# -*- coding: utf-8 -*-
import sqlite3
import calendar
from datetime import datetime, date, timedelta
from core.config import DB_NAME
from core.database import get_db_connection, get_payroll_config, get_payroll_cycle_dates
from core.utils import parse_time_str, format_time_ar


def calculate_employee_payroll(employee_id: int, year: int, month: int) -> dict:
    """
    حساب تفاصيل راتب الموظف لشهر محدد بناءً على إعدادات دورة الشهر وقواعد الغياب وحافز الانتظام.
    """
    cfg = get_payroll_config()
    start_date_str, end_date_str, cycle_desc = get_payroll_cycle_dates(year, month)
    daily_divisor = cfg.get("daily_divisor", 30.0)
    absence_rules = cfg.get("absence_rules", {})

    conn = get_db_connection()
    c = conn.cursor()

    # 1. Employee Info & Department Shift Schedule
    c.execute("""
        SELECT e.id, e.name, COALESCE(e.code, ''), COALESCE(e.base_salary, 0.0), 
               COALESCE(e.attendance_incentive, 500.0), COALESCE(e.daily_wage, 0.0),
               COALESCE(e.work_hours, d.default_work_hours, 11.0), 
               COALESCE(d.name, 'عام'), e.department_id,
               COALESCE(d.start_time, '08:00 ص'), COALESCE(d.end_time, '07:00 م'),
               COALESCE(d.default_work_hours, 11.0)
        FROM employees e
        LEFT JOIN departments d ON e.department_id = d.id
        WHERE e.id = ?
    """, (employee_id,))
    emp_row = c.fetchone()
    if not emp_row:
        conn.close()
        return None

    emp_id, emp_name, code, base_salary, max_incentive, daily_wage, work_hours, dept_name, dept_id, dept_start, dept_end, dept_hours = emp_row
    
    if daily_wage <= 0 and base_salary > 0:
        daily_wage = round(base_salary / daily_divisor, 2)

    # 2. Attendance & Absences in the active cycle range
    c.execute("""
        SELECT status, work_hours FROM attendance
        WHERE employee_id = ? AND date BETWEEN ? AND ?
    """, (employee_id, start_date_str, end_date_str))
    att_rows = c.fetchall()

    present_days = sum(1 for r in att_rows if r[0] == "حضور")
    raw_absent_days = sum(1 for r in att_rows if r[0] == "غياب")
    total_worked_hours = sum((r[1] or 0.0) for r in att_rows if r[0] == "حضور")

    # Fetch credited absence days (أيام احتساب مدفوعة على حساب المصنع)
    c.execute("""
        SELECT COALESCE(SUM(days), 0.0) FROM credited_leaves
        WHERE employee_id = ? AND date BETWEEN ? AND ?
    """, (employee_id, start_date_str, end_date_str))
    credited_days = float(c.fetchone()[0] or 0.0)

    # Calculate net effective absent days after subtracting credited days
    effective_absent_days = max(0.0, float(raw_absent_days) - credited_days)

    # Calculate total potential work days in cycle (excluding Fridays)
    start_dt = datetime.strptime(start_date_str, "%Y-%m-%d").date()
    end_dt = datetime.strptime(end_date_str, "%Y-%m-%d").date()
    total_work_days = 0
    curr_dt = start_dt
    while curr_dt <= end_dt:
        if curr_dt.weekday() != 4:  # 4 is Friday
            total_work_days += 1
        curr_dt += timedelta(days=1)

    # 3. Apply Configurable Absence & Incentive Rules (based on effective_absent_days)
    if effective_absent_days == 0:
        t0 = absence_rules.get("tier_0", {})
        incentive_deduction = float(t0.get("deduct_incentive", 0.0))
        absent_days_from_base = float(t0.get("deduct_base_days", 0.0))
    elif effective_absent_days <= 1:
        t1 = absence_rules.get("tier_1", {})
        incentive_deduction = float(t1.get("deduct_incentive", 100.0))
        absent_days_from_base = float(t1.get("deduct_base_days", 0.0))
    elif effective_absent_days <= 2:
        t2 = absence_rules.get("tier_2", {})
        incentive_deduction = float(t2.get("deduct_incentive", 200.0))
        absent_days_from_base = float(t2.get("deduct_base_days", 1.0))
    elif effective_absent_days <= 3:
        t3 = absence_rules.get("tier_3", {})
        incentive_deduction = float(t3.get("deduct_incentive", 300.0))
        absent_days_from_base = float(t3.get("deduct_base_days", 2.0))
    else:
        # 4+ days absent
        t4 = absence_rules.get("tier_4", {})
        incentive_deduction = float(t4.get("deduct_incentive", max_incentive))
        rule = t4.get("deduct_base_rule", "days_minus_1")
        if rule == "days_minus_1":
            absent_days_from_base = float(max(0, effective_absent_days - 1))
        else:
            absent_days_from_base = float(t4.get("deduct_base_fixed_days", effective_absent_days))

    # Clamp incentive deductions
    incentive_deduction = min(incentive_deduction, max_incentive)
    earned_incentive = max(0.0, max_incentive - incentive_deduction)
    absence_base_deduction = round(absent_days_from_base * daily_wage, 2)

    # 4. Overtime / Night Shifts (السهرات)
    ot_normal_mult = float(cfg.get("overtime_normal_multiplier", 1.5))
    ot_friday_mult = float(cfg.get("overtime_friday_multiplier", 2.0))

    c.execute("""
        SELECT date, start_time, end_time FROM overtime
        WHERE employee_id = ? AND date BETWEEN ? AND ?
    """, (employee_id, start_date_str, end_date_str))
    ot_rows = c.fetchall()

    shift_h = work_hours if work_hours > 0 else (dept_hours if dept_hours > 0 else 11.0)
    hourly_rate = (daily_wage / shift_h) if (daily_wage > 0 and shift_h > 0) else 0.0

    total_ot_hours = 0.0
    normal_ot_hours = 0.0
    friday_ot_hours = 0.0
    overtime_pay = 0.0

    for ot_date, s_time, e_time in ot_rows:
        t_start = parse_time_str(s_time)
        t_end = parse_time_str(e_time)
        if t_start.isValid() and t_end.isValid():
            dur_m = (t_end.hour() * 60 + t_end.minute()) - (t_start.hour() * 60 + t_start.minute())
            if dur_m < 0:
                dur_m += 24 * 60
            hrs = dur_m / 60.0
            total_ot_hours += hrs

            try:
                dt_obj = datetime.strptime(ot_date, "%Y-%m-%d")
                is_friday = (dt_obj.weekday() == 4)  # 4 is Friday
            except Exception:
                is_friday = False

            if is_friday:
                friday_ot_hours += hrs
                overtime_pay += hrs * hourly_rate * ot_friday_mult
            else:
                normal_ot_hours += hrs
                overtime_pay += hrs * hourly_rate * ot_normal_mult

    overtime_pay = round(overtime_pay, 2)

    # 5. General Deductions / Penalties
    c.execute("""
        SELECT COALESCE(SUM(amount), 0.0) FROM deductions
        WHERE employee_id = ? AND date BETWEEN ? AND ?
    """, (employee_id, start_date_str, end_date_str))
    penalties_total = c.fetchone()[0] or 0.0

    # 6. General Bonuses / Incentives
    c.execute("""
        SELECT COALESCE(SUM(amount), 0.0) FROM incentives
        WHERE employee_id = ? AND date BETWEEN ? AND ?
    """, (employee_id, start_date_str, end_date_str))
    extra_bonuses_total = c.fetchone()[0] or 0.0

    # 7. Monthly Advances (السلف الشهرية)
    c.execute("""
        SELECT COALESCE(SUM(amount), 0.0) FROM monthly_advances
        WHERE employee_id = ? AND date BETWEEN ? AND ?
    """, (employee_id, start_date_str, end_date_str))
    monthly_advances_total = c.fetchone()[0] or 0.0

    # 8. Long-Term Loan Installments (أقساط السلف طويلة الأجل)
    c.execute("""
        SELECT id, installment_amount, remaining_amount, start_date
        FROM long_term_loans
        WHERE employee_id = ? AND status = 'نشطة' AND remaining_amount > 0
    """, (employee_id,))
    loan_rows = c.fetchall()
    loan_installments_total = 0.0
    for lid, inst_amt, rem_amt, s_date in loan_rows:
        if s_date <= end_date_str or s_date[:7] <= end_date_str[:7]:
            due = min(inst_amt, rem_amt)
            loan_installments_total += due

    conn.close()

    # 9. Final Net Salary Calculations
    net_base = max(0.0, base_salary - absence_base_deduction)
    total_gross = net_base + earned_incentive + overtime_pay + extra_bonuses_total
    all_cash_deductions = penalties_total + monthly_advances_total + loan_installments_total
    total_deductions = absence_base_deduction + incentive_deduction + all_cash_deductions
    net_payout = max(0.0, total_gross - all_cash_deductions)

    return {
        "employee_id": emp_id,
        "name": emp_name,
        "code": code,
        "department_id": dept_id,
        "department_name": dept_name,
        "dept_start": dept_start,
        "dept_end": dept_end,
        "dept_hours": dept_hours,
        "base_salary": base_salary,
        "daily_wage": daily_wage,
        "hourly_rate": round(hourly_rate, 2),
        "total_work_days": total_work_days,
        "present_days": present_days,
        "raw_absent_days": raw_absent_days,
        "credited_days": credited_days,
        "absent_days": effective_absent_days,
        "effective_absent_days": effective_absent_days,
        "total_worked_hours": round(total_worked_hours, 1),
        "max_incentive": max_incentive,
        "earned_incentive": earned_incentive,
        "incentive_deduction": incentive_deduction,
        "absent_days_from_base": absent_days_from_base,
        "absence_base_deduction": absence_base_deduction,
        "overtime_hours": round(total_ot_hours, 1),
        "normal_ot_hours": round(normal_ot_hours, 1),
        "friday_ot_hours": round(friday_ot_hours, 1),
        "overtime_normal_mult": ot_normal_mult,
        "overtime_friday_mult": ot_friday_mult,
        "overtime_pay": overtime_pay,
        "penalties_total": round(penalties_total, 2),
        "extra_bonuses_total": round(extra_bonuses_total, 2),
        "monthly_advances_total": round(monthly_advances_total, 2),
        "loan_installments_total": round(loan_installments_total, 2),
        "total_gross": round(total_gross, 2),
        "total_deductions": round(total_deductions, 2),
        "net_payout": round(net_payout, 2),
        "start_date": start_date_str,
        "end_date": end_date_str,
        "cycle_desc": cycle_desc,
        "year": year,
        "month": month,
    }


def calculate_monthly_payroll_summary(year: int, month: int, department_id: int = None) -> dict:
    """
    حساب مسير الرواتب لجميع العاملين لشهر محدد، مع إمكانية التصفية بقسم معين،
    بناءً على التواريخ المحددة في دورة الشهر المالي.
    """
    start_date_str, end_date_str, cycle_desc = get_payroll_cycle_dates(year, month)

    conn = get_db_connection()
    c = conn.cursor()
    if department_id and department_id > 0:
        c.execute("SELECT id FROM employees WHERE department_id=? AND is_active=1 ORDER BY name ASC", (department_id,))
    else:
        c.execute("SELECT id FROM employees WHERE is_active=1 ORDER BY name ASC")
    emp_ids = [r[0] for r in c.fetchall()]
    conn.close()

    payroll_records = []
    tot_base_salary = 0.0
    tot_earned_incentive = 0.0
    tot_absence_base_deduction = 0.0
    tot_incentive_deduction = 0.0
    tot_overtime_pay = 0.0
    tot_penalties = 0.0
    tot_bonuses = 0.0
    tot_monthly_advances = 0.0
    tot_loan_installments = 0.0
    tot_net_payout = 0.0
    tot_raw_absent_days = 0.0
    tot_credited_days = 0.0
    tot_absent_days = 0.0
    tot_present_days = 0

    for eid in emp_ids:
        rec = calculate_employee_payroll(eid, year, month)
        if rec:
            payroll_records.append(rec)
            tot_base_salary += rec["base_salary"]
            tot_earned_incentive += rec["earned_incentive"]
            tot_absence_base_deduction += rec["absence_base_deduction"]
            tot_incentive_deduction += rec["incentive_deduction"]
            tot_overtime_pay += rec["overtime_pay"]
            tot_penalties += rec["penalties_total"]
            tot_bonuses += rec["extra_bonuses_total"]
            tot_monthly_advances += rec.get("monthly_advances_total", 0.0)
            tot_loan_installments += rec.get("loan_installments_total", 0.0)
            tot_net_payout += rec["net_payout"]
            tot_raw_absent_days += rec["raw_absent_days"]
            tot_credited_days += rec["credited_days"]
            tot_absent_days += rec["absent_days"]
            tot_present_days += rec["present_days"]

    return {
        "records": payroll_records,
        "count": len(payroll_records),
        "tot_base_salary": round(tot_base_salary, 2),
        "tot_earned_incentive": round(tot_earned_incentive, 2),
        "tot_absence_base_deduction": round(tot_absence_base_deduction, 2),
        "tot_incentive_deduction": round(tot_incentive_deduction, 2),
        "tot_overtime_pay": round(tot_overtime_pay, 2),
        "tot_penalties": round(tot_penalties, 2),
        "tot_bonuses": round(tot_bonuses, 2),
        "tot_monthly_advances": round(tot_monthly_advances, 2),
        "tot_loan_installments": round(tot_loan_installments, 2),
        "tot_net_payout": round(tot_net_payout, 2),
        "tot_raw_absent_days": tot_raw_absent_days,
        "tot_credited_days": tot_credited_days,
        "tot_absent_days": tot_absent_days,
        "tot_present_days": tot_present_days,
        "start_date": start_date_str,
        "end_date": end_date_str,
        "cycle_desc": cycle_desc,
        "year": year,
        "month": month,
    }
