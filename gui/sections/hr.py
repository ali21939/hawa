from core.config import *
from core.database import *
from core.utils import *
from core.payroll import calculate_employee_payroll, calculate_monthly_payroll_summary
from gui.widgets import *

from PySide6.QtCore import *
from PySide6.QtWidgets import *
from PySide6.QtGui import *
import sqlite3, os, time, csv, threading, calendar
from datetime import datetime, timedelta
import json
from openpyxl import Workbook
from openpyxl.styles import Font, Fill, Alignment, Border, Side, PatternFill


class HRMixin:
    def setup_hr_section(self, widget):
        self.hr_employee_combos = []
        self.hr_financial_views = []

        # Tabs Container
        self.hr_tabs = QTabWidget()
        self.hr_tabs.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.hr_tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                background: #ffffff;
                top: -1px;
            }
            QTabBar::tab {
                background: #f8fafc;
                color: #475569;
                padding: 9px 22px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-weight: bold;
                font-size: 13px;
                margin-right: 4px;
                border: 1px solid #e2e8f0;
                border-bottom: none;
            }
            QTabBar::tab:selected {
                background: #2563eb;
                color: #ffffff;
                border-color: #2563eb;
            }
            QTabBar::tab:hover:!selected {
                background: #e2e8f0;
                color: #0f172a;
            }
        """)
        
        # 1. Tab: Attendance & Absence
        tab1 = QWidget()
        t1_layout = QVBoxLayout(tab1)
        t1_layout.setSpacing(8)
        t1_layout.setContentsMargins(8, 8, 8, 8)

        # Control & Quick Action Bar
        ctrl_frame = QFrame()
        ctrl_frame.setStyleSheet("""
            QFrame {
                background: #f8fafc;
                border-radius: 8px;
                padding: 6px;
                border: 1.5px solid #e2e8f0;
            }
        """)
        ctrl_layout = QHBoxLayout(ctrl_frame)
        ctrl_layout.setSpacing(8)
        ctrl_layout.setContentsMargins(8, 6, 8, 6)

        # Date Label & Quick Nav
        date_label = QLabel("📅 التاريخ:")
        date_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        date_label.setStyleSheet("color: #1e293b; border: none; background: transparent;")
        ctrl_layout.addWidget(date_label)

        self.att_date = QDateEdit()
        self.att_date.setDate(QDate.currentDate())
        self.att_date.setDisplayFormat("yyyy-MM-dd")
        self.att_date.setCalendarPopup(True)
        self.att_date.setStyleSheet("""
            QDateEdit {
                font-size: 12px; font-weight: bold; padding: 5px 8px;
                border: 1.5px solid #cbd5e1; border-radius: 6px;
                background: white; color: #0f172a; min-width: 110px;
            }
            QDateEdit:focus { border: 2px solid #2563eb; }
        """)
        self.att_date.dateChanged.connect(self.load_attendance)
        ctrl_layout.addWidget(self.att_date)

        prev_day_btn = QPushButton("◀ السابق")
        prev_day_btn.setCursor(Qt.PointingHandCursor)
        prev_day_btn.setToolTip("الانتقال لليوم السابق (تخطي الجمعة)")
        prev_day_btn.setStyleSheet("""
            QPushButton { background: #e2e8f0; color: #334155; padding: 5px 10px; font-weight: bold; border-radius: 5px; border: none; font-size: 11px; }
            QPushButton:hover { background: #cbd5e1; }
        """)
        prev_day_btn.clicked.connect(self.go_previous_day)
        ctrl_layout.addWidget(prev_day_btn)

        today_btn = QPushButton("اليوم")
        today_btn.setCursor(Qt.PointingHandCursor)
        today_btn.setToolTip("العودة لتاريخ اليوم")
        today_btn.setStyleSheet("""
            QPushButton { background: #e0e7ff; color: #3730a3; padding: 5px 12px; font-weight: bold; border-radius: 5px; border: none; font-size: 11px; }
            QPushButton:hover { background: #c7d2fe; }
        """)
        today_btn.clicked.connect(lambda: self.att_date.setDate(QDate.currentDate()))
        ctrl_layout.addWidget(today_btn)

        next_day_btn = QPushButton("التالي ▶")
        next_day_btn.setCursor(Qt.PointingHandCursor)
        next_day_btn.setToolTip("الانتقال لليوم التالي (تخطي الجمعة)")
        next_day_btn.setStyleSheet("""
            QPushButton { background: #e2e8f0; color: #334155; padding: 5px 10px; font-weight: bold; border-radius: 5px; border: none; font-size: 11px; }
            QPushButton:hover { background: #cbd5e1; }
        """)
        next_day_btn.clicked.connect(self.go_next_day)
        ctrl_layout.addWidget(next_day_btn)

        # Search Bar
        search_lbl = QLabel("🔍 بحث:")
        search_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        search_lbl.setStyleSheet("color: #334155; border: none; background: transparent; margin-right: 4px;")
        ctrl_layout.addWidget(search_lbl)

        self.att_search_input = QLineEdit()
        self.att_search_input.setPlaceholderText("ابحث باسم العامل...")
        self.att_search_input.setClearButtonEnabled(True)
        self.att_search_input.setStyleSheet("""
            QLineEdit {
                font-size: 12px; padding: 5px 10px; border: 1.5px solid #cbd5e1;
                border-radius: 6px; background: white; color: #0f172a; min-width: 170px;
            }
            QLineEdit:focus { border: 2px solid #2563eb; }
        """)
        self.att_search_input.textChanged.connect(self.filter_attendance_table)
        ctrl_layout.addWidget(self.att_search_input)

        # Department Filter
        dept_filter_lbl = QLabel("🏢 القسم:")
        dept_filter_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        dept_filter_lbl.setStyleSheet("color: #334155; border: none; background: transparent; margin-right: 4px;")
        ctrl_layout.addWidget(dept_filter_lbl)

        self.att_dept_filter = QComboBox()
        self.att_dept_filter.setStyleSheet("""
            QComboBox {
                font-size: 12px; font-weight: bold; padding: 5px 10px; border: 1.5px solid #cbd5e1;
                border-radius: 6px; background: white; color: #0f172a; min-width: 150px;
            }
            QComboBox:focus { border: 2px solid #2563eb; }
        """)
        self.populate_dept_filter_combo(self.att_dept_filter)
        self.att_dept_filter.currentIndexChanged.connect(self.filter_attendance_table)
        ctrl_layout.addWidget(self.att_dept_filter)

        # Quick Bulk Presence Actions
        btn_mark_all_pres = QPushButton("✅ تحضير الكل")
        btn_mark_all_pres.setCursor(Qt.PointingHandCursor)
        btn_mark_all_pres.setToolTip("تحديد جميع العمال المعروضين كـ 'حضور'")
        btn_mark_all_pres.setStyleSheet("""
            QPushButton {
                background-color: #dcfce7; color: #15803d; border: 1px solid #86efac;
                padding: 6px 12px; font-weight: bold; border-radius: 6px; font-size: 11px;
            }
            QPushButton:hover { background-color: #bbf7d0; }
        """)
        btn_mark_all_pres.clicked.connect(self.mark_all_present)
        ctrl_layout.addWidget(btn_mark_all_pres)

        btn_mark_all_abs = QPushButton("❌ تغييب الكل")
        btn_mark_all_abs.setCursor(Qt.PointingHandCursor)
        btn_mark_all_abs.setToolTip("تحديد جميع العمال المعروضين كـ 'غياب'")
        btn_mark_all_abs.setStyleSheet("""
            QPushButton {
                background-color: #fee2e2; color: #b91c1c; border: 1px solid #fca5a5;
                padding: 6px 12px; font-weight: bold; border-radius: 6px; font-size: 11px;
            }
            QPushButton:hover { background-color: #fecaca; }
        """)
        btn_mark_all_abs.clicked.connect(self.mark_all_absent)
        ctrl_layout.addWidget(btn_mark_all_abs)

        ctrl_layout.addStretch()

        # Save Attendance Button
        self.save_att_btn = QPushButton("💾 حفظ كشوف اليوم")
        self.save_att_btn.setCursor(Qt.PointingHandCursor)
        self.save_att_btn.setStyleSheet("""
            QPushButton {
                background-color: #16a34a; color: white; padding: 7px 18px;
                font-weight: bold; border-radius: 6px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #15803d; }
            QPushButton:pressed { background-color: #166534; }
        """)
        self.save_att_btn.clicked.connect(self.save_attendance)
        ctrl_layout.addWidget(self.save_att_btn)

        # Manage Employees Button
        manage_employees_btn = QPushButton("⚙️ أسماء العاملين")
        manage_employees_btn.setCursor(Qt.PointingHandCursor)
        manage_employees_btn.setStyleSheet("""
            QPushButton {
                background-color: #0f172a; color: white; padding: 7px 16px;
                font-weight: bold; border-radius: 6px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #1e293b; }
        """)
        manage_employees_btn.clicked.connect(self.open_employees_manager)
        ctrl_layout.addWidget(manage_employees_btn)

        t1_layout.addWidget(ctrl_frame)

        # Summary Metrics Cards
        self.att_metrics_frame = QFrame()
        self.att_metrics_frame.setStyleSheet("border: none; background: transparent;")
        metrics_layout = QHBoxLayout(self.att_metrics_frame)
        metrics_layout.setContentsMargins(0, 0, 0, 0)
        metrics_layout.setSpacing(10)

        self.card_total_workers = self.create_hr_metric_card("👥 إجمالي العمال", "0", "#475569", "#f1f5f9")
        self.card_present_workers = self.create_hr_metric_card("✅ الحاضرين", "0", "#15803d", "#f0fdf4")
        self.card_absent_workers = self.create_hr_metric_card("❌ الغائبين", "0", "#b91c1c", "#fef2f2")
        self.card_attendance_pct = self.create_hr_metric_card("📊 نسبة الحضور", "0%", "#2563eb", "#eff6ff")
        self.card_total_work_hours = self.create_hr_metric_card("⏱️ إجمالي ساعات العمل", "0.0 س", "#7c3aed", "#f5f3ff")

        metrics_layout.addWidget(self.card_total_workers)
        metrics_layout.addWidget(self.card_present_workers)
        metrics_layout.addWidget(self.card_absent_workers)
        metrics_layout.addWidget(self.card_attendance_pct)
        metrics_layout.addWidget(self.card_total_work_hours)
        t1_layout.addWidget(self.att_metrics_frame)

        # Absent card interactive popup
        self.card_absent_workers.setCursor(Qt.PointingHandCursor)
        self.card_absent_workers.setToolTip("اضغط لعرض قائمة وأسماء الغائبين اليوم")
        self.card_absent_workers.mousePressEvent = lambda e: self.show_todays_absentees()

        # Attendance Table
        self.att_table = QTableWidget()
        self.att_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.att_table.setColumnCount(8)
        self.att_table.setHorizontalHeaderLabels([
            "#", "اسم العامل والوردية", "حالة الحضور", "🕒 وقت الحضور (البداية)",
            "🕑 وقت الانصراف (النهاية)", "⏱️ ساعات العمل", "ملاحظات وتنبيهات", "📅 أيام الغياب"
        ])
        self.att_table.setAlternatingRowColors(True)
        self.att_table.setStyleSheet("""
            QTableWidget {
                border: 1.5px solid #cbd5e1;
                border-radius: 8px;
                background: #ffffff;
                gridline-color: #e2e8f0;
                alternate-background-color: #f8fafc;
                selection-background-color: #dbeafe;
                selection-color: #1e3a8a;
            }
            QTableWidget::item {
                padding: 6px 10px;
                border-bottom: 1px solid #e2e8f0;
                color: #0f172a;
                font-size: 13px;
            }
            QTableWidget::item:selected {
                background: #dbeafe;
                color: #1e3a8a;
                font-weight: bold;
            }
            QHeaderView::section {
                background-color: #0f172a;
                color: #ffffff;
                padding: 10px 8px;
                font-weight: bold;
                font-size: 13px;
                border: 1px solid #1e293b;
                border-bottom: 3px solid #2563eb;
            }
        """)
        header = self.att_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.Stretch)
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        self.att_table.verticalHeader().setVisible(False)
        self.att_table.verticalHeader().setDefaultSectionSize(44)
        t1_layout.addWidget(self.att_table)

        self.att_summary = QLabel("")
        self.att_summary.setVisible(False)
        t1_layout.addWidget(self.att_summary)

        self.hr_tabs.addTab(tab1, "📋 الحضور والغياب")

        # 2. Tab: Deductions
        tab2 = QWidget()
        self.setup_hr_financial_tab(tab2, "deductions", "الخصومات", "➖ الخصومات")
        self.hr_tabs.addTab(tab2, "➖ الخصومات")

        # 3. Tab: Incentives
        tab3 = QWidget()
        self.setup_hr_financial_tab(tab3, "incentives", "الحوافز", "➕ الحوافز")
        self.hr_tabs.addTab(tab3, "➕ الحوافز")

        # 4. Tab: Credited Leaves (أيام الاحتساب على المصنع)
        tab_cl = QWidget()
        self.setup_credited_leaves_tab(tab_cl)
        self.hr_tabs.addTab(tab_cl, "✨ أيام الاحتساب (على المصنع)")

        # 5. Tab: Attendance Summary
        tab4 = QWidget()
        self.setup_attendance_summary_tab(tab4)
        self.hr_tabs.addTab(tab4, "📊 تجميع الحضور والتقارير")

        # 5. Tab: Overtime (السهرات)
        tab5 = QWidget()
        self.setup_overtime_tab(tab5)
        self.hr_tabs.addTab(tab5, "🌙 السهرات")

        # 6. Tab: Departments Management (إدارة الأقسام)
        tab6 = QWidget()
        self.setup_department_management_tab(tab6)
        self.hr_tabs.addTab(tab6, "🏢 إدارة الأقسام")

        # 7. Tab: Payroll & Salary Sheet (مسير الرواتب والأجور)
        tab7 = QWidget()
        self.setup_payroll_tab(tab7)
        self.hr_tabs.addTab(tab7, "💵 مسير الرواتب والأجور")

        # 8. Tab: Settings & Rules (الإعدادات وقواعد الرواتب)
        tab8 = QWidget()
        self.setup_payroll_settings_tab(tab8)
        self.hr_tabs.addTab(tab8, "⚙️ الإعدادات وقواعد الرواتب")

        widget.layout.addWidget(self.hr_tabs)
        self.load_attendance()

    def is_friday(self, qdate: QDate) -> bool:
        """التحقق مما إذا كان التاريخ المحدد يوافق يوم جمعة."""
        return qdate.dayOfWeek() == 5

    def previous_work_date(self, qdate: QDate) -> str:
        """حساب يوم العمل السابق مع تخطي يوم الجمعة."""
        d = qdate.addDays(-1)
        if d.dayOfWeek() == 5:
            d = d.addDays(-1)
        return d.toString("yyyy-MM-dd")

    def go_previous_day(self):
        """الانتقال إلى يوم العمل السابق."""
        cur = self.att_date.date()
        prev_d = cur.addDays(-1)
        if prev_d.dayOfWeek() == 5:
            prev_d = prev_d.addDays(-1)
        self.att_date.setDate(prev_d)

    def go_next_day(self):
        """الانتقال إلى يوم العمل التالي."""
        cur = self.att_date.date()
        next_d = cur.addDays(1)
        if next_d.dayOfWeek() == 5:
            next_d = next_d.addDays(1)
        self.att_date.setDate(next_d)

    def mark_all_present(self):
        """تسجيل جميع العمال المعروضين كـ 'حضور' بمواعيد وراديهم الافتراضية."""
        for r in range(self.att_table.rowCount()):
            if not self.att_table.isRowHidden(r):
                combo = self.att_table.cellWidget(r, 2)
                if combo:
                    combo.setCurrentText("حضور")
        self.recalculate_attendance_metrics_from_table()
        if hasattr(self, "show_toast"):
            self.show_toast("✓ تم تحديد جميع العمال المعروضين كـ 'حضور'")

    def mark_all_absent(self):
        """تسجيل جميع العمال المعروضين كـ 'غياب'."""
        for r in range(self.att_table.rowCount()):
            if not self.att_table.isRowHidden(r):
                combo = self.att_table.cellWidget(r, 2)
                if combo:
                    combo.setCurrentText("غياب")
        self.recalculate_attendance_metrics_from_table()
        if hasattr(self, "show_toast"):
            self.show_toast("✓ تم تحديد جميع العمال المعروضين كـ 'غياب'")

    def create_hr_metric_card(self, title, initial_value, text_color, bg_color):
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {bg_color};
                border: 1.5px solid #cbd5e1;
                border-radius: 8px;
                padding: 4px 10px;
            }}
        """)
        layout = QHBoxLayout(card)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(8)

        lbl_title = QLabel(title)
        lbl_title.setFont(QFont("Segoe UI", 9, QFont.Bold))
        lbl_title.setStyleSheet("color: #475569; border: none; background: transparent;")

        lbl_val = QLabel(initial_value)
        lbl_val.setFont(QFont("Segoe UI", 12, QFont.Bold))
        lbl_val.setStyleSheet(f"color: {text_color}; border: none; background: transparent;")

        layout.addWidget(lbl_title)
        layout.addWidget(lbl_val)
        card.val_label = lbl_val
        return card

    def setup_hr_financial_tab(self, tab, table_name, title, tab_label):
        layout = QVBoxLayout(tab)
        layout.setSpacing(12)
        layout.setContentsMargins(14, 14, 14, 14)

        # Form Card
        form_frame = QFrame()
        form_frame.setStyleSheet("background: #f8fafc; border-radius: 8px; padding: 12px; border: 1.5px solid #e2e8f0;")
        form_layout = QHBoxLayout(form_frame)
        form_layout.setSpacing(12)
        form_layout.setContentsMargins(10, 8, 10, 8)

        # Emp Combo
        emp_label = QLabel("اسم العامل:")
        emp_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        emp_label.setStyleSheet("color: #334155; border: none; background: transparent;")
        form_layout.addWidget(emp_label)

        emp_combo = QComboBox()
        emp_combo.setStyleSheet("font-size: 12px; padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; min-width: 180px; color: #0f172a;")
        self.populate_employee_combo(emp_combo)
        self.hr_employee_combos.append(emp_combo)
        form_layout.addWidget(emp_combo)

        # Amount Input
        amt_label = QLabel("المبلغ (ج.م):")
        amt_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        amt_label.setStyleSheet("color: #334155; border: none; background: transparent;")
        form_layout.addWidget(amt_label)

        amount_input = QDoubleSpinBox()
        amount_input.setRange(0, 100000)
        amount_input.setDecimals(2)
        amount_input.setStyleSheet("font-size: 12px; font-weight: bold; padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; min-width: 110px; color: #0f172a;")
        form_layout.addWidget(amount_input)

        # Reason Input
        reason_label = QLabel("السبب / التفاصيل:")
        reason_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        reason_label.setStyleSheet("color: #334155; border: none; background: transparent;")
        form_layout.addWidget(reason_label)

        reason_input = QLineEdit()
        reason_input.setPlaceholderText(f"أدخل سبب تسجيل {title}...")
        reason_input.setStyleSheet("font-size: 12px; padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;")
        form_layout.addWidget(reason_input)

        # Add Button
        add_btn = QPushButton(f"➕ إضافة {title}")
        add_btn.setCursor(Qt.PointingHandCursor)
        btn_color = "#2563eb" if table_name == "incentives" else "#dc2626"
        add_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {btn_color};
                color: white;
                padding: 8px 18px;
                font-weight: bold;
                border-radius: 6px;
                border: none;
                font-size: 12px;
            }}
            QPushButton:hover {{ opacity: 0.9; }}
        """)
        add_btn.clicked.connect(lambda: self.save_financial(table_name, emp_combo.currentData(), amount_input.value(), reason_input.text(), table_widget, sum_label, amount_input, reason_input))
        form_layout.addWidget(add_btn)

        layout.addWidget(form_frame)

        # Table
        table_widget = QTableWidget()
        table_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.prepare_table_with_delete(table_widget, ["ID", "التاريخ", "اسم العامل", "المبلغ (ج.م)", "السبب والتفاصيل"])

        header = table_widget.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        table_widget.verticalHeader().setDefaultSectionSize(40)
        table_widget.verticalHeader().setVisible(False)

        layout.addWidget(table_widget)

        # Footer summary label
        sum_label = QLabel(f"إجمالي {title}: 0.00 ج.م")
        sum_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
        sum_label.setStyleSheet("color: #1e293b; padding: 8px 14px; background: #f1f5f9; border-radius: 6px; border: 1.5px solid #cbd5e1;")
        layout.addWidget(sum_label)

        self.hr_financial_views.append({"table": table_name, "table_widget": table_widget, "sum_label": sum_label})
        self.load_financial(table_name, table_widget, sum_label)

    def save_financial(self, table_name, emp_id, amount, reason, table_widget, sum_label, amt_input, reason_input):
        if not emp_id:
            QMessageBox.warning(self, "تنبيه", "يرجى اختيار العامل أولاً.")
            return
        if amount <= 0:
            QMessageBox.warning(self, "تنبيه", "يرجى إدخال مبلغ صحيح أكبر من الصفر.")
            return

        date_str = self.att_date.date().toString("yyyy-MM-dd") if hasattr(self, "att_date") else datetime.now().strftime("%Y-%m-%d")
        conn = get_db_connection()
        c = conn.cursor()
        c.execute(f"""
            INSERT INTO {table_name} (date, employee_id, amount, reason)
            VALUES (?, ?, ?, ?)
        """, (date_str, emp_id, amount, reason.strip()))
        conn.commit()
        conn.close()

        amt_input.setValue(0)
        reason_input.clear()
        self.load_financial(table_name, table_widget, sum_label)
        if hasattr(self, "load_payroll_table"):
            self.load_payroll_table()
        self.refresh_dashboard_data()
        item_title = "الخصم" if table_name == "deductions" else "الحافز"
        if hasattr(self, "show_toast"):
            self.show_toast(f"✓ تم تسجيل {item_title} بمبلغ {amount:,.2f} ج.م بنجاح")

    def load_financial(self, table_name, table_widget, sum_label):
        if not table_widget:
            return
        conn = get_db_connection()
        c = conn.cursor()
        c.execute(f"""
            SELECT f.id, f.date, e.name, f.amount, COALESCE(f.reason, '')
            FROM {table_name} f
            JOIN employees e ON f.employee_id = e.id
            ORDER BY f.date DESC, f.id DESC
        """)
        rows = c.fetchall()
        conn.close()

        with self.preserve_table_scroll(table_widget):
            table_widget.setUpdatesEnabled(False)
            table_widget.setRowCount(len(rows))
            total_sum = 0.0

            for r_idx, (rec_id, date_val, emp_name, amt, reason_text) in enumerate(rows):
                total_sum += (amt or 0.0)

                # Col 0: ID
                id_item = QTableWidgetItem(str(rec_id))
                id_item.setTextAlignment(Qt.AlignCenter)
                table_widget.setItem(r_idx, 0, id_item)

                # Col 1: Date
                d_item = QTableWidgetItem(str(date_val))
                d_item.setTextAlignment(Qt.AlignCenter)
                table_widget.setItem(r_idx, 1, d_item)

                # Col 2: Name
                n_item = QTableWidgetItem(emp_name)
                n_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                table_widget.setItem(r_idx, 2, n_item)

                # Col 3: Amount
                amt_item = QTableWidgetItem(f"{amt:,.2f} ج.م")
                amt_item.setTextAlignment(Qt.AlignCenter)
                amt_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                color = "#dc2626" if table_name == "deductions" else "#16a34a"
                amt_item.setForeground(QColor(color))
                table_widget.setItem(r_idx, 3, amt_item)

                # Col 4: Reason
                table_widget.setItem(r_idx, 4, QTableWidgetItem(reason_text))

                # Col 5: Edit Button
                self.add_edit_button(table_widget, r_idx, 5, table_name, rec_id, lambda tn=table_name, tw=table_widget, sl=sum_label: self.load_financial(tn, tw, sl))

                # Col 6: Delete Button
                self.add_delete_button(table_widget, r_idx, 6, table_name, rec_id, lambda tn=table_name, tw=table_widget, sl=sum_label: self.load_financial(tn, tw, sl))

            table_widget.setUpdatesEnabled(True)

        title = "الخصومات" if table_name == "deductions" else "الحوافز"
        if sum_label:
            sum_label.setText(f"إجمالي {title}: {total_sum:,.2f} ج.م  |  عدد العمليات: {len(rows)}")

    def edit_financial_record_dialog(self, table_name, record_id, refresh_callback):
        """نافذة تعديل سجل الخصم أو الحافز."""
        conn = get_db_connection()
        c = conn.cursor()
        c.execute(f"SELECT date, employee_id, amount, reason FROM {table_name} WHERE id=?", (record_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            return

        cur_date, emp_id, cur_amt, cur_reason = row
        dlg = QDialog(self)
        item_title = "خصم" if table_name == "deductions" else "حافز"
        dlg.setWindowTitle(f"تعديل سجل {item_title} #{record_id}")
        dlg.setLayoutDirection(Qt.RightToLeft)
        dlg.resize(420, 320)
        dlg_layout = QVBoxLayout(dlg)
        dlg_layout.setSpacing(12)

        form = QFormLayout()
        form.setSpacing(10)

        date_edit = QDateEdit()
        date_edit.setDate(QDate.fromString(cur_date, "yyyy-MM-dd"))
        date_edit.setCalendarPopup(True)
        form.addRow("التاريخ:", date_edit)

        emp_combo = QComboBox()
        self.populate_employee_combo(emp_combo, emp_id)
        form.addRow("اسم العامل:", emp_combo)

        amt_spin = QDoubleSpinBox()
        amt_spin.setRange(0.0, 100000.0)
        amt_spin.setValue(float(cur_amt or 0))
        amt_spin.setSuffix(" ج.م")
        form.addRow("المبلغ:", amt_spin)

        reason_input = QLineEdit(cur_reason or "")
        form.addRow("السبب:", reason_input)

        dlg_layout.addLayout(form)

        btn_box = QHBoxLayout()
        save_btn = QPushButton("💾 حفظ التعديل")
        save_btn.setStyleSheet("background: #16a34a; color: white; padding: 8px 18px; font-weight: bold; border-radius: 6px;")
        cancel_btn = QPushButton("إلغاء")
        cancel_btn.setStyleSheet("background: #64748b; color: white; padding: 8px 18px; border-radius: 6px;")
        btn_box.addWidget(save_btn)
        btn_box.addWidget(cancel_btn)
        dlg_layout.addLayout(btn_box)

        cancel_btn.clicked.connect(dlg.reject)

        def save_changes():
            new_date = date_edit.date().toString("yyyy-MM-dd")
            new_eid = emp_combo.currentData()
            new_amt = amt_spin.value()
            new_r = reason_input.text().strip()

            con = get_db_connection()
            cr = con.cursor()
            cr.execute(f"UPDATE {table_name} SET date=?, employee_id=?, amount=?, reason=? WHERE id=?",
                       (new_date, new_eid, new_amt, new_r, record_id))
            con.commit()
            con.close()

            dlg.accept()
            refresh_callback()
            if hasattr(self, "load_payroll_table"):
                self.load_payroll_table()
            self.refresh_dashboard_data()
            if hasattr(self, "show_toast"):
                self.show_toast(f"✓ تم تعديل سجل {item_title} بنجاح")

        save_btn.clicked.connect(save_changes)
        dlg.exec()

    def setup_credited_leaves_tab(self, tab):
        """إعداد تبويب إدارة أيام الاحتساب (الإجازات المدفوعة على حساب المصنع)."""
        layout = QVBoxLayout(tab)
        layout.setSpacing(10)
        layout.setContentsMargins(12, 12, 12, 12)

        # ── 1. نموذج الإدخال السريع (Entry Form Card) ────────────────────────
        form_frame = QFrame()
        form_frame.setStyleSheet("""
            QFrame {
                background: #f8fafc;
                border-radius: 8px;
                padding: 10px;
                border: 1.5px solid #e2e8f0;
            }
        """)
        form_layout = QHBoxLayout(form_frame)
        form_layout.setSpacing(10)
        form_layout.setContentsMargins(8, 6, 8, 6)

        # Emp Combo
        emp_lbl = QLabel("👤 العامل:")
        emp_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        emp_lbl.setStyleSheet("color: #334155; border: none; background: transparent;")
        form_layout.addWidget(emp_lbl)

        self.cl_emp_combo = QComboBox()
        self.cl_emp_combo.setStyleSheet("""
            QComboBox {
                font-size: 12px; font-weight: bold; padding: 5px 8px;
                border: 1.5px solid #cbd5e1; border-radius: 6px;
                background: white; min-width: 180px; color: #0f172a;
            }
            QComboBox:focus { border: 2px solid #7c3aed; }
        """)
        self.populate_employee_combo(self.cl_emp_combo)
        self.hr_employee_combos.append(self.cl_emp_combo)
        form_layout.addWidget(self.cl_emp_combo)

        # Days Input
        days_lbl = QLabel("📅 الأيام المحتسبة:")
        days_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        days_lbl.setStyleSheet("color: #334155; border: none; background: transparent;")
        form_layout.addWidget(days_lbl)

        self.cl_days_spin = QDoubleSpinBox()
        self.cl_days_spin.setRange(0.5, 30.0)
        self.cl_days_spin.setSingleStep(0.5)
        self.cl_days_spin.setValue(1.0)
        self.cl_days_spin.setSuffix(" يوم")
        self.cl_days_spin.setStyleSheet("""
            QDoubleSpinBox {
                font-size: 12px; font-weight: bold; padding: 5px 8px;
                border: 1.5px solid #cbd5e1; border-radius: 6px;
                background: white; min-width: 95px; color: #0f172a;
            }
            QDoubleSpinBox:focus { border: 2px solid #7c3aed; }
        """)
        form_layout.addWidget(self.cl_days_spin)

        # Date Input
        date_lbl = QLabel("📆 التاريخ:")
        date_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        date_lbl.setStyleSheet("color: #334155; border: none; background: transparent;")
        form_layout.addWidget(date_lbl)

        self.cl_date_edit = QDateEdit()
        self.cl_date_edit.setDate(QDate.currentDate())
        self.cl_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.cl_date_edit.setCalendarPopup(True)
        self.cl_date_edit.setStyleSheet("""
            QDateEdit {
                font-size: 12px; font-weight: bold; padding: 5px 8px;
                border: 1.5px solid #cbd5e1; border-radius: 6px;
                background: white; min-width: 110px; color: #0f172a;
            }
            QDateEdit:focus { border: 2px solid #7c3aed; }
        """)
        form_layout.addWidget(self.cl_date_edit)

        # Reason Input
        reason_lbl = QLabel("السبب:")
        reason_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        reason_lbl.setStyleSheet("color: #334155; border: none; background: transparent;")
        form_layout.addWidget(reason_lbl)

        self.cl_reason_input = QLineEdit()
        self.cl_reason_input.setPlaceholderText("مثال: إذن إدارة / احتساب غياب / ظرف طارئ...")
        self.cl_reason_input.setStyleSheet("""
            QLineEdit {
                font-size: 12px; padding: 5px 8px;
                border: 1.5px solid #cbd5e1; border-radius: 6px;
                background: white; color: #0f172a; min-width: 160px;
            }
            QLineEdit:focus { border: 2px solid #7c3aed; }
        """)
        form_layout.addWidget(self.cl_reason_input)

        # Notes Input
        notes_lbl = QLabel("ملاحظات:")
        notes_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        notes_lbl.setStyleSheet("color: #334155; border: none; background: transparent;")
        form_layout.addWidget(notes_lbl)

        self.cl_notes_input = QLineEdit()
        self.cl_notes_input.setPlaceholderText("ملاحظات إضافية...")
        self.cl_notes_input.setStyleSheet("""
            QLineEdit {
                font-size: 12px; padding: 5px 8px;
                border: 1.5px solid #cbd5e1; border-radius: 6px;
                background: white; color: #0f172a; min-width: 130px;
            }
            QLineEdit:focus { border: 2px solid #7c3aed; }
        """)
        form_layout.addWidget(self.cl_notes_input)

        # Add Button
        add_cl_btn = QPushButton("➕ تسجيل احتساب")
        add_cl_btn.setCursor(Qt.PointingHandCursor)
        add_cl_btn.setStyleSheet("""
            QPushButton {
                background-color: #7c3aed; color: white; padding: 7px 18px;
                font-weight: bold; border-radius: 6px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #6d28d9; }
            QPushButton:pressed { background-color: #5b21b6; }
        """)
        add_cl_btn.clicked.connect(self.save_credited_leave_handler)
        form_layout.addWidget(add_cl_btn)

        layout.addWidget(form_frame)

        # ── 2. شريط الفلترة والبحث (Filter Bar) ─────────────────────────────
        filter_frame = QFrame()
        filter_frame.setStyleSheet("""
            QFrame {
                background: #f1f5f9; border-radius: 6px; padding: 6px; border: 1px solid #cbd5e1;
            }
        """)
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setSpacing(8)
        filter_layout.setContentsMargins(6, 4, 6, 4)

        # Search Query
        f_search_lbl = QLabel("🔍 بحث:")
        f_search_lbl.setFont(QFont("Segoe UI", 9, QFont.Bold))
        f_search_lbl.setStyleSheet("color: #334155; border: none; background: transparent;")
        filter_layout.addWidget(f_search_lbl)

        self.cl_search_input = QLineEdit()
        self.cl_search_input.setPlaceholderText("ابحث باسم العامل أو الكود أو السبب...")
        self.cl_search_input.setClearButtonEnabled(True)
        self.cl_search_input.setStyleSheet("""
            QLineEdit {
                font-size: 11px; padding: 4px 8px; border: 1px solid #cbd5e1;
                border-radius: 4px; background: white; min-width: 170px;
            }
        """)
        self.cl_search_input.textChanged.connect(self.filter_credited_leaves_table)
        filter_layout.addWidget(self.cl_search_input)

        # Dept Filter
        f_dept_lbl = QLabel("🏢 القسم:")
        f_dept_lbl.setFont(QFont("Segoe UI", 9, QFont.Bold))
        f_dept_lbl.setStyleSheet("color: #334155; border: none; background: transparent;")
        filter_layout.addWidget(f_dept_lbl)

        self.cl_dept_filter = QComboBox()
        self.cl_dept_filter.setStyleSheet("""
            QComboBox {
                font-size: 11px; font-weight: bold; padding: 4px 6px;
                border: 1px solid #cbd5e1; border-radius: 4px; background: white; min-width: 140px;
            }
        """)
        self.populate_dept_filter_combo(self.cl_dept_filter)
        self.cl_dept_filter.currentIndexChanged.connect(self.load_credited_leaves_table)
        filter_layout.addWidget(self.cl_dept_filter)

        filter_layout.addStretch()

        refresh_btn = QPushButton("🔄 تحديث")
        refresh_btn.setCursor(Qt.PointingHandCursor)
        refresh_btn.setStyleSheet("""
            QPushButton {
                background: #e2e8f0; color: #334155; padding: 4px 12px;
                font-weight: bold; border-radius: 4px; border: none; font-size: 11px;
            }
            QPushButton:hover { background: #cbd5e1; }
        """)
        refresh_btn.clicked.connect(self.load_credited_leaves_table)
        filter_layout.addWidget(refresh_btn)

        layout.addWidget(filter_frame)

        # ── 3. جدول أيام الاحتساب (Table Widget) ───────────────────────────
        self.credited_leaves_table = QTableWidget()
        self.credited_leaves_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.prepare_table_with_delete(self.credited_leaves_table, [
            "ID", "التاريخ", "كود العامل", "اسم العامل", "القسم", "الأيام المحتسبة", "السبب والتفاصيل", "ملاحظات"
        ])

        hdr = self.credited_leaves_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.Stretch)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(6, QHeaderView.Stretch)
        hdr.setSectionResizeMode(7, QHeaderView.Stretch)
        self.credited_leaves_table.verticalHeader().setDefaultSectionSize(40)
        self.credited_leaves_table.verticalHeader().setVisible(False)
        layout.addWidget(self.credited_leaves_table)

        # ── 4. شريط الإحصائيات (Summary Ribbon) ────────────────────────────
        self.cl_summary_label = QLabel("إجمالي أيام الاحتساب: 0.0 يوم | عدد العمليات: 0")
        self.cl_summary_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self.cl_summary_label.setStyleSheet("""
            QLabel {
                color: #5b21b6; padding: 8px 14px; background: #f5f3ff;
                border-radius: 6px; border: 1.5px solid #c4b5fd;
            }
        """)
        layout.addWidget(self.cl_summary_label)

        self.load_credited_leaves_table()

    def save_credited_leave_handler(self):
        """تسجيل أيام احتساب جديدة من نموذج الإدخال السريع."""
        emp_id = self.cl_emp_combo.currentData()
        if not emp_id:
            QMessageBox.warning(self, "تنبيه", "يرجى اختيار العامل أولاً.")
            return

        days = self.cl_days_spin.value()
        if days <= 0:
            QMessageBox.warning(self, "تنبيه", "يرجى إدخال عدد أيام احتساب أكبر من الصفر.")
            return

        date_str = self.cl_date_edit.date().toString("yyyy-MM-dd")
        reason = self.cl_reason_input.text().strip()
        notes = self.cl_notes_input.text().strip()

        res = add_credited_leave(emp_id, days, date_str, reason, notes)
        if res.get("success"):
            self.cl_days_spin.setValue(1.0)
            self.cl_reason_input.clear()
            self.cl_notes_input.clear()
            self.load_credited_leaves_table()
            if hasattr(self, "load_payroll_table"):
                self.load_payroll_table()
            self.refresh_dashboard_data()
            if hasattr(self, "show_toast"):
                self.show_toast(f"✓ تم تسجيل {days} يوم احتساب على المصنع بنجاح")
        else:
            QMessageBox.critical(self, "خطأ", f"تعذر تسجيل أيام الاحتساب: {res.get('error')}")

    def load_credited_leaves_table(self):
        """تحميل وعرض سجلات أيام الاحتساب في الجدول."""
        if not hasattr(self, "credited_leaves_table") or self.credited_leaves_table is None:
            return

        dept_id = self.cl_dept_filter.currentData() if hasattr(self, "cl_dept_filter") else None
        rows = get_credited_leaves(department_id=dept_id)

        with self.preserve_table_scroll(self.credited_leaves_table):
            self.credited_leaves_table.setUpdatesEnabled(False)
            self.credited_leaves_table.setRowCount(len(rows))
            total_days = 0.0

            for r_idx, (rec_id, date_val, emp_id, emp_name, emp_code, dept_name, days_cnt, reason_txt, notes_txt, created_at, dept_id_val) in enumerate(rows):
                total_days += (days_cnt or 0.0)

                # Col 0: ID
                id_item = QTableWidgetItem(str(rec_id))
                id_item.setTextAlignment(Qt.AlignCenter)
                self.credited_leaves_table.setItem(r_idx, 0, id_item)

                # Col 1: Date
                d_item = QTableWidgetItem(str(date_val))
                d_item.setTextAlignment(Qt.AlignCenter)
                self.credited_leaves_table.setItem(r_idx, 1, d_item)

                # Col 2: Code
                c_item = QTableWidgetItem(str(emp_code or "-"))
                c_item.setTextAlignment(Qt.AlignCenter)
                c_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                self.credited_leaves_table.setItem(r_idx, 2, c_item)

                # Col 3: Employee Name
                n_item = QTableWidgetItem(emp_name)
                n_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                n_item.setData(Qt.UserRole, emp_id)
                self.credited_leaves_table.setItem(r_idx, 3, n_item)

                # Col 4: Dept Name
                dept_item = QTableWidgetItem(dept_name)
                dept_item.setTextAlignment(Qt.AlignCenter)
                dept_item.setForeground(QColor("#0369a1"))
                self.credited_leaves_table.setItem(r_idx, 4, dept_item)

                # Col 5: Credited Days Badge
                days_item = QTableWidgetItem(f"{days_cnt:g} يوم")
                days_item.setTextAlignment(Qt.AlignCenter)
                days_item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                days_item.setForeground(QColor("#7c3aed"))
                days_item.setBackground(QColor("#f5f3ff"))
                self.credited_leaves_table.setItem(r_idx, 5, days_item)

                # Col 6: Reason
                self.credited_leaves_table.setItem(r_idx, 6, QTableWidgetItem(reason_txt or ""))

                # Col 7: Notes
                self.credited_leaves_table.setItem(r_idx, 7, QTableWidgetItem(notes_txt or ""))

                # Col 8: Edit Button
                self.add_edit_button(self.credited_leaves_table, r_idx, 8, "credited_leaves", rec_id, self.load_credited_leaves_table)

                # Col 9: Delete Button
                self.add_delete_button(self.credited_leaves_table, r_idx, 9, "credited_leaves", rec_id, self.load_credited_leaves_table)

            self.credited_leaves_table.setUpdatesEnabled(True)

        if hasattr(self, "cl_summary_label") and self.cl_summary_label:
            self.cl_summary_label.setText(f"🎁 إجمالي أيام الاحتساب (على المصنع): {total_days:g} يوم  |  عدد العمليات: {len(rows)}")

        if hasattr(self, "cl_search_input") and self.cl_search_input.text():
            self.filter_credited_leaves_table(self.cl_search_input.text())

    def filter_credited_leaves_table(self, text=""):
        """فلترة جدول أيام الاحتساب بالبحث النصي."""
        if not hasattr(self, "credited_leaves_table") or self.credited_leaves_table is None:
            return
        q = str(text or (self.cl_search_input.text() if hasattr(self, "cl_search_input") else "")).strip().lower()
        for r in range(self.credited_leaves_table.rowCount()):
            c_it = self.credited_leaves_table.item(r, 2)
            n_it = self.credited_leaves_table.item(r, 3)
            d_it = self.credited_leaves_table.item(r, 4)
            r_it = self.credited_leaves_table.item(r, 6)

            t_code = c_it.text().lower() if c_it else ""
            t_name = n_it.text().lower() if n_it else ""
            t_dept = d_it.text().lower() if d_it else ""
            t_reason = r_it.text().lower() if r_it else ""

            match = (not q or q in t_code or q in t_name or q in t_dept or q in t_reason)
            self.credited_leaves_table.setRowHidden(r, not match)

    def edit_credited_leave_dialog(self, record_id, refresh_callback=None):
        """نافذة تعديل سجل أيام الاحتساب."""
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT date, employee_id, days, reason, notes FROM credited_leaves WHERE id=?", (record_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            return

        cur_date, emp_id, cur_days, cur_reason, cur_notes = row
        dlg = QDialog(self)
        dlg.setWindowTitle(f"تعديل سجل أيام الاحتساب #{record_id}")
        dlg.setLayoutDirection(Qt.RightToLeft)
        dlg.resize(440, 340)
        dlg_layout = QVBoxLayout(dlg)
        dlg_layout.setSpacing(12)

        form = QFormLayout()
        form.setSpacing(10)

        date_edit = QDateEdit()
        date_edit.setDate(QDate.fromString(cur_date, "yyyy-MM-dd"))
        date_edit.setCalendarPopup(True)
        form.addRow("التاريخ:", date_edit)

        emp_combo = QComboBox()
        self.populate_employee_combo(emp_combo, emp_id)
        form.addRow("اسم العامل:", emp_combo)

        days_spin = QDoubleSpinBox()
        days_spin.setRange(0.5, 30.0)
        days_spin.setSingleStep(0.5)
        days_spin.setValue(float(cur_days or 1.0))
        days_spin.setSuffix(" يوم")
        form.addRow("الأيام المحتسبة:", days_spin)

        reason_input = QLineEdit(cur_reason or "")
        form.addRow("السبب:", reason_input)

        notes_input = QLineEdit(cur_notes or "")
        form.addRow("ملاحظات:", notes_input)

        dlg_layout.addLayout(form)

        btn_box = QHBoxLayout()
        save_btn = QPushButton("💾 حفظ التعديل")
        save_btn.setStyleSheet("background: #7c3aed; color: white; padding: 8px 18px; font-weight: bold; border-radius: 6px;")
        cancel_btn = QPushButton("إلغاء")
        cancel_btn.setStyleSheet("background: #64748b; color: white; padding: 8px 18px; border-radius: 6px;")
        btn_box.addWidget(save_btn)
        btn_box.addWidget(cancel_btn)
        dlg_layout.addLayout(btn_box)

        cancel_btn.clicked.connect(dlg.reject)

        def save_changes():
            new_date = date_edit.date().toString("yyyy-MM-dd")
            new_eid = emp_combo.currentData()
            new_days = days_spin.value()
            new_r = reason_input.text().strip()
            new_n = notes_input.text().strip()

            res = update_credited_leave(record_id, new_eid, new_days, new_date, new_r, new_n)
            if res.get("success"):
                dlg.accept()
                if refresh_callback:
                    refresh_callback()
                self.load_credited_leaves_table()
                if hasattr(self, "load_payroll_table"):
                    self.load_payroll_table()
                self.refresh_dashboard_data()
                if hasattr(self, "show_toast"):
                    self.show_toast("✓ تم تعديل سجل أيام الاحتساب بنجاح")
            else:
                QMessageBox.critical(dlg, "خطأ", f"تعذر حفظ التعديل: {res.get('error')}")

        save_btn.clicked.connect(save_changes)
        dlg.exec()

    def open_add_credited_leave_dialog(self, employee_id=None, default_days=1.0, callback=None):
        """نافذة سريعة لإضافة أيام احتساب لعامل محدد مباشرة من أي شاشة."""
        dlg = QDialog(self)
        dlg.setWindowTitle("✨ تسجيل أيام احتساب (مدفوعة على المصنع)")
        dlg.setLayoutDirection(Qt.RightToLeft)
        dlg.resize(440, 320)
        dlg.setStyleSheet("background: white; font-family: 'Segoe UI';")

        dlg_layout = QVBoxLayout(dlg)
        dlg_layout.setSpacing(12)
        dlg_layout.setContentsMargins(16, 16, 16, 16)

        header_frame = QFrame()
        header_frame.setStyleSheet("background: #f5f3ff; border: 1.5px solid #c4b5fd; border-radius: 8px; padding: 10px;")
        h_box = QVBoxLayout(header_frame)
        h_box.setSpacing(2)
        t_lbl = QLabel("🎁 تسجيل أيام احتساب (على المصنع)")
        t_lbl.setFont(QFont("Segoe UI", 11, QFont.Bold))
        t_lbl.setStyleSheet("color: #6b21a8; border: none;")
        s_lbl = QLabel("أيام الاحتساب تعفي العامل من الخصم وتتحملها إدارة المصنع بالكامل.")
        s_lbl.setFont(QFont("Segoe UI", 9))
        s_lbl.setStyleSheet("color: #7c3aed; border: none;")
        h_box.addWidget(t_lbl)
        h_box.addWidget(s_lbl)
        dlg_layout.addWidget(header_frame)

        form = QFormLayout()
        form.setSpacing(10)

        emp_combo = QComboBox()
        emp_combo.setStyleSheet("font-size: 12px; font-weight: bold; padding: 5px 8px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white;")
        self.populate_employee_combo(emp_combo, employee_id)
        form.addRow("اسم العامل:", emp_combo)

        days_spin = QDoubleSpinBox()
        days_spin.setRange(0.5, 30.0)
        days_spin.setSingleStep(0.5)
        days_spin.setValue(float(default_days or 1.0))
        days_spin.setSuffix(" يوم")
        days_spin.setStyleSheet("font-size: 12px; font-weight: bold; padding: 5px 8px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white;")
        form.addRow("الأيام المحتسبة:", days_spin)

        date_edit = QDateEdit()
        date_edit.setDate(self.att_date.date() if hasattr(self, "att_date") else QDate.currentDate())
        date_edit.setDisplayFormat("yyyy-MM-dd")
        date_edit.setCalendarPopup(True)
        date_edit.setStyleSheet("font-size: 12px; font-weight: bold; padding: 5px 8px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white;")
        form.addRow("التاريخ:", date_edit)

        reason_input = QLineEdit()
        reason_input.setPlaceholderText("مثال: إذن إدارة / ظرف طارئ / احتساب غياب...")
        reason_input.setStyleSheet("font-size: 12px; padding: 5px 8px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white;")
        form.addRow("السبب:", reason_input)

        notes_input = QLineEdit()
        notes_input.setPlaceholderText("ملاحظات إضافية...")
        notes_input.setStyleSheet("font-size: 12px; padding: 5px 8px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white;")
        form.addRow("ملاحظات:", notes_input)

        dlg_layout.addLayout(form)

        btn_box = QHBoxLayout()
        save_btn = QPushButton("💾 تسجيل الاحتساب")
        save_btn.setCursor(Qt.PointingHandCursor)
        save_btn.setStyleSheet("background: #7c3aed; color: white; padding: 8px 18px; font-weight: bold; border-radius: 6px; font-size: 12px;")
        cancel_btn = QPushButton("إلغاء")
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.setStyleSheet("background: #64748b; color: white; padding: 8px 18px; border-radius: 6px; font-size: 12px;")
        btn_box.addWidget(save_btn)
        btn_box.addWidget(cancel_btn)
        dlg_layout.addLayout(btn_box)

        cancel_btn.clicked.connect(dlg.reject)

        def do_save():
            sel_eid = emp_combo.currentData()
            if not sel_eid:
                QMessageBox.warning(dlg, "تنبيه", "يرجى اختيار عامل أولاً.")
                return
            num_days = days_spin.value()
            d_str = date_edit.date().toString("yyyy-MM-dd")
            r_str = reason_input.text().strip()
            n_str = notes_input.text().strip()

            res = add_credited_leave(sel_eid, num_days, d_str, r_str, n_str)
            if res.get("success"):
                dlg.accept()
                if callback:
                    callback()
                if hasattr(self, "load_credited_leaves_table"):
                    self.load_credited_leaves_table()
                if hasattr(self, "load_payroll_table"):
                    self.load_payroll_table()
                self.refresh_dashboard_data()
                if hasattr(self, "show_toast"):
                    self.show_toast(f"✓ تم تسجيل {num_days} يوم احتساب للعامل بنجاح")
            else:
                QMessageBox.critical(dlg, "خطأ", f"تعذر تسجيل أيام الاحتساب: {res.get('error')}")

        save_btn.clicked.connect(do_save)
        dlg.exec()

    def calculate_work_hours(self, t_in, t_out):
        """حساب عدد ساعات العمل بدقة بين وقت الحضور ووقت الانصراف."""
        if not t_in.isValid() or not t_out.isValid():
            return 0.0
        minutes = (t_out.hour() * 60 + t_out.minute()) - (t_in.hour() * 60 + t_in.minute())
        if minutes < 0:
            minutes += 24 * 60  # عبر منتصف الليل
        return round(minutes / 60.0, 2)

    def on_row_time_changed(self, row_idx):
        """تحديث ساعات العمل فورياً في الصف عند تغيير وقت الحضور أو الانصراف."""
        combo = self.att_table.cellWidget(row_idx, 2)
        if combo and combo.currentText() == "حضور":
            in_edit = self.att_table.cellWidget(row_idx, 3)
            out_edit = self.att_table.cellWidget(row_idx, 4)
            hours_item = self.att_table.item(row_idx, 5)
            if in_edit and out_edit:
                hours = self.calculate_work_hours(in_edit.time(), out_edit.time())
                if hours_item:
                    hours_item.setText(f"{hours:.1f} س")
                    hours_item.setForeground(QColor("#059669" if hours > 0 else "#94a3b8"))
        if getattr(self, "_att_is_saved", False):
            self.recalculate_attendance_metrics_from_table()

    def load_attendance(self):
        """تحميل وتجهيز كشوف الحضور والغياب لليوم المحدد."""
        if not hasattr(self, "att_table"):
            return

        self.att_table.blockSignals(True)
        selected_date = self.att_date.date()
        date_str = selected_date.toString("yyyy-MM-dd")

        if self.is_friday(selected_date):
            self.att_table.setEnabled(False)
            self.save_att_btn.setEnabled(False)
            if hasattr(self, "card_total_workers"):
                self.card_total_workers.val_label.setText("0")
                self.card_present_workers.val_label.setText("0")
                self.card_absent_workers.val_label.setText("0")
                self.card_attendance_pct.val_label.setText("إجازة جمعة 🕌")
                self.card_total_work_hours.val_label.setText("0.0 س")
            self.att_table.blockSignals(False)
            return

        with self.preserve_table_scroll(self.att_table):
            self.att_table.setUpdatesEnabled(False)
            self.att_table.setRowCount(0)
            self.att_table.setEnabled(True)
            self.save_att_btn.setEnabled(True)
            prev_date = self.previous_work_date(selected_date)

            conn = get_db_connection()
            c = conn.cursor()
            c.execute("""
                SELECT e.id, e.name, COALESCE(d.name, 'عام'), e.department_id,
                       COALESCE(d.start_time, '08:00 ص'), COALESCE(d.end_time, '07:00 م'),
                       COALESCE(d.default_work_hours, 11.0), COALESCE(e.code, '')
                FROM employees e
                LEFT JOIN departments d ON e.department_id = d.id
                WHERE e.is_active = 1
                ORDER BY e.name ASC
            """)
            employees = c.fetchall()

            c.execute("SELECT COUNT(*) FROM attendance WHERE date=?", (date_str,))
            is_saved = (c.fetchone()[0] > 0)
            self._att_is_saved = is_saved

            present_count = 0
            absent_count = 0
            total_hours_today = 0.0

            for emp_idx, (emp_id, name, dept_name, dept_id, dept_start, dept_end, dept_hours, emp_code) in enumerate(employees):
                r = self.att_table.rowCount()
                self.att_table.insertRow(r)

                # Col 0: # (Index)
                idx_item = QTableWidgetItem(str(emp_idx + 1))
                idx_item.setTextAlignment(Qt.AlignCenter)
                idx_item.setFont(QFont("Segoe UI", 10))
                idx_item.setFlags(idx_item.flags() & ~Qt.ItemIsEditable)
                self.att_table.setItem(r, 0, idx_item)

                # Col 1: Employee Name & Department Tag
                code_tag = f"({emp_code}) " if emp_code else ""
                name_item = QTableWidgetItem(f"{code_tag}{name}  [{dept_name}]")
                name_item.setData(Qt.UserRole, emp_id)
                name_item.setData(Qt.UserRole + 1, name)
                name_item.setData(Qt.UserRole + 2, dept_id)
                name_item.setData(Qt.UserRole + 3, dept_name)
                name_item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                self.att_table.setItem(r, 1, name_item)

                # Check saved record
                c.execute("""SELECT status, check_in_time, check_out_time, work_hours, notes 
                             FROM attendance WHERE employee_id=? AND date=?""", (emp_id, date_str))
                saved_att = c.fetchone()

                status = saved_att[0] if saved_att and saved_att[0] else "حضور"
                in_str = saved_att[1] if saved_att and saved_att[1] else dept_start
                out_str = saved_att[2] if saved_att and saved_att[2] else dept_end
                hours_val = saved_att[3] if (saved_att and saved_att[3] is not None) else dept_hours

                # Col 2: Status ComboBox (حضور / غياب)
                combo = QComboBox()
                combo.addItems(["حضور", "غياب"])
                combo.setCursor(Qt.PointingHandCursor)
                combo.setCurrentText(status if status in ["حضور", "غياب"] else "حضور")
                self.style_attendance_combo(combo, status)

                # Col 3: Check-in Time
                in_edit = QTimeEdit()
                in_edit.setLocale(ARABIC_LOCALE)
                in_edit.setDisplayFormat("hh:mm AP")
                t_in = parse_time_str(in_str)
                in_edit.setTime(t_in if t_in.isValid() else QTime(8, 0))
                in_edit.setStyleSheet("""
                    QTimeEdit {
                        font-size: 12px; font-weight: bold; padding: 4px 8px;
                        border: 1.5px solid #cbd5e1; border-radius: 6px;
                        background: white; min-width: 120px; min-height: 28px;
                        color: #0f172a;
                    }
                    QTimeEdit:focus { border: 2px solid #2563eb; }
                """)

                # Col 4: Check-out Time
                out_edit = QTimeEdit()
                out_edit.setLocale(ARABIC_LOCALE)
                out_edit.setDisplayFormat("hh:mm AP")
                t_out = parse_time_str(out_str)
                out_edit.setTime(t_out if t_out.isValid() else QTime(19, 0))
                out_edit.setStyleSheet("""
                    QTimeEdit {
                        font-size: 12px; font-weight: bold; padding: 4px 8px;
                        border: 1.5px solid #cbd5e1; border-radius: 6px;
                        background: white; min-width: 120px; min-height: 28px;
                        color: #0f172a;
                    }
                    QTimeEdit:focus { border: 2px solid #2563eb; }
                """)

                # Col 5: Hours
                if status == "حضور":
                    calculated_h = self.calculate_work_hours(in_edit.time(), out_edit.time())
                else:
                    calculated_h = 0.0
                    in_edit.setEnabled(False)
                    out_edit.setEnabled(False)

                hours_item = QTableWidgetItem(f"{calculated_h:.1f} س")
                hours_item.setTextAlignment(Qt.AlignCenter)
                hours_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                if status == "حضور":
                    hours_item.setForeground(QColor("#059669" if calculated_h > 0 else "#94a3b8"))
                else:
                    hours_item.setForeground(QColor("#94a3b8"))
                self.att_table.setItem(r, 5, hours_item)

                # Wire signals for instant hours update
                in_edit.timeChanged.connect(lambda _, row_idx=r: self.on_row_time_changed(row_idx))
                out_edit.timeChanged.connect(lambda _, row_idx=r: self.on_row_time_changed(row_idx))

                # Connect combo signal
                combo.currentTextChanged.connect(
                    lambda new_status, eid=emp_id, row_idx=r: self.on_attendance_combo_changed(eid, new_status, row_idx)
                )

                self.att_table.setCellWidget(r, 2, combo)
                self.att_table.setCellWidget(r, 3, in_edit)
                self.att_table.setCellWidget(r, 4, out_edit)

                if status == "حضور":
                    present_count += 1
                    total_hours_today += calculated_h
                else:
                    absent_count += 1

                # Col 6: Consecutive absence warning / Notes
                c.execute("SELECT status FROM attendance WHERE employee_id=? AND date=?", (emp_id, prev_date))
                prev_stat = c.fetchone()
                if status == "غياب" and prev_stat and prev_stat[0] == "غياب":
                    item = QTableWidgetItem("⚠️ تنبيه: غياب يومين متتاليين!")
                    item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                    item.setForeground(QColor("#dc2626"))
                    self.att_table.setItem(r, 6, item)
                else:
                    note_str = saved_att[4] if saved_att and saved_att[4] else ""
                    self.att_table.setItem(r, 6, QTableWidgetItem(note_str))

                # Col 7: Absence details button
                abs_btn = QPushButton("📅 أيام الغياب")
                abs_btn.setCursor(Qt.PointingHandCursor)
                abs_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #3b82f6; color: white; font-size: 11px; font-weight: bold;
                        padding: 4px 10px; border-radius: 5px; border: none; min-height: 26px;
                    }
                    QPushButton:hover { background-color: #2563eb; }
                """)
                abs_btn.clicked.connect(lambda _, eid=emp_id, ename=name: self.show_employee_absences(eid, ename))
                self.att_table.setCellWidget(r, 7, abs_btn)

            self.att_table.setUpdatesEnabled(True)
            conn.close()
        self.att_table.blockSignals(False)

        if is_saved:
            self.update_attendance_metrics(len(employees), present_count, absent_count, total_hours_today)
        else:
            if hasattr(self, "card_total_workers"):
                self.card_total_workers.val_label.setText(str(len(employees)))
                self.card_present_workers.val_label.setText("-")
                self.card_absent_workers.val_label.setText("-")
                self.card_attendance_pct.val_label.setText("لم يتم الحفظ")
                self.card_total_work_hours.val_label.setText("لم يتم الحفظ")

        if hasattr(self, "att_search_input") and self.att_search_input.text():
            self.filter_attendance_table(self.att_search_input.text())

    def on_attendance_combo_changed(self, emp_id, new_status, row_idx):
        selected_date = self.att_date.date()
        if self.is_friday(selected_date):
            return

        # 1. Update style of combo box
        combo = self.att_table.cellWidget(row_idx, 2)
        in_edit = self.att_table.cellWidget(row_idx, 3)
        out_edit = self.att_table.cellWidget(row_idx, 4)
        hours_item = self.att_table.item(row_idx, 5)

        if combo:
            self.style_attendance_combo(combo, new_status)

        if new_status == "حضور":
            if in_edit: in_edit.setEnabled(True)
            if out_edit: out_edit.setEnabled(True)
            if in_edit and out_edit:
                h = self.calculate_work_hours(in_edit.time(), out_edit.time())
                if hours_item:
                    hours_item.setText(f"{h:.1f} س")
                    hours_item.setForeground(QColor("#059669" if h > 0 else "#94a3b8"))
        else:
            # غياب
            if in_edit: in_edit.setEnabled(False)
            if out_edit: out_edit.setEnabled(False)
            if hours_item:
                hours_item.setText("0.0 س")
                hours_item.setForeground(QColor("#94a3b8"))

        # 2. Update consecutive absence warning
        prev_date = self.previous_work_date(selected_date)
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT status FROM attendance WHERE employee_id=? AND date=?", (emp_id, prev_date))
        prev_stat = c.fetchone()
        conn.close()

        if new_status == "غياب" and prev_stat and prev_stat[0] == "غياب":
            item = QTableWidgetItem("⚠️ تنبيه: غياب يومين متتاليين!")
            item.setFont(QFont("Segoe UI", 10, QFont.Bold))
            item.setForeground(QColor("#dc2626"))
            self.att_table.setItem(row_idx, 6, item)
        else:
            self.att_table.setItem(row_idx, 6, QTableWidgetItem(""))

        self.recalculate_attendance_metrics_from_table()

    def recalculate_attendance_metrics_from_table(self):
        total = self.att_table.rowCount()
        present = 0
        absent = 0
        total_hours = 0.0
        for r in range(total):
            combo = self.att_table.cellWidget(r, 2)
            hours_item = self.att_table.item(r, 5)
            if combo:
                st = combo.currentText()
                if st == "حضور":
                    present += 1
                    if hours_item:
                        try:
                            total_hours += float(hours_item.text().replace(" س", "").strip())
                        except Exception:
                            pass
                else:
                    absent += 1
        self.update_attendance_metrics(total, present, absent, total_hours)

    def update_attendance_metrics(self, total, present, absent, total_hours=0.0):
        if hasattr(self, "card_total_workers"):
            self.card_total_workers.val_label.setText(str(total))
            self.card_present_workers.val_label.setText(str(present))
            self.card_absent_workers.val_label.setText(str(absent))
            pct = (present / total * 100) if total > 0 else 0
            self.card_attendance_pct.val_label.setText(f"{pct:.1f}%")
            self.card_total_work_hours.val_label.setText(f"{total_hours:.1f} س")

    def save_attendance(self):
        selected_date = self.att_date.date()
        if self.is_friday(selected_date):
            QMessageBox.information(self, "إجازة", "يوم الجمعة إجازة رسمية ولا يتم حفظ كشوف حضور له.")
            return

        date_str = selected_date.toString("yyyy-MM-dd")
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("DELETE FROM attendance WHERE date=?", (date_str,))

        for r in range(self.att_table.rowCount()):
            name_item = self.att_table.item(r, 1)
            combo = self.att_table.cellWidget(r, 2)
            in_edit = self.att_table.cellWidget(r, 3)
            out_edit = self.att_table.cellWidget(r, 4)
            notes_item = self.att_table.item(r, 6)

            if name_item and combo:
                emp_id = name_item.data(Qt.UserRole)
                if emp_id is not None:
                    status = combo.currentText()
                    if status == "حضور" and in_edit and out_edit:
                        in_str = format_time_ar(in_edit.time())
                        out_str = format_time_ar(out_edit.time())
                        hours = self.calculate_work_hours(in_edit.time(), out_edit.time())
                    else:
                        in_str = ""
                        out_str = ""
                        hours = 0.0

                    notes_str = notes_item.text() if notes_item else ""
                    c.execute("""
                        INSERT INTO attendance (date, employee_id, status, check_in_time, check_out_time, work_hours, notes) 
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (date_str, emp_id, status, in_str, out_str, hours, notes_str))

        conn.commit()
        conn.close()

        self.load_attendance()
        self.refresh_dashboard_data()
        if hasattr(self, "show_toast"):
            self.show_toast("✓ تم حفظ كشوف الحضور ومواعيد الورديات لليوم بنجاح!")
        else:
            QMessageBox.information(self, "حفظ", "تم حفظ كشوف الحضور ومواعيد الدخول والانصراف بنجاح.")

    def populate_dept_filter_combo(self, combo):
        combo.blockSignals(True)
        cur_data = combo.currentData()
        combo.clear()
        combo.addItem("🏢 جميع الأقسام (الكل)", 0)
        for row in get_departments():
            dept_id = row[0]
            dept_name = row[1]
            combo.addItem(f"📁 {dept_name}", dept_id)
        if cur_data is not None:
            idx = combo.findData(cur_data)
            if idx >= 0:
                combo.setCurrentIndex(idx)
        combo.blockSignals(False)

    def filter_attendance_table(self, _=None):
        query = self.att_search_input.text().strip().lower() if hasattr(self, "att_search_input") else ""
        selected_dept_id = self.att_dept_filter.currentData() if hasattr(self, "att_dept_filter") else 0

        for r in range(self.att_table.rowCount()):
            name_item = self.att_table.item(r, 1)
            combo = self.att_table.cellWidget(r, 2)

            worker_full_text = name_item.text().lower() if name_item else ""
            status = combo.currentText().lower() if combo else ""
            dept_id = name_item.data(Qt.UserRole + 2) if name_item else None

            match_query = (not query or query in worker_full_text or query in status)
            match_dept = (selected_dept_id == 0 or selected_dept_id is None or dept_id == selected_dept_id)

            if match_query and match_dept:
                self.att_table.setRowHidden(r, False)
            else:
                self.att_table.setRowHidden(r, True)

    def show_employee_absences(self, emp_id, emp_name):
        """عرض نافذة تفاصيل أيام غياب العامل وأيام الاحتساب ضمن الدورة المالية الحالية."""
        sel = self.att_date.date() if hasattr(self, "att_date") else QDate.currentDate()
        year, month = sel.year(), sel.month()
        start_date_str, end_date_str, cycle_desc = get_payroll_cycle_dates(year, month)

        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            SELECT date, status FROM attendance 
            WHERE employee_id=? AND status='غياب' AND date BETWEEN ? AND ? 
            ORDER BY date ASC
        """, (emp_id, start_date_str, end_date_str))
        rows = c.fetchall()

        c.execute("""
            SELECT COALESCE(SUM(days), 0.0) FROM credited_leaves
            WHERE employee_id=? AND date BETWEEN ? AND ?
        """, (emp_id, start_date_str, end_date_str))
        credited_days = float(c.fetchone()[0] or 0.0)
        conn.close()

        effective_absent = max(0.0, float(len(rows)) - credited_days)

        dlg = QDialog(self)
        dlg.setWindowTitle(f"سجل غياب واحتساب العامل: {emp_name}")
        dlg.setLayoutDirection(Qt.RightToLeft)
        dlg.resize(500, 480)
        dlg.setStyleSheet("background: white; font-family: 'Segoe UI';")
        dlg_layout = QVBoxLayout(dlg)
        dlg_layout.setSpacing(12)
        dlg_layout.setContentsMargins(16, 16, 16, 16)

        header_frame = QFrame()
        header_frame.setStyleSheet("background: #f8fafc; border: 1.5px solid #cbd5e1; border-radius: 8px; padding: 10px;")
        h_box = QVBoxLayout(header_frame)
        h_box.setSpacing(4)
        t_lbl = QLabel(f"👤 سجل أيام الغياب والاحتساب للعامل: <b>{emp_name}</b>")
        t_lbl.setFont(QFont("Segoe UI", 11, QFont.Bold))
        t_lbl.setStyleSheet("color: #0f172a; border: none;")
        p_lbl = QLabel(f"الفترة المالية: {start_date_str} إلى {end_date_str} ({cycle_desc})")
        p_lbl.setFont(QFont("Segoe UI", 9))
        p_lbl.setStyleSheet("color: #64748b; border: none;")
        h_box.addWidget(t_lbl)
        h_box.addWidget(p_lbl)
        dlg_layout.addWidget(header_frame)

        # 3 KPI Badges (إجمالي الغياب | أيام الاحتساب على المصنع | الغياب المخصوم الصافي)
        badge_frame = QFrame()
        badge_frame.setStyleSheet("background: transparent; border: none;")
        b_layout = QHBoxLayout(badge_frame)
        b_layout.setContentsMargins(0, 0, 0, 0)
        b_layout.setSpacing(8)

        def _make_badge(title, val, color, bg):
            f = QFrame()
            f.setStyleSheet(f"background: {bg}; border: 1px solid #cbd5e1; border-radius: 6px; padding: 6px 10px;")
            l = QVBoxLayout(f)
            l.setContentsMargins(0, 0, 0, 0)
            l.setSpacing(2)
            lbl_t = QLabel(title)
            lbl_t.setFont(QFont("Segoe UI", 8, QFont.Bold))
            lbl_t.setStyleSheet("color: #64748b;")
            lbl_v = QLabel(val)
            lbl_v.setFont(QFont("Segoe UI", 12, QFont.Bold))
            lbl_v.setStyleSheet(f"color: {color};")
            l.addWidget(lbl_t)
            l.addWidget(lbl_v)
            return f

        b_layout.addWidget(_make_badge("❌ إجمالي الغياب", f"{len(rows)} يوم", "#dc2626", "#fef2f2"))
        b_layout.addWidget(_make_badge("🎁 أيام الاحتساب (المصنع)", f"{credited_days:g} يوم", "#7c3aed", "#f5f3ff"))
        b_layout.addWidget(_make_badge("📉 الغياب المخصوم", f"{effective_absent:g} يوم", "#b45309" if effective_absent > 0 else "#15803d", "#fefce8" if effective_absent > 0 else "#f0fdf4"))
        dlg_layout.addWidget(badge_frame)

        if rows:
            tbl = QTableWidget(len(rows), 3)
            tbl.setHorizontalHeaderLabels(["#", "التاريخ", "اليوم"])
            tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            tbl.verticalHeader().setVisible(False)
            tbl.setStyleSheet("""
                QTableWidget { border: 1px solid #e2e8f0; border-radius: 6px; background: white; }
                QTableWidget::item { padding: 8px; font-size: 12px; }
                QHeaderView::section { background: #0f172a; color: white; padding: 8px; font-weight: bold; }
            """)
            day_names = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]
            for i, (date_val, st) in enumerate(rows):
                dt = datetime.strptime(date_val, "%Y-%m-%d")
                dname = day_names[dt.weekday()]

                idx_item = QTableWidgetItem(str(i + 1))
                idx_item.setTextAlignment(Qt.AlignCenter)
                d_item = QTableWidgetItem(date_val)
                d_item.setTextAlignment(Qt.AlignCenter)
                dname_item = QTableWidgetItem(dname)
                dname_item.setTextAlignment(Qt.AlignCenter)
                dname_item.setFont(QFont("Segoe UI", 10, QFont.Bold))

                tbl.setItem(i, 0, idx_item)
                tbl.setItem(i, 1, d_item)
                tbl.setItem(i, 2, dname_item)
            dlg_layout.addWidget(tbl)
        else:
            no_lbl = QLabel("🎉 ممتاز! لا توجد أي أيام غياب مسجلة لهذا العامل في هذه الفترة.")
            no_lbl.setFont(QFont("Segoe UI", 11, QFont.Bold))
            no_lbl.setStyleSheet("color: #15803d; padding: 20px; background: #f0fdf4; border-radius: 8px; border: 1px solid #86efac;")
            no_lbl.setAlignment(Qt.AlignCenter)
            dlg_layout.addWidget(no_lbl)

        btn_bar = QHBoxLayout()
        add_credit_btn = QPushButton("✨ إضافة أيام احتساب لهذا العامل")
        add_credit_btn.setCursor(Qt.PointingHandCursor)
        add_credit_btn.setStyleSheet("background: #7c3aed; color: white; padding: 8px 16px; border-radius: 6px; font-weight: bold; border: none;")
        add_credit_btn.clicked.connect(lambda: [dlg.accept(), self.open_add_credited_leave_dialog(emp_id, callback=lambda: self.show_employee_absences(emp_id, emp_name))])
        btn_bar.addWidget(add_credit_btn)

        btn_bar.addStretch()

        close_btn = QPushButton("إغلاق")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet("background: #64748b; color: white; padding: 8px 20px; border-radius: 6px; font-weight: bold; border: none;")
        close_btn.clicked.connect(dlg.accept)
        btn_bar.addWidget(close_btn)

        dlg_layout.addLayout(btn_bar)
        dlg.exec()

    def style_attendance_combo(self, combo, status):
        """تنسيق وتجميل قائمة حالة الحضور والغياب في جدول الحضور."""
        if not combo:
            return
        combo.setFont(QFont("Segoe UI", 10, QFont.Bold))
        status_clean = str(status).strip()
        if "حضور" in status_clean:
            combo.setStyleSheet("""
                QComboBox {
                    background-color: #dcfce7;
                    color: #15803d;
                    border: 1.5px solid #86efac;
                    border-radius: 6px;
                    padding: 3px 8px;
                    font-weight: bold;
                    font-size: 12px;
                    min-width: 110px;
                    min-height: 28px;
                }
                QComboBox:hover {
                    background-color: #bbf7d0;
                    border-color: #4ade80;
                }
                QComboBox::drop-down { border: none; width: 18px; }
                QComboBox::down-arrow {
                    border-left: 4px solid transparent;
                    border-right: 4px solid transparent;
                    border-top: 5px solid #15803d;
                    width: 0px; height: 0px; margin-right: 4px;
                }
                QComboBox QAbstractItemView {
                    background-color: #ffffff;
                    color: #0f172a;
                    border: 1.5px solid #cbd5e1;
                    border-radius: 6px;
                    padding: 4px;
                    selection-background-color: #dcfce7;
                    selection-color: #15803d;
                    font-weight: bold;
                }
            """)
        else:
            # غياب
            combo.setStyleSheet("""
                QComboBox {
                    background-color: #fee2e2;
                    color: #b91c1c;
                    border: 1.5px solid #fca5a5;
                    border-radius: 6px;
                    padding: 3px 8px;
                    font-weight: bold;
                    font-size: 12px;
                    min-width: 110px;
                    min-height: 28px;
                }
                QComboBox:hover {
                    background-color: #fecaca;
                    border-color: #f87171;
                }
                QComboBox::drop-down { border: none; width: 18px; }
                QComboBox::down-arrow {
                    border-left: 4px solid transparent;
                    border-right: 4px solid transparent;
                    border-top: 5px solid #b91c1c;
                    width: 0px; height: 0px; margin-right: 4px;
                }
                QComboBox QAbstractItemView {
                    background-color: #ffffff;
                    color: #0f172a;
                    border: 1.5px solid #cbd5e1;
                    border-radius: 6px;
                    padding: 4px;
                    selection-background-color: #fee2e2;
                    selection-color: #b91c1c;
                    font-weight: bold;
                }
            """)

    def show_todays_absentees(self):
        """عرض قائمة بأسماء الغائبين لليوم المحدد."""
        selected_date = self.att_date.date().toString("yyyy-MM-dd")
        display_date = self.att_date.date().toString("yyyy/MM/dd")

        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            SELECT e.name, COALESCE(d.name, 'عام')
            FROM attendance a
            JOIN employees e ON a.employee_id = e.id
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE a.date = ? AND a.status = 'غياب'
            ORDER BY e.name ASC
        """, (selected_date,))
        rows = c.fetchall()
        conn.close()

        dlg = QDialog(self)
        dlg.setWindowTitle(f"الغائبون - {display_date}")
        dlg.setLayoutDirection(Qt.RightToLeft)
        dlg.resize(420, 460)
        dlg.setStyleSheet("background: white; font-family: 'Segoe UI';")
        v = QVBoxLayout(dlg)
        v.setSpacing(12)
        v.setContentsMargins(18, 18, 18, 18)

        title = QLabel(f"❌ قائمة الغائبين بتاريخ  <b>{display_date}</b>")
        title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        title.setStyleSheet("color: #b91c1c;")
        title.setAlignment(Qt.AlignCenter)
        v.addWidget(title)

        if not rows:
            no_lbl = QLabel("✅ لا يوجد أي غائبين في هذا اليوم — نسبة الحضور 100%!")
            no_lbl.setFont(QFont("Segoe UI", 11, QFont.Bold))
            no_lbl.setStyleSheet("color: #15803d; padding: 16px; background: #f0fdf4; border-radius: 8px; border: 1px solid #86efac;")
            no_lbl.setAlignment(Qt.AlignCenter)
            v.addWidget(no_lbl)
        else:
            count_lbl = QLabel(f"عدد الغائبين: <b style='color:#b91c1c'>{len(rows)}</b> عامل")
            count_lbl.setFont(QFont("Segoe UI", 11))
            count_lbl.setAlignment(Qt.AlignCenter)
            v.addWidget(count_lbl)

            list_frame = QFrame()
            list_frame.setStyleSheet("background: #fef2f2; border: 1.5px solid #fca5a5; border-radius: 8px; padding: 6px;")
            list_layout = QVBoxLayout(list_frame)
            list_layout.setSpacing(6)
            list_layout.setContentsMargins(10, 8, 10, 8)

            for idx, (name, dept) in enumerate(rows, 1):
                row_lbl = QLabel(f"<b style='color:#b91c1c'>{idx}.</b>  {name}  <span style='color:#64748b;'>[{dept}]</span>")
                row_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
                row_lbl.setStyleSheet("background: white; border-radius: 5px; padding: 6px 10px; border: 1px solid #fed7d7;")
                list_layout.addWidget(row_lbl)

            v.addWidget(list_frame)

        close_btn = QPushButton("إغلاق")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet("background: #64748b; color: white; font-size: 12px; font-weight: bold; padding: 7px 20px; border-radius: 6px; border: none;")
        close_btn.clicked.connect(dlg.accept)
        v.addWidget(close_btn, alignment=Qt.AlignLeft)
        dlg.exec()

    def open_employees_manager(self):
        """فتح نافذة متكاملة وشاملة لإدارة العاملين (إضافة، تعديل، عرض الملف الكامل، تصدير، وتنشيط وإلغاء)."""
        dlg = QDialog(self)
        dlg.setWindowTitle("👥 إدارة أسماء وملفات وبيانات العاملين والموظفين")
        dlg.setLayoutDirection(Qt.RightToLeft)
        dlg.resize(1180, 700)
        dlg.setStyleSheet("background: #ffffff; font-family: 'Segoe UI';")

        layout = QVBoxLayout(dlg)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # Header
        hdr = QFrame()
        hdr.setStyleSheet("background: #0f172a; border-radius: 8px; padding: 12px;")
        h_layout = QHBoxLayout(hdr)
        lbl_h = QLabel("👥 إدارة ملفات وبيانات العاملين والموظفين")
        lbl_h.setFont(QFont("Segoe UI", 13, QFont.Bold))
        lbl_h.setStyleSheet("color: white; border: none;")
        h_layout.addWidget(lbl_h)
        h_layout.addStretch()

        btn_export_emps = QPushButton("📥 تصدير كشف العاملين (Excel)")
        btn_export_emps.setCursor(Qt.PointingHandCursor)
        btn_export_emps.setStyleSheet("""
            QPushButton { background-color: #0284c7; color: white; padding: 7px 16px; font-weight: bold; border-radius: 6px; border: none; font-size: 12px; }
            QPushButton:hover { background-color: #0369a1; }
        """)
        h_layout.addWidget(btn_export_emps)

        btn_add_emp = QPushButton("➕ إضافة عامل جديد")
        btn_add_emp.setCursor(Qt.PointingHandCursor)
        btn_add_emp.setStyleSheet("""
            QPushButton { background-color: #16a34a; color: white; padding: 7px 18px; font-weight: bold; border-radius: 6px; border: none; font-size: 12px; }
            QPushButton:hover { background-color: #15803d; }
        """)
        h_layout.addWidget(btn_add_emp)
        layout.addWidget(hdr)

        # Filter Bar
        f_bar = QFrame()
        f_bar.setStyleSheet("background: #f8fafc; border-radius: 8px; padding: 8px; border: 1.5px solid #e2e8f0;")
        f_layout = QHBoxLayout(f_bar)
        f_layout.setSpacing(10)

        f_layout.addWidget(QLabel("🔍 بحث شامل:"))
        search_inp = QLineEdit()
        search_inp.setPlaceholderText("ابحث بالاسم، الكود، الموبايل، أو الرقم القومي...")
        search_inp.setStyleSheet("padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; min-width: 250px; font-size: 12px;")
        f_layout.addWidget(search_inp)

        f_layout.addWidget(QLabel("🏢 القسم:"))
        dept_cb = QComboBox()
        dept_cb.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; min-width: 160px; font-size: 12px;")
        self.populate_dept_filter_combo(dept_cb)
        f_layout.addWidget(dept_cb)

        f_layout.addWidget(QLabel("الحالة:"))
        status_cb = QComboBox()
        status_cb.addItems(["الكل", "نشط فقط", "غير نشط فقط"])
        status_cb.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; min-width: 120px; font-size: 12px;")
        f_layout.addWidget(status_cb)

        f_layout.addWidget(QLabel("التأمين:"))
        ins_cb = QComboBox()
        ins_cb.addItems(["الكل", "مؤمن عليه", "غير مؤمن"])
        ins_cb.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; min-width: 110px; font-size: 12px;")
        f_layout.addWidget(ins_cb)

        f_layout.addStretch()
        layout.addWidget(f_bar)

        # Employees Table
        tbl = QTableWidget()
        headers = [
            "ID", "الكود", "اسم العامل", "🏢 القسم", "📱 الموبايل",
            "🪪 الرقم القومي", "🛡️ التأمين", "💵 الراتب الأساسي",
            "🎁 حافز الانتظام", "الحالة", "👁️ ملف العامل", "✏️ تعديل", "🔄 الحالة"
        ]
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.setAlternatingRowColors(True)
        tbl.verticalHeader().setVisible(False)
        tbl.setStyleSheet("""
            QTableWidget { border: 1.5px solid #cbd5e1; border-radius: 8px; background: white; gridline-color: #e2e8f0; }
            QTableWidget::item { padding: 6px 8px; font-size: 12px; }
            QHeaderView::section { background-color: #0f172a; color: white; padding: 9px 8px; font-weight: bold; font-size: 12px; }
        """)
        thdr = tbl.horizontalHeader()
        thdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(2, QHeaderView.Stretch)
        thdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(8, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(9, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(10, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(11, QHeaderView.ResizeToContents)
        thdr.setSectionResizeMode(12, QHeaderView.ResizeToContents)
        layout.addWidget(tbl)

        raw_employees_cache = []

        def load_emp_table():
            nonlocal raw_employees_cache
            q_txt = search_inp.text().strip().lower()
            sel_dept = dept_cb.currentData()
            sel_st = status_cb.currentIndex()
            sel_ins = ins_cb.currentIndex()

            conn = get_db_connection()
            c = conn.cursor()
            c.execute("""
                SELECT e.id, COALESCE(e.code, ''), e.name, COALESCE(d.name, 'عام'),
                       COALESCE(e.base_salary, 0.0), COALESCE(e.attendance_incentive, 500.0),
                       COALESCE(e.work_hours, 11.0), e.is_active, e.department_id,
                       COALESCE(e.notes, ''), COALESCE(e.hire_date, ''), COALESCE(e.birth_date, ''),
                       COALESCE(e.national_id, ''), COALESCE(e.insurance_status, 'غير مؤمن'),
                       COALESCE(e.insurance_no, ''), COALESCE(e.marital_status, 'أعزب'),
                       COALESCE(e.phone, ''), COALESCE(e.has_relatives, 'لا'),
                       COALESCE(e.relative_name, ''), COALESCE(e.relative_relation, ''),
                       COALESCE(e.address, ''), COALESCE(e.daily_wage, 0.0)
                FROM employees e
                LEFT JOIN departments d ON e.department_id = d.id
                ORDER BY e.is_active DESC, e.name ASC
            """)
            all_emps = c.fetchall()
            conn.close()

            raw_employees_cache = all_emps
            filtered = []
            for emp in all_emps:
                (eid, code, ename, dname, sal, inc, hrs, active, did, notes,
                 hire_d, birth_d, nid, ins_st, ins_no, marital, phone,
                 has_rel, rel_name, rel_rel, addr, dwage) = emp

                if q_txt:
                    match_any = (
                        q_txt in ename.lower() or
                        q_txt in str(code).lower() or
                        q_txt in str(phone).lower() or
                        q_txt in str(nid).lower() or
                        q_txt in str(dname).lower() or
                        q_txt in str(addr).lower()
                    )
                    if not match_any:
                        continue
                if sel_dept and sel_dept > 0 and did != sel_dept:
                    continue
                if sel_st == 1 and active != 1:
                    continue
                if sel_st == 2 and active != 0:
                    continue
                if sel_ins == 1 and "مؤمن" not in ins_st:
                    continue
                if sel_ins == 2 and "غير مؤمن" not in ins_st:
                    continue
                filtered.append(emp)

            tbl.setRowCount(len(filtered))
            for r_idx, emp in enumerate(filtered):
                (eid, code, ename, dname, sal, inc, hrs, active, did, notes,
                 hire_d, birth_d, nid, ins_st, ins_no, marital, phone,
                 has_rel, rel_name, rel_rel, addr, dwage) = emp

                tbl.setRowHeight(r_idx, 38)

                # Col 0: ID
                i0 = QTableWidgetItem(str(eid))
                i0.setTextAlignment(Qt.AlignCenter)
                tbl.setItem(r_idx, 0, i0)

                # Col 1: Code
                i1 = QTableWidgetItem(str(code or "-"))
                i1.setTextAlignment(Qt.AlignCenter)
                i1.setFont(QFont("Segoe UI", 10, QFont.Bold))
                tbl.setItem(r_idx, 1, i1)

                # Col 2: Name
                i2 = QTableWidgetItem(ename)
                i2.setFont(QFont("Segoe UI", 10, QFont.Bold))
                tbl.setItem(r_idx, 2, i2)

                # Col 3: Dept
                i3 = QTableWidgetItem(dname)
                i3.setTextAlignment(Qt.AlignCenter)
                tbl.setItem(r_idx, 3, i3)

                # Col 4: Phone
                i4 = QTableWidgetItem(phone or "-")
                i4.setTextAlignment(Qt.AlignCenter)
                i4.setForeground(QColor("#0369a1") if phone else QColor("#94a3b8"))
                tbl.setItem(r_idx, 4, i4)

                # Col 5: National ID
                i5 = QTableWidgetItem(nid or "-")
                i5.setTextAlignment(Qt.AlignCenter)
                tbl.setItem(r_idx, 5, i5)

                # Col 6: Insurance Status
                is_ins = ("مؤمن" in ins_st and "غير" not in ins_st)
                i6 = QTableWidgetItem("🛡️ مؤمن" if is_ins else "⚪ غير مؤمن")
                i6.setTextAlignment(Qt.AlignCenter)
                i6.setFont(QFont("Segoe UI", 9, QFont.Bold))
                i6.setForeground(QColor("#15803d" if is_ins else "#64748b"))
                tbl.setItem(r_idx, 6, i6)

                # Col 7: Base Salary
                i7 = QTableWidgetItem(f"{sal:,.2f} ج.م")
                i7.setTextAlignment(Qt.AlignCenter)
                i7.setFont(QFont("Segoe UI", 10, QFont.Bold))
                i7.setForeground(QColor("#059669"))
                tbl.setItem(r_idx, 7, i7)

                # Col 8: Incentive
                i8 = QTableWidgetItem(f"{inc:,.2f} ج.م")
                i8.setTextAlignment(Qt.AlignCenter)
                i8.setFont(QFont("Segoe UI", 10, QFont.Bold))
                i8.setForeground(QColor("#2563eb"))
                tbl.setItem(r_idx, 8, i8)

                # Col 9: Status
                st_txt = "نشط ✅" if active == 1 else "معطل ❌"
                st_color = "#16a34a" if active == 1 else "#dc2626"
                i9 = QTableWidgetItem(st_txt)
                i9.setTextAlignment(Qt.AlignCenter)
                i9.setFont(QFont("Segoe UI", 10, QFont.Bold))
                i9.setForeground(QColor(st_color))
                tbl.setItem(r_idx, 9, i9)

                # Col 10: View Profile Button
                btn_view = QPushButton("👁️ ملف العامل")
                btn_view.setCursor(Qt.PointingHandCursor)
                btn_view.setStyleSheet("background: #0f172a; color: white; border-radius: 5px; padding: 4px 8px; font-weight: bold; border: none; font-size: 11px;")
                btn_view.clicked.connect(lambda _, e_rec=emp: show_employee_profile_dialog(e_rec))
                tbl.setCellWidget(r_idx, 10, btn_view)

                # Col 11: Edit Button
                btn_ed = QPushButton("✏️ تعديل")
                btn_ed.setCursor(Qt.PointingHandCursor)
                btn_ed.setStyleSheet("background: #2563eb; color: white; border-radius: 5px; padding: 4px 10px; font-weight: bold; border: none; font-size: 11px;")
                btn_ed.clicked.connect(lambda _, rec_id=eid: open_emp_form_dialog(rec_id))
                tbl.setCellWidget(r_idx, 11, btn_ed)

                # Col 12: Toggle Active Button
                btn_tog = QPushButton("إيقاف" if active == 1 else "تفعيل")
                btn_tog.setCursor(Qt.PointingHandCursor)
                tog_color = "#dc2626" if active == 1 else "#16a34a"
                btn_tog.setStyleSheet(f"background: {tog_color}; color: white; border-radius: 5px; padding: 4px 10px; font-weight: bold; border: none; font-size: 11px;")
                btn_tog.clicked.connect(lambda _, rec_id=eid, act=active, ename=ename: toggle_emp_active(rec_id, act, ename))
                tbl.setCellWidget(r_idx, 12, btn_tog)

        def toggle_emp_active(rec_id, cur_active, ename):
            new_act = 0 if cur_active == 1 else 1
            action_name = "إيقاف وتجميد" if new_act == 0 else "تفعيل وتشغيل"
            reply = QMessageBox.question(dlg, "تأكيد تغيير الحالة", f"هل تريد بالتأكيد {action_name} ملف العامل <b>{ename}</b>؟", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                conn = get_db_connection()
                c = conn.cursor()
                c.execute("UPDATE employees SET is_active=? WHERE id=?", (new_act, rec_id))
                conn.commit()
                conn.close()
                load_emp_table()
                self.refresh_hr_views()
                if hasattr(self, "show_toast"):
                    self.show_toast(f"✓ تم {action_name} ملف العامل '{ename}' بنجاح")

        def show_employee_profile_dialog(emp_data):
            """عرض ملف بيانات العامل الشامل في بطاقة معلومات متكاملة."""
            (eid, code, ename, dname, sal, inc, hrs, active, did, notes,
             hire_d, birth_d, nid, ins_st, ins_no, marital, phone,
             has_rel, rel_name, rel_rel, addr, dwage) = emp_data

            pd = QDialog(dlg)
            pd.setWindowTitle(f"👤 الملف الشامل للعامل: {ename}")
            pd.setLayoutDirection(Qt.RightToLeft)
            pd.resize(650, 580)
            pd.setStyleSheet("background: #f8fafc; font-family: 'Segoe UI';")

            p_layout = QVBoxLayout(pd)
            p_layout.setSpacing(12)
            p_layout.setContentsMargins(16, 16, 16, 16)

            # Profile Header
            p_hdr = QFrame()
            p_hdr.setStyleSheet("background: #0f172a; border-radius: 8px; padding: 14px;")
            ph_layout = QHBoxLayout(p_hdr)
            lbl_title = QLabel(f"👤 {ename}")
            lbl_title.setFont(QFont("Segoe UI", 14, QFont.Bold))
            lbl_title.setStyleSheet("color: white;")
            ph_layout.addWidget(lbl_title)

            if code:
                lbl_code = QLabel(f"كود: {code}")
                lbl_code.setFont(QFont("Segoe UI", 11, QFont.Bold))
                lbl_code.setStyleSheet("background: #2563eb; color: white; padding: 4px 10px; border-radius: 5px;")
                ph_layout.addWidget(lbl_code)

            st_badge = QLabel("نشط ✅" if active == 1 else "معطل ❌")
            st_badge.setFont(QFont("Segoe UI", 11, QFont.Bold))
            st_badge.setStyleSheet(f"background: {'#16a34a' if active == 1 else '#dc2626'}; color: white; padding: 4px 10px; border-radius: 5px;")
            ph_layout.addWidget(st_badge)
            ph_layout.addStretch()
            p_layout.addWidget(p_hdr)

            def create_info_card(title, items):
                card = QFrame()
                card.setStyleSheet("background: white; border: 1.5px solid #e2e8f0; border-radius: 8px; padding: 10px;")
                c_lay = QVBoxLayout(card)
                c_lay.setSpacing(8)

                c_title = QLabel(title)
                c_title.setFont(QFont("Segoe UI", 11, QFont.Bold))
                c_title.setStyleSheet("color: #1e293b; border-bottom: 1.5px solid #cbd5e1; padding-bottom: 4px;")
                c_lay.addWidget(c_title)

                grid = QGridLayout()
                grid.setSpacing(8)
                for idx, (lbl, val) in enumerate(items):
                    r, col_idx = divmod(idx, 2)
                    lbl_w = QLabel(f"<b>{lbl}:</b>")
                    lbl_w.setStyleSheet("color: #475569; font-size: 12px;")
                    val_w = QLabel(str(val or "—"))
                    val_w.setStyleSheet("color: #0f172a; font-size: 12px; font-weight: bold;")
                    grid.addWidget(lbl_w, r, col_idx * 2)
                    grid.addWidget(val_w, r, col_idx * 2 + 1)
                c_lay.addLayout(grid)
                return card

            # Card 1: Job Info
            p_layout.addWidget(create_info_card("🏢 البيانات الوظيفية والمالية", [
                ("القسم", dname),
                ("تاريخ التعيين", hire_d or "غير محدد"),
                ("الراتب الأساسي", f"{sal:,.2f} ج.م"),
                ("حافز الانتظام", f"{inc:,.2f} ج.م"),
                ("ساعات الوردية", f"{hrs:.1f} ساعة"),
                ("أجر اليوم", f"{dwage:,.2f} ج.م"),
            ]))

            # Card 2: Personal & Contact
            p_layout.addWidget(create_info_card("👤 البيانات الشخصية والاتصال", [
                ("الرقم القومي", nid or "—"),
                ("تاريخ الميلاد", birth_d or "—"),
                ("الحالة الاجتماعية", marital or "—"),
                ("رقم الموبايل", phone or "—"),
                ("العنوان ومحل الإقامة", addr or "—"),
            ]))

            # Card 3: Insurance & Relatives
            rel_info_str = f"{rel_name} ({rel_rel})" if (has_rel == "نعم" and rel_name) else ("نعم" if has_rel == "نعم" else "لا يوجد")
            p_layout.addWidget(create_info_card("🛡️ التأمينات الاجتماعية والأقارب بالشركة", [
                ("الموقف التأميني", ins_st or "غير مؤمن"),
                ("الرقم التأميني", ins_no or "—"),
                ("أقارب بالشركة", rel_info_str),
                ("صلة القرابة", rel_rel or "—"),
                ("ملاحظات إضافية", notes or "لا توجد"),
            ]))

            b_lay = QHBoxLayout()
            btn_close = QPushButton("إغلاق")
            btn_close.setCursor(Qt.PointingHandCursor)
            btn_close.setStyleSheet("background: #64748b; color: white; padding: 7px 24px; font-weight: bold; border-radius: 6px; border: none;")
            btn_close.clicked.connect(pd.accept)
            b_lay.addStretch()
            b_lay.addWidget(btn_close)
            p_layout.addLayout(b_lay)

            pd.exec()

        def open_emp_form_dialog(rec_id=None):
            """فتح نافذة إضافة أو تعديل عامل بكل الحقول المطلوبة."""
            is_edit = (rec_id is not None)
            ed = QDialog(dlg)
            ed.setWindowTitle(f"✏️ تعديل بيانات العامل (ID: {rec_id})" if is_edit else "➕ إضافة عامل جديد بكافة البيانات")
            ed.setLayoutDirection(Qt.RightToLeft)
            ed.resize(750, 560)
            ed.setStyleSheet("background: #f8fafc; font-family: 'Segoe UI';")

            ed_layout = QVBoxLayout(ed)
            ed_layout.setSpacing(12)
            ed_layout.setContentsMargins(16, 16, 16, 16)

            # Dialog Header
            d_hdr = QFrame()
            d_hdr.setStyleSheet(f"background: {'#2563eb' if is_edit else '#16a34a'}; border-radius: 8px; padding: 10px 14px;")
            dh_lay = QHBoxLayout(d_hdr)
            dh_lbl = QLabel(f"✏️ تعديل بيانات العامل" if is_edit else "➕ تسجيل وإضافة عامل جديد")
            dh_lbl.setFont(QFont("Segoe UI", 12, QFont.Bold))
            dh_lbl.setStyleSheet("color: white;")
            dh_lay.addWidget(dh_lbl)
            ed_layout.addWidget(d_hdr)

            # Tabs Container
            ed_tabs = QTabWidget()
            ed_tabs.setStyleSheet("""
                QTabWidget::pane { border: 1.5px solid #cbd5e1; border-radius: 8px; background: white; padding: 12px; }
                QTabBar::tab { background: #e2e8f0; color: #334155; padding: 8px 16px; font-weight: bold; font-size: 12px; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-left: 4px; }
                QTabBar::tab:selected { background: #0f172a; color: white; }
            """)

            # ── Tab 1: البيانات الوظيفية والأساسية ──────────────────────
            t1 = QWidget()
            t1_lay = QFormLayout(t1)
            t1_lay.setSpacing(10)
            t1_lay.setContentsMargins(8, 8, 8, 8)

            name_in = QLineEdit()
            name_in.setPlaceholderText("أدخل اسم العامل الثلاثي أو الرباعي...")
            name_in.setStyleSheet("padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px; font-weight: bold;")
            t1_lay.addRow("اسم العامل *:", name_in)

            code_in = QLineEdit()
            code_in.setPlaceholderText("رقم الكود الوظيفي (مثال: 101)...")
            code_in.setStyleSheet("padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t1_lay.addRow("رقم الكود:", code_in)

            dept_sel = QComboBox()
            dept_sel.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            self.populate_dept_combo(dept_sel)
            t1_lay.addRow("🏢 القسم:", dept_sel)

            hire_date_in = QDateEdit()
            hire_date_in.setCalendarPopup(True)
            hire_date_in.setDisplayFormat("yyyy-MM-dd")
            hire_date_in.setDate(QDate.currentDate())
            hire_date_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t1_lay.addRow("📅 تاريخ التعيين:", hire_date_in)

            sal_in = QDoubleSpinBox()
            sal_in.setRange(0, 1000000)
            sal_in.setDecimals(2)
            sal_in.setValue(4000.0)
            sal_in.setSuffix(" ج.م")
            sal_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px; font-weight: bold; color: #15803d;")
            t1_lay.addRow("💵 الراتب الأساسي:", sal_in)

            inc_in = QDoubleSpinBox()
            inc_in.setRange(0, 100000)
            inc_in.setDecimals(2)
            inc_in.setValue(500.0)
            inc_in.setSuffix(" ج.م")
            inc_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px; font-weight: bold; color: #2563eb;")
            t1_lay.addRow("🎁 حافز الانتظام:", inc_in)

            hrs_in = QDoubleSpinBox()
            hrs_in.setRange(1, 24)
            hrs_in.setDecimals(1)
            hrs_in.setValue(11.0)
            hrs_in.setSuffix(" ساعة")
            hrs_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t1_lay.addRow("⏱️ ساعات الوردية:", hrs_in)

            # Auto update hours from dept
            def update_dept_hours(idx):
                did = dept_sel.currentData()
                if did:
                    d_info = get_department_by_id(did)
                    if d_info:
                        hrs_in.setValue(float(d_info.get("default_work_hours", 11.0)))
            dept_sel.currentIndexChanged.connect(update_dept_hours)

            ed_tabs.addTab(t1, "🏢 البيانات الوظيفية والمالية")

            # ── Tab 2: البيانات الشخصية والاتصال ────────────────────────
            t2 = QWidget()
            t2_lay = QFormLayout(t2)
            t2_lay.setSpacing(10)
            t2_lay.setContentsMargins(8, 8, 8, 8)

            nid_in = QLineEdit()
            nid_in.setPlaceholderText("الرقم القومي المكون من 14 رقم...")
            nid_in.setMaxLength(14)
            nid_in.setStyleSheet("padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t2_lay.addRow("🪪 الرقم القومي:", nid_in)

            birth_date_in = QDateEdit()
            birth_date_in.setCalendarPopup(True)
            birth_date_in.setDisplayFormat("yyyy-MM-dd")
            birth_date_in.setDate(QDate(1998, 1, 1))
            birth_date_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t2_lay.addRow("🎂 تاريخ الميلاد:", birth_date_in)

            marital_sel = QComboBox()
            marital_sel.addItems(["أعزب", "متزوج", "متزوج ويعول", "مطلق", "أرمل"])
            marital_sel.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t2_lay.addRow("💍 الحالة الاجتماعية:", marital_sel)

            phone_in = QLineEdit()
            phone_in.setPlaceholderText("رقم الموبايل (مثال: 01012345678)...")
            phone_in.setStyleSheet("padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t2_lay.addRow("📱 رقم الموبايل:", phone_in)

            addr_in = QLineEdit()
            addr_in.setPlaceholderText("العنوان ومحل الإقامة بالتفصيل...")
            addr_in.setStyleSheet("padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t2_lay.addRow("🏠 العنوان ومحل الإقامة:", addr_in)

            ed_tabs.addTab(t2, "👤 البيانات الشخصية والاتصال")

            # ── Tab 3: التأمينات والأقارب والملاحظات ────────────────────
            t3 = QWidget()
            t3_lay = QFormLayout(t3)
            t3_lay.setSpacing(10)
            t3_lay.setContentsMargins(8, 8, 8, 8)

            ins_status_sel = QComboBox()
            ins_status_sel.addItems(["غير مؤمن", "مؤمن عليه"])
            ins_status_sel.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t3_lay.addRow("🛡️ الموقف التأميني:", ins_status_sel)

            ins_no_in = QLineEdit()
            ins_no_in.setPlaceholderText("الرقم التأميني إن وُجد...")
            ins_no_in.setStyleSheet("padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t3_lay.addRow("📄 الرقم التأميني:", ins_no_in)

            relatives_sel = QComboBox()
            relatives_sel.addItems(["لا", "نعم"])
            relatives_sel.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t3_lay.addRow("👥 هل يوجد أقارب بالشركة؟:", relatives_sel)

            rel_name_in = QLineEdit()
            rel_name_in.setPlaceholderText("اسم القريب العامل بالشركة...")
            rel_name_in.setEnabled(False)
            rel_name_in.setStyleSheet("padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t3_lay.addRow("اسم القريب بالشركة:", rel_name_in)

            rel_relation_in = QComboBox()
            rel_relation_in.setEditable(True)
            rel_relation_in.addItems(["—", "أخ", "أخت", "والد", "والدة", "ابن", "ابنة", "عم", "خال", "ابن عم", "ابن خال", "صهر", "أخرى"])
            rel_relation_in.setEnabled(False)
            rel_relation_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t3_lay.addRow("صلة القرابة:", rel_relation_in)

            def on_relatives_toggle(idx):
                has_r = (relatives_sel.currentText() == "نعم")
                rel_name_in.setEnabled(has_r)
                rel_relation_in.setEnabled(has_r)
                if not has_r:
                    rel_name_in.clear()
                    rel_relation_in.setCurrentIndex(0)
            relatives_sel.currentIndexChanged.connect(on_relatives_toggle)

            notes_in = QLineEdit()
            notes_in.setPlaceholderText("ملاحظات إضافية حول العامل...")
            notes_in.setStyleSheet("padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; font-size: 12px;")
            t3_lay.addRow("📝 ملاحظات عامة:", notes_in)

            ed_tabs.addTab(t3, "🛡️ التأمينات والأقارب والملاحظات")
            ed_layout.addWidget(ed_tabs)

            # Pre-fill data if editing
            if is_edit:
                conn = get_db_connection()
                c = conn.cursor()
                c.execute("""
                    SELECT name, code, department_id, base_salary, attendance_incentive,
                           work_hours, hire_date, birth_date, national_id, insurance_status,
                           insurance_no, marital_status, phone, has_relatives, relative_name,
                           relative_relation, address, notes
                    FROM employees WHERE id=?
                """, (rec_id,))
                row = c.fetchone()
                conn.close()
                if row:
                    (r_name, r_code, r_did, r_sal, r_inc, r_hrs, r_hire, r_birth,
                     r_nid, r_ins_st, r_ins_no, r_marital, r_phone, r_has_rel,
                     r_rel_name, r_rel_rel, r_addr, r_notes) = row

                    name_in.setText(r_name or "")
                    code_in.setText(r_code or "")
                    if r_did:
                        d_idx = dept_sel.findData(r_did)
                        if d_idx >= 0: dept_sel.setCurrentIndex(d_idx)
                    sal_in.setValue(float(r_sal or 0.0))
                    inc_in.setValue(float(r_inc or 500.0))
                    hrs_in.setValue(float(r_hrs or 11.0))
                    if r_hire:
                        qd_hire = QDate.fromString(r_hire, "yyyy-MM-dd")
                        if qd_hire.isValid(): hire_date_in.setDate(qd_hire)
                    if r_birth:
                        qd_birth = QDate.fromString(r_birth, "yyyy-MM-dd")
                        if qd_birth.isValid(): birth_date_in.setDate(qd_birth)

                    nid_in.setText(r_nid or "")
                    if r_marital:
                        m_idx = marital_sel.findText(r_marital)
                        if m_idx >= 0: marital_sel.setCurrentIndex(m_idx)
                    phone_in.setText(r_phone or "")
                    addr_in.setText(r_addr or "")

                    if r_ins_st:
                        ins_idx = ins_status_sel.findText(r_ins_st)
                        if ins_idx >= 0: ins_status_sel.setCurrentIndex(ins_idx)
                    ins_no_in.setText(r_ins_no or "")

                    if r_has_rel:
                        rel_idx = relatives_sel.findText(r_has_rel)
                        if rel_idx >= 0: relatives_sel.setCurrentIndex(rel_idx)
                    on_relatives_toggle(0)
                    rel_name_in.setText(r_rel_name or "")
                    if r_rel_rel:
                        rel_rel_idx = rel_relation_in.findText(r_rel_rel)
                        if rel_rel_idx >= 0:
                            rel_relation_in.setCurrentIndex(rel_rel_idx)
                        else:
                            rel_relation_in.setEditText(r_rel_rel)
                    notes_in.setText(r_notes or "")

            # Bottom Actions
            b_layout = QHBoxLayout()
            save_btn = QPushButton("💾 حفظ بيانات العامل" if not is_edit else "💾 حفظ التعديلات")
            save_btn.setCursor(Qt.PointingHandCursor)
            save_btn.setStyleSheet(f"background: {'#2563eb' if is_edit else '#16a34a'}; color: white; padding: 9px 24px; font-weight: bold; border-radius: 6px; border: none; font-size: 13px;")
            canc_btn = QPushButton("إلغاء")
            canc_btn.setCursor(Qt.PointingHandCursor)
            canc_btn.setStyleSheet("background: #64748b; color: white; padding: 9px 20px; border-radius: 6px; border: none; font-size: 13px;")
            b_layout.addWidget(save_btn)
            b_layout.addWidget(canc_btn)
            ed_layout.addLayout(b_layout)

            canc_btn.clicked.connect(ed.reject)

            def do_save():
                nm = name_in.text().strip()
                if not nm:
                    QMessageBox.warning(ed, "تنبيه", "يرجى كتابة اسم العامل أولاً.")
                    ed_tabs.setCurrentIndex(0)
                    name_in.setFocus()
                    return

                cd = code_in.text().strip()
                did = dept_sel.currentData()
                bsal = sal_in.value()
                ainc = inc_in.value()
                whrs = hrs_in.value()
                dwage = round(bsal / 30.0, 2)
                h_date = hire_date_in.date().toString("yyyy-MM-dd")

                nid_val = nid_in.text().strip()
                b_date = birth_date_in.date().toString("yyyy-MM-dd")
                marital_val = marital_sel.currentText().strip()
                phone_val = phone_in.text().strip()
                addr_val = addr_in.text().strip()

                ins_st_val = ins_status_sel.currentText().strip()
                ins_no_val = ins_no_in.text().strip()
                has_rel_val = relatives_sel.currentText().strip()
                rel_nm_val = rel_name_in.text().strip() if has_rel_val == "نعم" else ""
                rel_rel_val = rel_relation_in.currentText().strip() if has_rel_val == "نعم" else ""
                nts_val = notes_in.text().strip()

                conn = get_db_connection()
                c = conn.cursor()

                if not is_edit:
                    c.execute("""
                        INSERT INTO employees (
                            name, code, department_id, base_salary, attendance_incentive,
                            daily_wage, work_hours, hire_date, birth_date, national_id,
                            insurance_status, insurance_no, marital_status, phone,
                            has_relatives, relative_name, relative_relation, address,
                            notes, is_active
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
                    """, (
                        nm, cd, did, bsal, ainc, dwage, whrs, h_date, b_date, nid_val,
                        ins_st_val, ins_no_val, marital_val, phone_val,
                        has_rel_val, rel_nm_val, rel_rel_val, addr_val, nts_val
                    ))
                    conn.commit()
                    conn.close()
                    if hasattr(self, "show_toast"):
                        self.show_toast(f"✓ تم إضافة العامل '{nm}' بنجاح")
                else:
                    c.execute("""
                        UPDATE employees
                        SET name=?, code=?, department_id=?, base_salary=?, attendance_incentive=?,
                            daily_wage=?, work_hours=?, hire_date=?, birth_date=?, national_id=?,
                            insurance_status=?, insurance_no=?, marital_status=?, phone=?,
                            has_relatives=?, relative_name=?, relative_relation=?, address=?,
                            notes=?
                        WHERE id=?
                    """, (
                        nm, cd, did, bsal, ainc, dwage, whrs, h_date, b_date, nid_val,
                        ins_st_val, ins_no_val, marital_val, phone_val,
                        has_rel_val, rel_nm_val, rel_rel_val, addr_val, nts_val, rec_id
                    ))
                    conn.commit()
                    conn.close()
                    if hasattr(self, "show_toast"):
                        self.show_toast(f"✓ تم حفظ تعديلات العامل '{nm}' بنجاح")

                ed.accept()
                load_emp_table()
                self.refresh_hr_views()

            save_btn.clicked.connect(do_save)
            ed.exec()

        def export_employees_to_excel():
            """تصدير قاعدة بيانات العاملين الشاملة بكافة الحقول إلى ملف Excel."""
            if not raw_employees_cache:
                QMessageBox.warning(dlg, "تنبيه", "لا توجد بيانات عمال للتصدير.")
                return

            default_name = f"كشف_بيانات_العاملين_الشامل_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            file_path, _ = QFileDialog.getSaveFileName(dlg, "تصدير كشف العاملين إلى Excel", default_name, "Excel Files (*.xlsx)")
            if not file_path:
                return

            export_headers = [
                "ID", "كود العامل", "اسم العامل", "القسم", "الموبايل",
                "الرقم القومي", "تاريخ الميلاد", "الحالة الاجتماعية", "العنوان ومحل الإقامة",
                "تاريخ التعيين", "الموقف التأميني", "الرقم التأميني", "أقارب بالشركة",
                "اسم القريب", "صلة القرابة", "الراتب الأساسي (ج.م)", "حافز الانتظام (ج.م)",
                "أجر اليوم (ج.م)", "ساعات الوردية", "حالة الملف", "ملاحظات"
            ]

            rows_to_export = []
            for emp in raw_employees_cache:
                (eid, code, ename, dname, sal, inc, hrs, active, did, notes,
                 hire_d, birth_d, nid, ins_st, ins_no, marital, phone,
                 has_rel, rel_name, rel_rel, addr, dwage) = emp

                rows_to_export.append([
                    eid, code or "—", ename, dname or "عام", phone or "—",
                    nid or "—", birth_d or "—", marital or "—", addr or "—",
                    hire_d or "—", ins_st or "غير مؤمن", ins_no or "—", has_rel or "لا",
                    rel_name or "—", rel_rel or "—", sal, inc, dwage, hrs,
                    "نشط" if active == 1 else "معطل", notes or ""
                ])

            try:
                self._write_styled_xlsx(
                    file_path,
                    "السجل العام وبيانات العاملين والموظفين",
                    export_headers,
                    rows_to_export,
                    sheet_name="بيانات العاملين",
                    extra_info=f"إجمالي عدد العاملين: {len(rows_to_export)} | تاريخ التصدير: {datetime.now().strftime('%Y-%m-%d %I:%M %p')}"
                )
                QMessageBox.information(dlg, "نجاح التصدير", f"تم تصدير كشف العاملين الشامل بنجاح إلى:\n{file_path}")
            except Exception as ex:
                QMessageBox.critical(dlg, "خطأ في التصدير", f"تعذر تصدير الملف: {ex}")

        search_inp.textChanged.connect(load_emp_table)
        dept_cb.currentIndexChanged.connect(load_emp_table)
        status_cb.currentIndexChanged.connect(load_emp_table)
        ins_cb.currentIndexChanged.connect(load_emp_table)
        btn_add_emp.clicked.connect(lambda: open_emp_form_dialog(None))
        btn_export_emps.clicked.connect(export_employees_to_excel)

        load_emp_table()
        dlg.exec()

    def open_monthly_calendar(self):
        """عرض التقويم الشهري التفاعلي للحضور والسهرات."""
        dialog = QDialog(self)
        dialog.setWindowTitle("📅 التقويم الشهري للحضور والسهرات لجميع العاملين")
        dialog.setLayoutDirection(Qt.RightToLeft)
        dialog.resize(1150, 680)
        dialog.setStyleSheet("background-color: #f8fafc; font-family: 'Segoe UI';")
        d_layout = QVBoxLayout(dialog)
        d_layout.setSpacing(10)
        d_layout.setContentsMargins(14, 14, 14, 14)

        bar = QFrame()
        bar.setStyleSheet("background: white; border-radius: 8px; padding: 8px; border: 1.5px solid #cbd5e1;")
        bar_l = QHBoxLayout(bar)
        bar_l.setSpacing(10)

        bar_l.addWidget(QLabel("الشهر:"))
        m_combo = QComboBox()
        for m_num in range(1, 13):
            m_combo.addItem(f"شهر {m_num}", m_num)
        m_combo.setCurrentIndex(datetime.now().month - 1)
        m_combo.setStyleSheet("padding: 5px; font-weight: bold;")
        bar_l.addWidget(m_combo)

        bar_l.addWidget(QLabel("السنة:"))
        y_spin = QSpinBox()
        y_spin.setRange(2020, 2035)
        y_spin.setValue(datetime.now().year)
        y_spin.setStyleSheet("padding: 5px; font-weight: bold;")
        bar_l.addWidget(y_spin)

        btn_load_cal = QPushButton("🔄 عرض التقويم")
        btn_load_cal.setCursor(Qt.PointingHandCursor)
        btn_load_cal.setStyleSheet("background: #2563eb; color: white; padding: 6px 18px; font-weight: bold; border-radius: 6px;")
        bar_l.addWidget(btn_load_cal)
        bar_l.addStretch()
        d_layout.addWidget(bar)

        cal_table = QTableWidget()
        cal_table.setStyleSheet("""
            QTableWidget { background: white; border: 1.5px solid #cbd5e1; border-radius: 8px; }
            QTableWidget::item { padding: 4px; font-size: 11px; }
            QHeaderView::section { background-color: #0f172a; color: white; padding: 6px; font-weight: bold; }
        """)
        d_layout.addWidget(cal_table)

        def load_cal_data():
            sel_m = m_combo.currentData() or (m_combo.currentIndex() + 1)
            sel_y = y_spin.value()
            days_in_m = calendar.monthrange(sel_y, sel_m)[1]

            conn = get_db_connection()
            c = conn.cursor()
            c.execute("SELECT id, name FROM employees WHERE is_active=1 ORDER BY name ASC")
            emps = c.fetchall()

            cal_table.setColumnCount(days_in_m + 1)
            headers = ["اسم العامل"] + [str(d) for d in range(1, days_in_m + 1)]
            cal_table.setHorizontalHeaderLabels(headers)
            cal_table.setRowCount(len(emps))

            for r_idx, (eid, ename) in enumerate(emps):
                n_item = QTableWidgetItem(ename)
                n_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                cal_table.setItem(r_idx, 0, n_item)

                for day in range(1, days_in_m + 1):
                    dt_str = f"{sel_y:04d}-{sel_m:02d}-{day:02d}"
                    qd = QDate(sel_y, sel_m, day)
                    if qd.dayOfWeek() == 5:
                        item = QTableWidgetItem("🕌")
                        item.setBackground(QColor("#f3e8ff"))
                    else:
                        c.execute("SELECT status, work_hours FROM attendance WHERE employee_id=? AND date=?", (eid, dt_str))
                        att = c.fetchone()
                        if att:
                            st = att[0]
                            if st == "حضور":
                                hrs = att[1] or 0.0
                                item = QTableWidgetItem(f"✅ {hrs:.0f}س" if hrs > 0 else "✅")
                                item.setBackground(QColor("#dcfce7"))
                                item.setForeground(QColor("#15803d"))
                            else:
                                item = QTableWidgetItem("❌")
                                item.setBackground(QColor("#fee2e2"))
                                item.setForeground(QColor("#b91c1c"))
                        else:
                            item = QTableWidgetItem("-")
                    item.setTextAlignment(Qt.AlignCenter)
                    cal_table.setItem(r_idx, day, item)

            conn.close()
            cal_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            for i in range(1, days_in_m + 1):
                cal_table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeToContents)

        btn_load_cal.clicked.connect(load_cal_data)
        load_cal_data()
        dialog.exec()

    def upload_hr_to_sheets_silent(self, date_str=None):
        pass

    def setup_attendance_summary_tab(self, tab):
        layout = QVBoxLayout(tab)
        layout.setSpacing(6)
        layout.setContentsMargins(6, 6, 6, 6)

        bar = QFrame()
        bar.setStyleSheet("background: #f8fafc; border-radius: 8px; border: 1px solid #e2e8f0;")
        bar_layout = QHBoxLayout(bar)
        bar_layout.setSpacing(8)
        bar_layout.setContentsMargins(8, 5, 8, 5)

        bar_layout.addWidget(QLabel("📅 من:"))
        self.att_summary_from = QDateEdit()
        self.att_summary_from.setDate(QDate.currentDate().addDays(-30))
        self.att_summary_from.setDisplayFormat("yyyy-MM-dd")
        self.att_summary_from.setCalendarPopup(True)
        self.att_summary_from.setStyleSheet("font-size: 12px; padding: 5px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;")
        bar_layout.addWidget(self.att_summary_from)

        bar_layout.addWidget(QLabel("إلى:"))
        self.att_summary_to = QDateEdit()
        self.att_summary_to.setDate(QDate.currentDate())
        self.att_summary_to.setDisplayFormat("yyyy-MM-dd")
        self.att_summary_to.setCalendarPopup(True)
        self.att_summary_to.setStyleSheet("font-size: 12px; padding: 5px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;")
        bar_layout.addWidget(self.att_summary_to)

        btn_this_month = QPushButton("هذا الشهر")
        btn_this_month.setCursor(Qt.PointingHandCursor)
        btn_this_month.setStyleSheet("background: #e2e8f0; color: #334155; padding: 5px 10px; font-weight: bold; border-radius: 4px; border: none; font-size: 11px;")
        btn_this_month.clicked.connect(self.set_summary_date_this_month)
        bar_layout.addWidget(btn_this_month)

        calc_btn = QPushButton("📊 حساب")
        calc_btn.setCursor(Qt.PointingHandCursor)
        calc_btn.setStyleSheet("""
            QPushButton { background: #0f766e; color: white; padding: 5px 14px; font-weight: bold; border-radius: 6px; border: none; font-size: 12px; }
            QPushButton:hover { background: #115e59; }
        """)
        calc_btn.clicked.connect(self.calculate_attendance_summary)
        bar_layout.addWidget(calc_btn)

        exp_btn = QPushButton("📥 Excel")
        exp_btn.setCursor(Qt.PointingHandCursor)
        exp_btn.setStyleSheet("""
            QPushButton { background: #16a34a; color: white; padding: 5px 14px; font-weight: bold; border-radius: 6px; border: none; font-size: 12px; }
            QPushButton:hover { background: #15803d; }
        """)
        exp_btn.clicked.connect(self.export_attendance_summary)
        bar_layout.addWidget(exp_btn)

        cal_btn = QPushButton("📅 التقويم الشهري")
        cal_btn.setCursor(Qt.PointingHandCursor)
        cal_btn.setStyleSheet("""
            QPushButton { background: #7c3aed; color: white; padding: 5px 14px; font-weight: bold; border-radius: 6px; border: none; font-size: 12px; }
            QPushButton:hover { background: #6d28d9; }
        """)
        cal_btn.clicked.connect(self.open_monthly_calendar)
        bar_layout.addWidget(cal_btn)

        perf_btn = QPushButton("👤 كشف أداء وإنتاجية العامل")
        perf_btn.setCursor(Qt.PointingHandCursor)
        perf_btn.setStyleSheet("""
            QPushButton { background: #2563eb; color: white; padding: 5px 14px; font-weight: bold; border-radius: 6px; border: none; font-size: 12px; }
            QPushButton:hover { background: #1d4ed8; }
        """)
        perf_btn.clicked.connect(lambda: self.open_employee_monthly_performance_dialog())
        bar_layout.addWidget(perf_btn)

        bar_layout.addStretch()
        layout.addWidget(bar)

        # Results Table with hours breakdown
        self.att_summary_table = QTableWidget()
        self.att_summary_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.att_summary_table.setColumnCount(8)
        self.att_summary_table.setHorizontalHeaderLabels([
            "اسم العامل", "أيام الحضور", "إجمالي الأيام",
            "نسبة الحضور", "ساعات العمل الأساسية", "ساعات السهرات", "إجمالي الخصومات", "إجمالي الحوافز"
        ])
        self.att_summary_table.setAlternatingRowColors(True)
        self.att_summary_table.doubleClicked.connect(self._on_summary_table_double_clicked)
        self.att_summary_table.setStyleSheet("""
            QTableWidget {
                border: 1.5px solid #cbd5e1; border-radius: 8px;
                background: #ffffff; gridline-color: #e2e8f0;
                alternate-background-color: #f8fafc;
                selection-background-color: #dbeafe; selection-color: #1e3a8a;
            }
            QTableWidget::item { padding: 8px 12px; font-size: 13px; color: #0f172a; border-bottom: 1px solid #e2e8f0; }
            QTableWidget::item:selected { background: #dbeafe; color: #1e3a8a; font-weight: bold; }
            QHeaderView::section {
                background: #0f172a; color: #ffffff;
                padding: 10px 10px; font-weight: bold; font-size: 13px;
                border: 1px solid #1e293b; border-bottom: 3px solid #2563eb;
            }
        """)
        hdr = self.att_summary_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        self.att_summary_table.verticalHeader().setDefaultSectionSize(44)
        layout.addWidget(self.att_summary_table)

        self.att_summary_info = QLabel("")
        self.att_summary_info.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.att_summary_info.setStyleSheet("color: #1e40af; padding: 6px 14px; background: #eff6ff; border-radius: 6px; border: 1px solid #bfdbfe;")
        layout.addWidget(self.att_summary_info)

    def set_summary_date_this_month(self):
        now = QDate.currentDate()
        first_day = QDate(now.year(), now.month(), 1)
        self.att_summary_from.setDate(first_day)
        self.att_summary_to.setDate(now)
        self.calculate_attendance_summary()

    def calculate_attendance_summary(self):
        from_date = self.att_summary_from.date().toString("yyyy-MM-dd")
        to_date = self.att_summary_to.date().toString("yyyy-MM-dd")
        
        if from_date > to_date:
            QMessageBox.warning(self, "تنبيه", "تاريخ البداية يجب أن يكون قبل تاريخ النهاية.")
            return
        
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        
        c.execute("SELECT id, name FROM employees ORDER BY id ASC")
        employees = c.fetchall()
        
        from_date_obj = datetime.strptime(from_date, "%Y-%m-%d")
        to_date_obj = datetime.strptime(to_date, "%Y-%m-%d")
        
        total_work_days = 0
        current_date = from_date_obj
        while current_date <= to_date_obj:
            if current_date.weekday() != 4:  # Not Friday
                total_work_days += 1
            current_date += timedelta(days=1)
        
        if not hasattr(self, "att_summary_table"):
            conn.close()
            return

        with self.preserve_table_scroll(self.att_summary_table):
            self.att_summary_table.setUpdatesEnabled(False)
            self.att_summary_table.setRowCount(0)
            
            for emp_id, emp_name in employees:
                c.execute(
                    """SELECT COUNT(*), COALESCE(SUM(work_hours), 0) FROM attendance
                       WHERE employee_id=? AND date BETWEEN ? AND ? AND status='حضور'""",
                    (emp_id, from_date, to_date)
                )
                att_row = c.fetchone()
                present_days = att_row[0] or 0
                reg_hours = att_row[1] or 0.0

                c.execute(
                    """SELECT COALESCE(SUM(amount), 0) FROM deductions
                       WHERE employee_id=? AND date BETWEEN ? AND ?""",
                    (emp_id, from_date, to_date)
                )
                total_deductions = c.fetchone()[0] or 0

                c.execute(
                    """SELECT COALESCE(SUM(amount), 0) FROM incentives
                       WHERE employee_id=? AND date BETWEEN ? AND ?""",
                    (emp_id, from_date, to_date)
                )
                total_incentives = c.fetchone()[0] or 0

                # Calculate overtime hours
                c.execute(
                    """SELECT start_time, end_time FROM overtime
                       WHERE employee_id=? AND date BETWEEN ? AND ?""",
                    (emp_id, from_date, to_date)
                )
                overtime_records = c.fetchall()
                total_overtime_hours = 0
                for start, end in overtime_records:
                    t_start = parse_time_str(start)
                    t_end = parse_time_str(end)
                    if t_start.isValid() and t_end.isValid():
                        duration_minutes = (t_end.hour() * 60 + t_end.minute()) - (t_start.hour() * 60 + t_start.minute())
                        if duration_minutes < 0:
                            duration_minutes += 24 * 60
                    else:
                        duration_minutes = 0
                    total_overtime_hours += duration_minutes / 60

                attendance_percentage = (present_days / total_work_days * 100) if total_work_days > 0 else 0

                row_idx = self.att_summary_table.rowCount()
                self.att_summary_table.insertRow(row_idx)

                name_item = QTableWidgetItem(emp_name)
                name_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                self.att_summary_table.setItem(row_idx, 0, name_item)

                self.att_summary_table.setItem(row_idx, 1, QTableWidgetItem(str(present_days)))
                self.att_summary_table.setItem(row_idx, 2, QTableWidgetItem(str(total_work_days)))

                pct_item = QTableWidgetItem(f"{attendance_percentage:.1f}%")
                pct_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                if attendance_percentage >= 90:
                    pct_item.setForeground(QColor("#15803d"))
                elif attendance_percentage >= 75:
                    pct_item.setForeground(QColor("#d97706"))
                else:
                    pct_item.setForeground(QColor("#dc2626"))
                self.att_summary_table.setItem(row_idx, 3, pct_item)

                reg_item = QTableWidgetItem(f"{reg_hours:.1f} س")
                reg_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                reg_item.setForeground(QColor("#0284c7"))
                self.att_summary_table.setItem(row_idx, 4, reg_item)

                ot_item = QTableWidgetItem(f"{total_overtime_hours:.1f} س" if total_overtime_hours > 0 else "0")
                if total_overtime_hours > 0:
                    ot_item.setForeground(QColor("#7c3aed"))
                    ot_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                self.att_summary_table.setItem(row_idx, 5, ot_item)

                ded_item = QTableWidgetItem(f"{total_deductions:,.2f} ج.م" if total_deductions > 0 else "0")
                if total_deductions > 0:
                    ded_item.setForeground(QColor("#dc2626"))
                self.att_summary_table.setItem(row_idx, 6, ded_item)

                inc_item = QTableWidgetItem(f"{total_incentives:,.2f} ج.م" if total_incentives > 0 else "0")
                if total_incentives > 0:
                    inc_item.setForeground(QColor("#15803d"))
                self.att_summary_table.setItem(row_idx, 7, inc_item)

            self.att_summary_table.setUpdatesEnabled(True)
        
        conn.close()
        
        self.att_summary_info.setText(
            f"الفترة: من {from_date} إلى {to_date} | إجمالي أيام العمل (بدون جمعة): {total_work_days} يوم | عدد العاملين: {len(employees)}"
        )

    def export_attendance_summary(self):
        if self.att_summary_table.rowCount() == 0:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير. قم بحساب الحضور أولاً.")
            return

        from_date = self.att_summary_from.date().toString("yyyy-MM-dd")
        to_date = self.att_summary_to.date().toString("yyyy-MM-dd")
        default_name = f"attendance_summary_{from_date}_to_{to_date}.xlsx"

        file_path, _ = QFileDialog.getSaveFileName(self, "تصدير تقرير الحضور", default_name, "Excel Files (*.xlsx)")
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        headers = [
            "اسم العامل", "أيام الحضور", "إجمالي الأيام",
            "نسبة الحضور", "ساعات العمل الأساسية", "ساعات السهرات", "إجمالي الخصومات", "إجمالي الحوافز"
        ]
        data_rows = []
        for row_idx in range(self.att_summary_table.rowCount()):
            r = []
            for col_idx in range(self.att_summary_table.columnCount()):
                item = self.att_summary_table.item(row_idx, col_idx)
                r.append(item.text() if item else "")
            data_rows.append(r)

        summary_text = self.att_summary_info.text()
        self._write_styled_xlsx(
            file_path, f"تقرير الحضور والغياب وساعات العمل  {from_date} ← {to_date}",
            headers, data_rows, "تجميع الحضور",
            extra_info=summary_text
        )
        QMessageBox.information(self, "تصدير", "تم تصدير تقرير الحضور بنجاح ✔️")

    def setup_overtime_tab(self, tab):
        layout = QVBoxLayout(tab)
        layout.setSpacing(6)
        layout.setContentsMargins(6, 6, 6, 6)

        top_bar = QFrame()
        top_bar.setStyleSheet("background: #f8fafc; border-radius: 8px; border: 1px solid #e2e8f0;")
        top_layout = QHBoxLayout(top_bar)
        top_layout.setSpacing(8)
        top_layout.setContentsMargins(8, 5, 8, 5)

        top_layout.addWidget(QLabel("📅"))
        self.ot_date = QDateEdit()
        self.ot_date.setDate(QDate.currentDate())
        self.ot_date.setDisplayFormat("yyyy-MM-dd")
        self.ot_date.setCalendarPopup(True)
        self.ot_date.setStyleSheet("font-size: 12px; padding: 5px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;")
        top_layout.addWidget(self.ot_date)

        top_layout.addWidget(QLabel("👤"))
        self.ot_emp_combo = QComboBox()
        self.ot_emp_combo.setStyleSheet("font-size: 12px; padding: 5px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; min-width: 150px; color: #0f172a;")
        self.populate_employee_combo(self.ot_emp_combo)
        self.hr_employee_combos.append(self.ot_emp_combo)
        top_layout.addWidget(self.ot_emp_combo)

        top_layout.addWidget(QLabel("⏰ من:"))
        self.ot_start_time = QTimeEdit()
        self.ot_start_time.setLocale(ARABIC_LOCALE)
        self.ot_start_time.setTime(QTime(18, 0))
        self.ot_start_time.setDisplayFormat("hh:mm AP")
        self.ot_start_time.setStyleSheet("font-size: 12px; padding: 5px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;")
        top_layout.addWidget(self.ot_start_time)

        top_layout.addWidget(QLabel("إلى:"))
        self.ot_end_time = QTimeEdit()
        self.ot_end_time.setLocale(ARABIC_LOCALE)
        self.ot_end_time.setTime(QTime(22, 0))
        self.ot_end_time.setDisplayFormat("hh:mm AP")
        self.ot_end_time.setStyleSheet("font-size: 12px; padding: 5px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;")
        top_layout.addWidget(self.ot_end_time)

        self.ot_notes = QLineEdit()
        self.ot_notes.setPlaceholderText("ملاحظات...")
        self.ot_notes.setStyleSheet("font-size: 12px; padding: 5px 8px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;")
        top_layout.addWidget(self.ot_notes, 1)

        add_btn = QPushButton("➕ تسجيل السهرة")
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.setStyleSheet("""
            QPushButton {
                background-color: #7c3aed; color: white;
                padding: 6px 14px; font-weight: bold;
                border-radius: 6px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #6d28d9; }
        """)
        add_btn.clicked.connect(self.save_overtime)
        top_layout.addWidget(add_btn)

        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setStyleSheet("color: #cbd5e1;")
        top_layout.addWidget(sep)

        top_layout.addWidget(QLabel("📊 من:"))
        self.ot_from_date = QDateEdit()
        self.ot_from_date.setDate(QDate.currentDate().addDays(-30))
        self.ot_from_date.setDisplayFormat("yyyy-MM-dd")
        self.ot_from_date.setCalendarPopup(True)
        self.ot_from_date.setStyleSheet("font-size: 12px; padding: 5px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;")
        top_layout.addWidget(self.ot_from_date)

        top_layout.addWidget(QLabel("إلى:"))
        self.ot_to_date = QDateEdit()
        self.ot_to_date.setDate(QDate.currentDate())
        self.ot_to_date.setDisplayFormat("yyyy-MM-dd")
        self.ot_to_date.setCalendarPopup(True)
        self.ot_to_date.setStyleSheet("font-size: 12px; padding: 5px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;")
        top_layout.addWidget(self.ot_to_date)

        export_btn = QPushButton("📥 Excel")
        export_btn.setCursor(Qt.PointingHandCursor)
        export_btn.setStyleSheet("""
            QPushButton {
                background-color: #059669; color: white;
                padding: 6px 12px; font-weight: bold;
                border-radius: 6px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #047857; }
        """)
        export_btn.clicked.connect(self.export_overtime_report)
        top_layout.addWidget(export_btn)

        layout.addWidget(top_bar)

        self.ot_table = QTableWidget()
        self.ot_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.prepare_table_with_delete(self.ot_table, ["ID", "التاريخ", "اسم العامل", "وقت البداية", "وقت النهاية", "المدة", "ملاحظات"])

        hdr = self.ot_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(6, QHeaderView.Stretch)
        self.ot_table.verticalHeader().setDefaultSectionSize(36)
        self.ot_table.setAlternatingRowColors(True)
        self.ot_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #e2e8f0; border-radius: 8px;
                background: white; gridline-color: #f1f5f9;
                alternate-background-color: #f8fafc;
            }
            QTableWidget::item { padding: 6px 10px; font-size: 13px; }
            QTableWidget::item:selected { background: #f3e8ff; color: #5b21b6; }
            QHeaderView::section {
                background: #faf5ff; color: #1e293b;
                padding: 8px 12px; font-weight: bold; font-size: 13px;
                border: none; border-bottom: 2px solid #d8b4fe;
            }
        """)
        layout.addWidget(self.ot_table)

        self.ot_summary = QLabel("")
        self.ot_summary.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self.ot_summary.setStyleSheet("color: #7c3aed; padding: 6px 14px; background: #f3e8ff; border-radius: 6px; border: 1px solid #d8b4fe;")
        layout.addWidget(self.ot_summary)

        self.load_overtime()

    def save_overtime(self):
        emp_id = self.ot_emp_combo.currentData()
        if emp_id is None:
            QMessageBox.warning(self, "تنبيه", "اختر عاملاً أولاً.")
            return
        
        if self.ot_start_time.time() >= self.ot_end_time.time():
            QMessageBox.warning(self, "تنبيه", "وقت النهاية يجب أن يكون بعد وقت البداية.")
            return
        
        date_str = self.ot_date.date().toString("yyyy-MM-dd")
        start_time = format_time_ar(self.ot_start_time.time())
        end_time = format_time_ar(self.ot_end_time.time())
        notes = self.ot_notes.text().strip()
        
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute(
            "INSERT INTO overtime (date, employee_id, start_time, end_time, notes) VALUES (?, ?, ?, ?, ?)",
            (date_str, emp_id, start_time, end_time, notes)
        )
        conn.commit()
        conn.close()
        
        self.ot_notes.clear()
        self.load_overtime()
        self.show_toast("تم تسجيل السهرة بنجاح!")
        self.upload_hr_to_sheets_silent(date_str)

    def load_overtime(self):
        if not hasattr(self, "ot_table"):
            return

        with self.preserve_table_scroll(self.ot_table):
            self.ot_table.setUpdatesEnabled(False)
            self.ot_table.setRowCount(0)
            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()
            c.execute(
                """SELECT o.id, o.date, e.name, o.start_time, o.end_time, o.notes 
                   FROM overtime o JOIN employees e ON o.employee_id=e.id 
                   ORDER BY o.date DESC, o.id DESC"""
            )
            records = c.fetchall()
            conn.close()
            
            total_hours = 0
            for row_idx, (ot_id, date, emp_name, start, end, notes) in enumerate(records):
                self.ot_table.insertRow(row_idx)
                self.ot_table.setItem(row_idx, 0, QTableWidgetItem(str(ot_id)))
                self.ot_table.setItem(row_idx, 1, QTableWidgetItem(date))
                self.ot_table.setItem(row_idx, 2, QTableWidgetItem(emp_name))
                self.ot_table.setItem(row_idx, 3, QTableWidgetItem(format_time_ar(start)))
                self.ot_table.setItem(row_idx, 4, QTableWidgetItem(format_time_ar(end)))
                
                t_start = parse_time_str(start)
                t_end = parse_time_str(end)

                if t_start.isValid() and t_end.isValid():
                    duration_minutes = (t_end.hour() * 60 + t_end.minute()) - (t_start.hour() * 60 + t_start.minute())
                    if duration_minutes < 0:
                        duration_minutes += 24 * 60
                else:
                    duration_minutes = 0

                duration_hours = duration_minutes / 60
                total_hours += duration_hours
                
                duration_str = f"{duration_hours:.1f} ساعة"
                self.ot_table.setItem(row_idx, 5, QTableWidgetItem(duration_str))
                self.ot_table.setItem(row_idx, 6, QTableWidgetItem(notes or "-"))
                
                self.add_edit_button(
                    self.ot_table,
                    row_idx,
                    7,
                    "overtime",
                    ot_id,
                    self.load_overtime
                )
                self.add_delete_button(
                    self.ot_table,
                    row_idx,
                    8,
                    "overtime",
                    ot_id,
                    self.load_overtime
                )
            
            self.ot_table.setUpdatesEnabled(True)
        
        self.ot_summary.setText(f"إجمالي السهرات المسجلة: {len(records)} سهرة | إجمالي الساعات: {total_hours:.1f} ساعة")

    def export_overtime_report(self):
        from_date = self.ot_from_date.date().toString("yyyy-MM-dd")
        to_date = self.ot_to_date.date().toString("yyyy-MM-dd")

        if from_date > to_date:
            QMessageBox.warning(self, "تنبيه", "تاريخ البداية يجب أن يكون قبل أو مساوي لتاريخ النهاية.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "حفظ ملف Excel", f"تقرير_السهرات_{from_date}_إلى_{to_date}.xlsx", "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        cursor.execute(
            """SELECT o.id, o.date, e.name, o.start_time, o.end_time, o.notes
               FROM overtime o JOIN employees e ON o.employee_id=e.id
               WHERE o.date BETWEEN ? AND ?
               ORDER BY o.date ASC, o.id ASC""",
            (from_date, to_date)
        )
        records = cursor.fetchall()
        conn.close()

        if not records:
            QMessageBox.information(self, "معلومات", "لا توجد سهرات مسجلة في هذا النطاق.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "تقرير السهرات"

        headers = ["اسم العامل"] + [str(d) for d in range(1, days_in_month + 1)]
        calendar_table.setHorizontalHeaderLabels(headers)

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT id, name FROM employees ORDER BY id ASC")
        employees = c.fetchall()

        calendar_table.setRowCount(len(employees))

        for row_idx, (emp_id, emp_name) in enumerate(employees):
            calendar_table.setItem(row_idx, 0, QTableWidgetItem(emp_name))

            for day in range(1, days_in_month + 1):
                date_str = f"{year:04d}-{month_idx:02d}-{day:02d}"
                qdate = QDate(year, month_idx, day)

                cell_text = ""
                cell_color = "#ffffff"

                if qdate.dayOfWeek() == 5:
                    cell_text = "🕌"
                    cell_color = "#f3e8ff"
                else:
                    c.execute(
                        "SELECT status, work_hours FROM attendance WHERE employee_id=? AND date=?",
                        (emp_id, date_str)
                    )
                    att = c.fetchone()
                    if att:
                        if att[0] == "حضور":
                            h_val = att[1] if att[1] is not None else 8.0
                            cell_text = f"✅\n{h_val:.1f}س"
                            cell_color = "#f0fdf4"
                        else:
                            cell_text = "❌"
                            cell_color = "#fef2f2"

                    c.execute(
                        "SELECT start_time, end_time FROM overtime WHERE employee_id=? AND date=?",
                        (emp_id, date_str)
                    )
                    ot = c.fetchone()
                    if ot:
                        if cell_text:
                            cell_text += "\n🌙"
                        else:
                            cell_text = "🌙"
                        cell_color = "#fffbeb"

                item = QTableWidgetItem(cell_text)
                item.setTextAlignment(Qt.AlignCenter)
                if cell_color != "#ffffff":
                    item.setBackground(QColor(cell_color))
                calendar_table.setItem(row_idx, day, item)

        conn.close()

        header = calendar_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for i in range(1, days_in_month + 1):
            header.setSectionResizeMode(i, QHeaderView.ResizeToContents)

        refresh_btn.clicked.connect(load_calendar)
        load_calendar()

        dialog.exec()

    def _on_summary_table_double_clicked(self, index):
        row = index.row()
        name_item = self.att_summary_table.item(row, 0)
        if name_item:
            emp_name = name_item.text().strip()
            conn = sqlite3.connect(DB_NAME)
            r = conn.execute("SELECT id FROM employees WHERE name=?", (emp_name,)).fetchone()
            conn.close()
            emp_id = r[0] if r else None
            self.open_employee_monthly_performance_dialog(emp_id)

    def open_employee_monthly_performance_dialog(self, employee_id=None):
        dialog = QDialog(self)
        dialog.setWindowTitle("📊 كشف أداء وإنتاجية وحضور العامل الشهري")
        dialog.setLayoutDirection(Qt.RightToLeft)
        dialog.resize(1340, 780)
        dialog.setStyleSheet("background-color: #f8fafc; font-family: 'Segoe UI';")

        d_layout = QVBoxLayout(dialog)
        d_layout.setSpacing(12)
        d_layout.setContentsMargins(16, 16, 16, 16)

        # ── Control Bar ─────────────────────────────────────────────
        ctrl_frame = QFrame()
        ctrl_frame.setStyleSheet(
            "background: white; border-radius: 8px; padding: 10px 14px; border: 1px solid #e2e8f0;"
        )
        ctrl_layout = QHBoxLayout(ctrl_frame)
        ctrl_layout.setSpacing(12)

        lbl_emp = QLabel("👤 العامل:")
        lbl_emp.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_emp.setStyleSheet("color: #334155;")
        ctrl_layout.addWidget(lbl_emp)

        emp_combo = QComboBox()
        emp_combo.setStyleSheet(
            "font-size: 12px; padding: 6px 12px; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 180px; background: white; color: #0f172a;"
        )

        conn = sqlite3.connect(DB_NAME)
        employees = conn.execute("SELECT id, name FROM employees ORDER BY name").fetchall()
        conn.close()
        for eid, ename in employees:
            emp_combo.addItem(ename, eid)

        if employee_id is not None:
            idx = emp_combo.findData(employee_id)
            if idx >= 0:
                emp_combo.setCurrentIndex(idx)

        ctrl_layout.addWidget(emp_combo)

        lbl_month = QLabel("📅 الشهر:")
        lbl_month.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_month.setStyleSheet("color: #334155;")
        ctrl_layout.addWidget(lbl_month)

        month_combo = QComboBox()
        months_arabic = [
            "01 - يناير",
            "02 - فبراير",
            "03 - مارس",
            "04 - أبريل",
            "05 - مايو",
            "06 - يونيو",
            "07 - يوليو",
            "08 - أغسطس",
            "09 - سبتمبر",
            "10 - أكتوبر",
            "11 - نوفمبر",
            "12 - ديسمبر",
        ]
        month_combo.addItems(months_arabic)
        month_combo.setCurrentIndex(QDate.currentDate().month() - 1)
        month_combo.setStyleSheet(
            "font-size: 12px; padding: 6px 12px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;"
        )
        ctrl_layout.addWidget(month_combo)

        lbl_year = QLabel("📆 السنة:")
        lbl_year.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_year.setStyleSheet("color: #334155;")
        ctrl_layout.addWidget(lbl_year)

        year_spin = QSpinBox()
        year_spin.setRange(2020, 2035)
        year_spin.setValue(QDate.currentDate().year())
        year_spin.setStyleSheet(
            "font-size: 12px; padding: 6px 12px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; color: #0f172a;"
        )
        ctrl_layout.addWidget(year_spin)

        btn_refresh = QPushButton("🔄 عرض التقرير")
        btn_refresh.setCursor(Qt.PointingHandCursor)
        btn_refresh.setStyleSheet("""
            QPushButton { background: #2563eb; color: white; padding: 7px 18px; font-weight: bold; border-radius: 6px; border: none; font-size: 12px; }
            QPushButton:hover { background: #1d4ed8; }
        """)
        ctrl_layout.addWidget(btn_refresh)

        btn_export = QPushButton("📥 تصدير Excel")
        btn_export.setCursor(Qt.PointingHandCursor)
        btn_export.setStyleSheet("""
            QPushButton { background: #16a34a; color: white; padding: 7px 16px; font-weight: bold; border-radius: 6px; border: none; font-size: 12px; }
            QPushButton:hover { background: #15803d; }
        """)
        ctrl_layout.addWidget(btn_export)

        ctrl_layout.addStretch()
        d_layout.addWidget(ctrl_frame)

        # ── KPI Summary Cards ───────────────────────────────────────
        kpi_frame = QFrame()
        kpi_frame.setStyleSheet("background: transparent; border: none;")
        kpi_layout = QHBoxLayout(kpi_frame)
        kpi_layout.setSpacing(10)
        kpi_layout.setContentsMargins(0, 0, 0, 0)

        def _make_kpi_card(title, value, subtext, color_theme):
            card = QFrame()
            card.setStyleSheet(f"""
                QFrame {{
                    background: white; border: 1.5px solid #e2e8f0; border-radius: 8px; padding: 10px 14px;
                    border-right: 4px solid {color_theme};
                }}
            """)
            cl = QVBoxLayout(card)
            cl.setSpacing(3)
            cl.setContentsMargins(0, 0, 0, 0)

            t_lbl = QLabel(title)
            t_lbl.setFont(QFont("Segoe UI", 9, QFont.Bold))
            t_lbl.setStyleSheet("color: #64748b; border: none; background: transparent;")

            v_lbl = QLabel(value)
            v_lbl.setFont(QFont("Segoe UI", 15, QFont.Bold))
            v_lbl.setStyleSheet(f"color: {color_theme}; border: none; background: transparent;")

            s_lbl = QLabel(subtext)
            s_lbl.setFont(QFont("Segoe UI", 8))
            s_lbl.setStyleSheet("color: #94a3b8; border: none; background: transparent;")

            cl.addWidget(t_lbl)
            cl.addWidget(v_lbl)
            cl.addWidget(s_lbl)
            return card, v_lbl, s_lbl

        card_att, card_att_val, card_att_sub = _make_kpi_card(
            "📅 الحضور والغياب", "0 يوم", "0% نسبة الحضور", "#2563eb"
        )
        card_hrs, card_hrs_val, card_hrs_sub = _make_kpi_card(
            "⏱️ ساعات العمل", "0.0 س", "إضافي: 0.0 س", "#0891b2"
        )
        card_prod, card_prod_val, card_prod_sub = _make_kpi_card(
            "🏭 الإنتاج الفعلي", "0 قطعة", "0.0 كجم ألومنيوم", "#059669"
        )
        card_ach, card_ach_val, card_ach_sub = _make_kpi_card(
            "🎯 نسبة الإنجاز", "0%", "المستهدف: 0 قطعة", "#d97706"
        )
        card_fin, card_fin_val, card_fin_sub = _make_kpi_card(
            "💰 المعاملات المالية", "+0 ج.م", "خصومات: 0 ج.م", "#7c3aed"
        )

        kpi_layout.addWidget(card_att)
        kpi_layout.addWidget(card_hrs)
        kpi_layout.addWidget(card_prod)
        kpi_layout.addWidget(card_ach)
        kpi_layout.addWidget(card_fin)
        d_layout.addWidget(kpi_frame)

        # ── Day-by-Day Table ────────────────────────────────────────
        table = QTableWidget()
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        table.setColumnCount(13)
        table.setHorizontalHeaderLabels([
            "#",
            "اليوم والتاريخ",
            "حالة الحضور",
            "وقت الحضور",
            "وقت الانصراف",
            "ساعات العمل",
            "السهرات / الإضافي",
            "الماكينة والاسطمبة",
            "الإنتاج الفعلي (قطع)",
            "الإنتاج المستهدف",
            "نسبة الإنجاز",
            "حوافز (+)",
            "خصومات (-)",
        ])
        table.setAlternatingRowColors(True)
        table.setStyleSheet("""
            QTableWidget {
                border: 1.5px solid #cbd5e1;
                border-radius: 8px;
                background: #ffffff;
                gridline-color: #f1f5f9;
                alternate-background-color: #f8fafc;
            }
            QHeaderView::section {
                background-color: #f1f5f9;
                color: #1e293b;
                font-weight: bold;
                font-size: 11px;
                padding: 8px 6px;
                border: none;
                border-bottom: 2px solid #cbd5e1;
                border-left: 1px solid #e2e8f0;
            }
            QTableWidget::item {
                padding: 6px 8px;
                color: #1e293b;
                font-size: 11px;
                border-bottom: 1px solid #f1f5f9;
            }
            QTableWidget::item:selected {
                background-color: #eff6ff;
                color: #1d4ed8;
                font-weight: 600;
            }
        """)
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(36)
        table.setLayoutDirection(Qt.RightToLeft)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        hdr = table.horizontalHeader()
        hdr.setStretchLastSection(False)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(7, QHeaderView.Stretch)
        hdr.setSectionResizeMode(8, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(9, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(10, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(11, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(12, QHeaderView.ResizeToContents)
        d_layout.addWidget(table)

        def load_report():
            sel_emp_id = emp_combo.currentData()
            if not sel_emp_id:
                table.setRowCount(0)
                return

            month = month_combo.currentIndex() + 1
            year = year_spin.value()
            days_in_month = calendar.monthrange(year, month)[1]

            conn = sqlite3.connect(DB_NAME)
            c = conn.cursor()

            # 1. Fetch attendance for this month
            start_date_str = f"{year:04d}-{month:02d}-01"
            end_date_str = f"{year:04d}-{month:02d}-{days_in_month:02d}"

            c.execute(
                """SELECT date, status, check_in_time, check_out_time, work_hours, notes
                   FROM attendance WHERE employee_id=? AND date BETWEEN ? AND ?""",
                (sel_emp_id, start_date_str, end_date_str),
            )
            att_by_date = {row[0]: row for row in c.fetchall()}

            # 2. Fetch overtime
            c.execute(
                """SELECT date, start_time, end_time, notes
                   FROM overtime WHERE employee_id=? AND date BETWEEN ? AND ?""",
                (sel_emp_id, start_date_str, end_date_str),
            )
            ot_by_date = {}
            for row in c.fetchall():
                ot_by_date.setdefault(row[0], []).append(row)

            # 3. Fetch production records with machine & mold
            c.execute(
                """SELECT mp.date, mp.quantity, mp.total_weight_kg,
                          COALESCE(ma.machine_name,''), COALESCE(mo.mold_name,''),
                          COALESCE(mo.target_daily_rate, 0), mp.notes
                   FROM machine_production mp
                   LEFT JOIN machines ma ON mp.machine_id=ma.id
                   LEFT JOIN molds mo ON mp.mold_id=mo.id
                   WHERE mp.employee_id=? AND mp.date BETWEEN ? AND ?""",
                (sel_emp_id, start_date_str, end_date_str),
            )
            prod_by_date = {}
            for row in c.fetchall():
                prod_by_date.setdefault(row[0], []).append(row)

            # 4. Fetch deductions & incentives
            c.execute(
                """SELECT date, amount, reason FROM deductions
                   WHERE employee_id=? AND date BETWEEN ? AND ?""",
                (sel_emp_id, start_date_str, end_date_str),
            )
            ded_by_date = {}
            for row in c.fetchall():
                ded_by_date.setdefault(row[0], []).append(row)

            c.execute(
                """SELECT date, amount, reason FROM incentives
                   WHERE employee_id=? AND date BETWEEN ? AND ?""",
                (sel_emp_id, start_date_str, end_date_str),
            )
            inc_by_date = {}
            for row in c.fetchall():
                inc_by_date.setdefault(row[0], []).append(row)

            conn.close()

            # Aggregate stats
            total_work_days = 0
            present_days = 0
            absent_days = 0
            total_work_hours = 0.0
            total_ot_hours = 0.0
            total_produced_qty = 0
            total_produced_kg = 0.0
            total_target_qty = 0
            total_incentives = 0.0
            total_deductions = 0.0
            achievement_rates = []

            arabic_day_names = [
                "الاثنين",
                "الثلاثاء",
                "الأربعاء",
                "الخميس",
                "الجمعة",
                "السبت",
                "الأحد",
            ]

            table.setRowCount(0)

            for day in range(1, days_in_month + 1):
                cur_date = datetime(year, month, day)
                cur_date_str = f"{year:04d}-{month:02d}-{day:02d}"
                day_name = arabic_day_names[cur_date.weekday()]
                is_friday = cur_date.weekday() == 4

                if not is_friday:
                    total_work_days += 1

                att_record = att_by_date.get(cur_date_str)
                att_status = att_record[1] if att_record else ("جمعة" if is_friday else "غياب")
                check_in = format_time_ar(att_record[2]) if (att_record and att_record[2]) else ""
                check_out = format_time_ar(att_record[3]) if (att_record and att_record[3]) else ""
                work_hrs = (
                    (att_record[4] or 0.0) if att_record and att_status == "حضور" else 0.0
                )

                if att_status == "حضور":
                    present_days += 1
                    total_work_hours += work_hrs
                elif not is_friday:
                    absent_days += 1

                # Overtime
                ot_records = ot_by_date.get(cur_date_str, [])
                ot_hrs_day = 0.0
                ot_texts = []
                for _, s_time, e_time, ot_note in ot_records:
                    t_start = parse_time_str(s_time)
                    t_end = parse_time_str(e_time)
                    if t_start.isValid() and t_end.isValid():
                        dur = (t_end.hour() * 60 + t_end.minute()) - (
                            t_start.hour() * 60 + t_start.minute()
                        )
                        if dur < 0:
                            dur += 24 * 60
                        dur_hrs = dur / 60.0
                    else:
                        dur_hrs = 0.0
                    ot_hrs_day += dur_hrs
                    ot_texts.append(f"{format_time_ar(s_time)} - {format_time_ar(e_time)} ({dur_hrs:.1f}س)")
                total_ot_hours += ot_hrs_day

                # Production
                prod_records = prod_by_date.get(cur_date_str, [])
                day_qty = sum(pr[1] or 0 for pr in prod_records)
                day_kg = sum(pr[2] or 0.0 for pr in prod_records)
                total_produced_qty += day_qty
                total_produced_kg += day_kg

                mach_mold_list = []
                day_target_rate = 0
                for pr in prod_records:
                    _, q, kg, m_name, mo_name, trg, _ = pr
                    item_desc = f"{m_name} / {mo_name}" if (m_name or mo_name) else "إنتاج"
                    mach_mold_list.append(f"{item_desc} ({q:,} قطعة)")
                    day_target_rate += trg or 0

                total_target_qty += day_target_rate

                # Achievement %
                if day_target_rate > 0:
                    ach_pct = (day_qty / day_target_rate) * 100.0
                    achievement_rates.append(ach_pct)
                    ach_str = f"{ach_pct:.1f}%"
                elif day_qty > 0:
                    ach_pct = 100.0
                    achievement_rates.append(ach_pct)
                    ach_str = "100%"
                else:
                    ach_pct = 0.0
                    ach_str = "—" if (att_status != "حضور" or is_friday) else "0%"

                # Financials
                ded_records = ded_by_date.get(cur_date_str, [])
                day_ded = sum(dr[1] or 0.0 for dr in ded_records)
                total_deductions += day_ded

                inc_records = inc_by_date.get(cur_date_str, [])
                day_inc = sum(ir[1] or 0.0 for ir in inc_records)
                total_incentives += day_inc

                # Insert Row into Table
                row_idx = table.rowCount()
                table.insertRow(row_idx)

                # 0: #
                it0 = QTableWidgetItem(str(day))
                it0.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_idx, 0, it0)

                # 1: Date & Day
                it1 = QTableWidgetItem(f"{day_name} {cur_date_str}")
                it1.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                it1.setFont(QFont("Segoe UI", 10, QFont.Bold if is_friday else QFont.Normal))
                table.setItem(row_idx, 1, it1)

                # 2: Status
                if is_friday:
                    it2 = QTableWidgetItem("🕌 جمعة")
                    it2.setForeground(QColor("#7c3aed"))
                elif att_status == "حضور":
                    it2 = QTableWidgetItem("✅ حضور")
                    it2.setForeground(QColor("#16a34a"))
                    it2.setFont(QFont("Segoe UI", 10, QFont.Bold))
                else:
                    it2 = QTableWidgetItem("❌ غياب")
                    it2.setForeground(QColor("#dc2626"))
                    it2.setFont(QFont("Segoe UI", 10, QFont.Bold))
                it2.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_idx, 2, it2)

                # 3: Check in
                it3 = QTableWidgetItem(check_in if (att_status == "حضور" and check_in) else "—")
                it3.setTextAlignment(Qt.AlignCenter)
                if check_in and att_status == "حضور":
                    it3.setForeground(QColor("#0284c7"))
                table.setItem(row_idx, 3, it3)

                # 4: Check out
                it4 = QTableWidgetItem(check_out if (att_status == "حضور" and check_out) else "—")
                it4.setTextAlignment(Qt.AlignCenter)
                if check_out and att_status == "حضور":
                    it4.setForeground(QColor("#0284c7"))
                table.setItem(row_idx, 4, it4)

                # 5: Work hours
                it5 = QTableWidgetItem(
                    f"{work_hrs:.1f} س" if (att_status == "حضور" and work_hrs > 0) else "—"
                )
                it5.setTextAlignment(Qt.AlignCenter)
                if work_hrs > 0 and att_status == "حضور":
                    it5.setFont(QFont("Segoe UI", 10, QFont.Bold))
                table.setItem(row_idx, 5, it5)

                # 6: Overtime
                if ot_hrs_day > 0:
                    it6 = QTableWidgetItem(f"🌙 {ot_hrs_day:.1f} س\n({', '.join(ot_texts)})")
                    it6.setForeground(QColor("#7c3aed"))
                    it6.setFont(QFont("Segoe UI", 9, QFont.Bold))
                else:
                    it6 = QTableWidgetItem("—")
                it6.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_idx, 6, it6)

                # 7: Machine & Mold
                it7 = QTableWidgetItem("\n".join(mach_mold_list) if mach_mold_list else "—")
                it7.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(row_idx, 7, it7)

                # 8: Actual Production Qty
                it8 = QTableWidgetItem(f"{day_qty:,}" if day_qty > 0 else "—")
                it8.setTextAlignment(Qt.AlignCenter)
                if day_qty > 0:
                    it8.setFont(QFont("Segoe UI", 10, QFont.Bold))
                    it8.setForeground(QColor("#059669"))
                table.setItem(row_idx, 8, it8)

                # 9: Target Production Qty
                it9 = QTableWidgetItem(f"{day_target_rate:,}" if day_target_rate > 0 else "—")
                it9.setTextAlignment(Qt.AlignCenter)
                table.setItem(row_idx, 9, it9)

                # 10: Achievement %
                it10 = QTableWidgetItem(ach_str)
                it10.setTextAlignment(Qt.AlignCenter)
                it10.setFont(QFont("Segoe UI", 10, QFont.Bold))
                if ach_str != "—":
                    if ach_pct >= 100:
                        it10.setForeground(QColor("#15803d"))
                    elif ach_pct >= 75:
                        it10.setForeground(QColor("#d97706"))
                    else:
                        it10.setForeground(QColor("#dc2626"))
                table.setItem(row_idx, 10, it10)

                # 11: Incentives
                it11 = QTableWidgetItem(f"+{day_inc:,.2f}" if day_inc > 0 else "—")
                it11.setTextAlignment(Qt.AlignCenter)
                if day_inc > 0:
                    it11.setForeground(QColor("#16a34a"))
                    it11.setFont(QFont("Segoe UI", 10, QFont.Bold))
                table.setItem(row_idx, 11, it11)

                # 12: Deductions
                it12 = QTableWidgetItem(f"-{day_ded:,.2f}" if day_ded > 0 else "—")
                it12.setTextAlignment(Qt.AlignCenter)
                if day_ded > 0:
                    it12.setForeground(QColor("#dc2626"))
                    it12.setFont(QFont("Segoe UI", 10, QFont.Bold))
                table.setItem(row_idx, 12, it12)

                # Row background color highlight
                if is_friday:
                    bg_color = QColor("#f5f3ff")
                elif att_status == "غياب":
                    bg_color = QColor("#fef2f2")
                elif day_qty > 0 and ach_pct >= 100:
                    bg_color = QColor("#f0fdf4")
                else:
                    bg_color = None

                if bg_color:
                    for col in range(table.columnCount()):
                        item = table.item(row_idx, col)
                        if item:
                            item.setBackground(bg_color)

            # Update KPI Card values
            att_pct = (present_days / total_work_days * 100) if total_work_days > 0 else 0
            card_att_val.setText(f"{present_days} يوم")
            card_att_sub.setText(f"من أصل {total_work_days} يوم عمل (حضور {att_pct:.1f}%)")

            card_hrs_val.setText(f"{total_work_hours:.1f} س")
            card_hrs_sub.setText(f"إضافي/سهرات: {total_ot_hours:.1f} س")

            card_prod_val.setText(f"{total_produced_qty:,} قطعة")
            card_prod_sub.setText(f"خام مستهلك: {total_produced_kg:.2f} كجم")

            avg_ach = (
                (sum(achievement_rates) / len(achievement_rates))
                if achievement_rates
                else 0.0
            )
            card_ach_val.setText(f"{avg_ach:.1f}%")
            card_ach_sub.setText(f"إجمالي المستهدف: {total_target_qty:,} قطعة")

            net_fin = total_incentives - total_deductions
            card_fin_val.setText(f"{net_fin:+,.2f} ج.م")
            card_fin_sub.setText(
                f"حوافز: +{total_incentives:,.1f} | خصومات: -{total_deductions:,.1f}"
            )

        def export_excel():
            if table.rowCount() == 0:
                QMessageBox.warning(dialog, "تنبيه", "لا توجد بيانات لتصديرها.")
                return

            emp_name = emp_combo.currentText()
            month_text = month_combo.currentText()
            year_val = year_spin.value()
            default_filename = f"تقرير_انتاجية_وحضور_{emp_name}_{year_val}_{month_combo.currentIndex()+1}.xlsx"

            file_path, _ = QFileDialog.getSaveFileName(
                dialog, "تصدير تقرير أداء العامل", default_filename, "Excel Files (*.xlsx)"
            )
            if not file_path:
                return
            if not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"

            headers = [
                "#",
                "اليوم والتاريخ",
                "حالة الحضور",
                "وقت الحضور",
                "وقت الانصراف",
                "ساعات العمل",
                "السهرات / الإضافي",
                "الماكينة والاسطمبة",
                "الإنتاج الفعلي (قطع)",
                "الإنتاج المستهدف",
                "نسبة الإنجاز",
                "حوافز (+)",
                "خصومات (-)",
            ]
            data_rows = []
            for r in range(table.rowCount()):
                row_data = []
                for c in range(table.columnCount()):
                    item = table.item(r, c)
                    row_data.append(item.text() if item else "")
                data_rows.append(row_data)

            summary_info = (
                f"العامل: {emp_name} | الفترة: {month_text} {year_val} | "
                f"أيام الحضور: {card_att_val.text()} ({card_att_sub.text()}) | "
                f"ساعات العمل: {card_hrs_val.text()} | "
                f"إجمالي الإنتاج: {card_prod_val.text()} | "
                f"متوسط الإنجاز: {card_ach_val.text()}"
            )

            self._write_styled_xlsx(
                file_path,
                f"كشف أداء وإنتاجية العامل: {emp_name} ({month_text} {year_val})",
                headers,
                data_rows,
                "تقرير أداء العامل",
                extra_info=summary_info,
            )
            QMessageBox.information(
                dialog, "تصدير", f"تم تصدير تقرير العامل {emp_name} بنجاح ✔️"
            )

        btn_refresh.clicked.connect(load_report)
        btn_export.clicked.connect(export_excel)
        emp_combo.currentIndexChanged.connect(load_report)
        month_combo.currentIndexChanged.connect(load_report)
        year_spin.valueChanged.connect(load_report)

        load_report()
        dialog.exec()

    # ══════════════════════════════════════════════════════════════════════
    # 🏢 تبويب إدارة الأقسام (Department Management)
    # ══════════════════════════════════════════════════════════════════════

    def setup_department_management_tab(self, tab):
        layout = QVBoxLayout(tab)
        layout.setSpacing(10)
        layout.setContentsMargins(12, 12, 12, 12)

        # 1. Header Summary Cards
        kpi_frame = QFrame()
        kpi_frame.setStyleSheet("background: transparent; border: none;")
        kpi_layout = QHBoxLayout(kpi_frame)
        kpi_layout.setSpacing(10)
        kpi_layout.setContentsMargins(0, 0, 0, 0)

        self.card_dept_count = self.create_hr_metric_card("🏢 إجمالي الأقسام", "0", "#2563eb", "#eff6ff")
        self.card_dept_assigned = self.create_hr_metric_card("👥 العمال الموزعين", "0", "#16a34a", "#f0fdf4")
        self.card_dept_unassigned = self.create_hr_metric_card("⚠️ عمال بدون قسم", "0", "#dc2626", "#fef2f2")
        self.card_dept_total_sal = self.create_hr_metric_card("💰 إجمالي رواتب الأقسام", "0 ج.م", "#7c3aed", "#f5f3ff")

        kpi_layout.addWidget(self.card_dept_count)
        kpi_layout.addWidget(self.card_dept_assigned)
        kpi_layout.addWidget(self.card_dept_unassigned)
        kpi_layout.addWidget(self.card_dept_total_sal)
        layout.addWidget(kpi_frame)

        # 2. Control & Quick Add Bar
        ctrl_frame = QFrame()
        ctrl_frame.setStyleSheet("background: #f8fafc; border-radius: 8px; padding: 10px; border: 1px solid #e2e8f0;")
        ctrl_layout = QVBoxLayout(ctrl_frame)
        ctrl_layout.setSpacing(8)
        ctrl_layout.setContentsMargins(12, 10, 12, 10)

        # Row 1: Name, Code, Shift start & end, Hours, Notes
        r1_layout = QHBoxLayout()
        r1_layout.setSpacing(10)

        lbl_name = QLabel("اسم القسم:")
        lbl_name.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_name.setStyleSheet("color: #334155;")
        r1_layout.addWidget(lbl_name)

        self.dept_name_input = QLineEdit()
        self.dept_name_input.setPlaceholderText("اكتب اسم القسم...")
        self.dept_name_input.setStyleSheet("font-size: 12px; padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; min-width: 150px;")
        r1_layout.addWidget(self.dept_name_input)

        lbl_code = QLabel("كود:")
        lbl_code.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_code.setStyleSheet("color: #334155;")
        r1_layout.addWidget(lbl_code)

        self.dept_code_input = QLineEdit()
        self.dept_code_input.setPlaceholderText("كود...")
        self.dept_code_input.setStyleSheet("font-size: 12px; padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white; max-width: 90px;")
        r1_layout.addWidget(self.dept_code_input)

        lbl_in = QLabel("🕒 حضور:")
        lbl_in.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_in.setStyleSheet("color: #0369a1;")
        r1_layout.addWidget(lbl_in)

        self.dept_start_edit = QTimeEdit(QTime(8, 0))
        self.dept_start_edit.setDisplayFormat("hh:mm AP")
        self.dept_start_edit.setLocale(ARABIC_LOCALE)
        self.dept_start_edit.setStyleSheet("padding: 5px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white;")
        r1_layout.addWidget(self.dept_start_edit)

        lbl_out = QLabel("🕒 انصراف:")
        lbl_out.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_out.setStyleSheet("color: #b45309;")
        r1_layout.addWidget(lbl_out)

        self.dept_end_edit = QTimeEdit(QTime(19, 0))
        self.dept_end_edit.setDisplayFormat("hh:mm AP")
        self.dept_end_edit.setLocale(ARABIC_LOCALE)
        self.dept_end_edit.setStyleSheet("padding: 5px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white;")
        r1_layout.addWidget(self.dept_end_edit)

        lbl_notes = QLabel("ملاحظات:")
        lbl_notes.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_notes.setStyleSheet("color: #334155;")
        r1_layout.addWidget(lbl_notes)

        self.dept_notes_input = QLineEdit()
        self.dept_notes_input.setPlaceholderText("ملاحظات القسم...")
        self.dept_notes_input.setStyleSheet("font-size: 12px; padding: 6px 10px; border: 1.5px solid #cbd5e1; border-radius: 6px; background: white;")
        r1_layout.addWidget(self.dept_notes_input, 1)

        add_dept_btn = QPushButton("➕ إضافة قسم")
        add_dept_btn.setCursor(Qt.PointingHandCursor)
        add_dept_btn.setStyleSheet("""
            QPushButton {
                background-color: #16a34a; color: white; padding: 7px 16px;
                font-weight: bold; border-radius: 6px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #15803d; }
        """)
        add_dept_btn.clicked.connect(self.add_department_handler)
        r1_layout.addWidget(add_dept_btn)

        sync_btn = QPushButton("📥 استيراد من الإكسل")
        sync_btn.setCursor(Qt.PointingHandCursor)
        sync_btn.setStyleSheet("""
            QPushButton {
                background-color: #7c3aed; color: white; padding: 7px 14px;
                font-weight: bold; border-radius: 6px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #6d28d9; }
        """)
        sync_btn.clicked.connect(self.sync_excel_departments_handler)
        r1_layout.addWidget(sync_btn)

        ctrl_layout.addLayout(r1_layout)
        layout.addWidget(ctrl_frame)

        # 3. Search Bar
        search_frame = QHBoxLayout()
        search_frame.addWidget(QLabel("🔍 بحث في الأقسام:"))
        self.dept_search_input = QLineEdit()
        self.dept_search_input.setPlaceholderText("ابحث باسم القسم أو الكود أو المواعيد...")
        self.dept_search_input.setClearButtonEnabled(True)
        self.dept_search_input.setStyleSheet("font-size: 12px; padding: 6px 10px; border: 1px solid #cbd5e1; border-radius: 6px; background: white;")
        self.dept_search_input.textChanged.connect(self.filter_departments_table)
        search_frame.addWidget(self.dept_search_input)
        layout.addLayout(search_frame)

        # 4. Departments Table
        self.dept_table = QTableWidget()
        self.dept_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.dept_table.setColumnCount(9)
        self.dept_table.setHorizontalHeaderLabels([
            "ID", "كود القسم", "اسم القسم", "مواعيد الوردية", "عدد العمال المرتبطين", "إجمالي الرواتب الأساسية", "ملاحظات", "تعديل", "حذف"
        ])
        self.dept_table.setAlternatingRowColors(True)
        self.dept_table.setStyleSheet("""
            QTableWidget {
                border: 1.5px solid #cbd5e1; border-radius: 8px; background: #ffffff;
                gridline-color: #e2e8f0; alternate-background-color: #f8fafc;
                selection-background-color: #dbeafe; selection-color: #1e3a8a;
            }
            QTableWidget::item { padding: 8px 12px; font-size: 13px; color: #0f172a; border-bottom: 1px solid #e2e8f0; }
            QTableWidget::item:selected { background: #dbeafe; color: #1e3a8a; font-weight: bold; }
            QHeaderView::section {
                background: #0f172a; color: #ffffff; padding: 10px 10px;
                font-weight: bold; font-size: 13px; border: 1px solid #1e293b; border-bottom: 3px solid #2563eb;
            }
        """)
        dhdr = self.dept_table.horizontalHeader()
        dhdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        dhdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        dhdr.setSectionResizeMode(2, QHeaderView.Stretch)
        dhdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        dhdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        dhdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        dhdr.setSectionResizeMode(6, QHeaderView.Stretch)
        dhdr.setSectionResizeMode(7, QHeaderView.ResizeToContents)
        dhdr.setSectionResizeMode(8, QHeaderView.ResizeToContents)
        self.dept_table.verticalHeader().setDefaultSectionSize(46)
        layout.addWidget(self.dept_table)

        self.load_departments_table()

    def load_departments_table(self):
        """تحميل وعرض بيانات الأقسام وإحصائيات العمال ومواعيد العمل."""
        if not hasattr(self, "dept_table"):
            return

        with self.preserve_table_scroll(self.dept_table):
            self.dept_table.setUpdatesEnabled(False)
            self.dept_table.setRowCount(0)

            stats = get_all_departments_with_stats()
            total_depts = len(stats)
            total_assigned = sum(d["employee_count"] for d in stats)
            total_payroll = sum(d["total_base_salary"] for d in stats)

            conn = get_db_connection()
            c = conn.cursor()
            c.execute("SELECT COUNT(*) FROM employees WHERE department_id IS NULL AND is_active=1")
            unassigned_count = c.fetchone()[0] or 0
            conn.close()

            if hasattr(self, "card_dept_count"):
                self.card_dept_count.val_label.setText(str(total_depts))
                self.card_dept_assigned.val_label.setText(str(total_assigned))
                self.card_dept_unassigned.val_label.setText(str(unassigned_count))
                self.card_dept_total_sal.val_label.setText(f"{total_payroll:,.0f} ج.م")

            for r_idx, d in enumerate(stats):
                self.dept_table.insertRow(r_idx)
                dept_id = d["id"]
                dept_name = d["name"]
                emp_count = d["employee_count"]
                s_time = d.get("start_time", "08:00 ص")
                e_time = d.get("end_time", "07:00 م")
                w_hours = d.get("default_work_hours", 11.0)

                # 0. ID
                id_item = QTableWidgetItem(str(dept_id))
                id_item.setTextAlignment(Qt.AlignCenter)
                self.dept_table.setItem(r_idx, 0, id_item)

                # 1. Code
                code_item = QTableWidgetItem(d["code"] or "-")
                code_item.setTextAlignment(Qt.AlignCenter)
                self.dept_table.setItem(r_idx, 1, code_item)

                # 2. Name
                name_item = QTableWidgetItem(dept_name)
                name_item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                self.dept_table.setItem(r_idx, 2, name_item)

                # 3. Shift schedule badge
                shift_text = f"🕒 {s_time} - {e_time} ({w_hours:g} س)"
                shift_item = QTableWidgetItem(shift_text)
                shift_item.setTextAlignment(Qt.AlignCenter)
                shift_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                shift_item.setForeground(QColor("#0369a1"))
                self.dept_table.setItem(r_idx, 3, shift_item)

                # 4. Employee Count Badge / Button
                workers_btn = QPushButton(f"👥 {emp_count} عامل (عرض)")
                workers_btn.setCursor(Qt.PointingHandCursor)
                workers_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #e0f2fe; color: #0369a1; font-weight: bold;
                        border-radius: 6px; padding: 4px 10px; font-size: 12px; border: 1px solid #bae6fd;
                    }
                    QPushButton:hover { background-color: #0284c7; color: white; }
                """)
                workers_btn.clicked.connect(lambda _, did=dept_id, dname=dept_name: self.show_department_workers_dialog(did, dname))
                self.dept_table.setCellWidget(r_idx, 4, workers_btn)

                # 5. Total Salaries
                sal_item = QTableWidgetItem(f"{d['total_base_salary']:,.2f} ج.م")
                sal_item.setTextAlignment(Qt.AlignCenter)
                sal_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                sal_item.setForeground(QColor("#059669"))
                self.dept_table.setItem(r_idx, 5, sal_item)

                # 6. Notes
                notes_item = QTableWidgetItem(d["notes"] or "")
                self.dept_table.setItem(r_idx, 6, notes_item)

                # 7. Edit Button
                edit_btn = QPushButton("✏️ تعديل")
                edit_btn.setCursor(Qt.PointingHandCursor)
                edit_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #2563eb; color: white; border-radius: 6px;
                        padding: 4px 10px; font-weight: bold; font-size: 11px; border: none;
                    }
                    QPushButton:hover { background-color: #1d4ed8; }
                """)
                edit_btn.clicked.connect(lambda _, did=dept_id: self.edit_department_dialog(did))
                self.dept_table.setCellWidget(r_idx, 7, edit_btn)

                # 8. Delete Button
                del_btn = QPushButton("🗑️ حذف")
                del_btn.setCursor(Qt.PointingHandCursor)
                del_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #dc2626; color: white; border-radius: 6px;
                        padding: 4px 10px; font-weight: bold; font-size: 11px; border: none;
                    }
                    QPushButton:hover { background-color: #b91c1c; }
                """)
                del_btn.clicked.connect(lambda _, did=dept_id, dname=dept_name, ecnt=emp_count: self.delete_department_handler(did, dname, ecnt))
                self.dept_table.setCellWidget(r_idx, 8, del_btn)

            self.dept_table.setUpdatesEnabled(True)

    def filter_departments_table(self, text=""):
        q = text.strip().lower()
        for r in range(self.dept_table.rowCount()):
            code_it = self.dept_table.item(r, 1)
            name_it = self.dept_table.item(r, 2)
            shift_it = self.dept_table.item(r, 3)
            notes_it = self.dept_table.item(r, 6)

            t_code = code_it.text().lower() if code_it else ""
            t_name = name_it.text().lower() if name_it else ""
            t_shift = shift_it.text().lower() if shift_it else ""
            t_notes = notes_it.text().lower() if notes_it else ""

            match = (not q or q in t_code or q in t_name or q in t_shift or q in t_notes)
            self.dept_table.setRowHidden(r, not match)

    def add_department_handler(self):
        name = self.dept_name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "تنبيه", "يرجى كتابة اسم القسم أولاً.")
            return
        code = self.dept_code_input.text().strip() or None
        notes = self.dept_notes_input.text().strip() or None
        start_str = format_time_ar(self.dept_start_edit.time())
        end_str = format_time_ar(self.dept_end_edit.time())
        work_h = self.calculate_work_hours(self.dept_start_edit.time(), self.dept_end_edit.time())

        res = add_department(name, code, notes, start_str, end_str, work_h)
        if res.get("success"):
            self.dept_name_input.clear()
            self.dept_code_input.clear()
            self.dept_notes_input.clear()
            self.load_departments_table()
            self.refresh_hr_views()
            if hasattr(self, "show_toast"):
                self.show_toast(f"✅ تم إضافة قسم '{name}' بمواعيد {start_str} إلى {end_str} بنجاح")
        else:
            QMessageBox.critical(self, "خطأ", f"تعذر إضافة القسم: {res.get('error')}")

    def edit_department_dialog(self, dept_id):
        dept = get_department_by_id(dept_id)
        if not dept:
            QMessageBox.warning(self, "تنبيه", "لم يتم العثور على بيانات القسم.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle(f"تعديل قسم: {dept['name']}")
        dialog.setLayoutDirection(Qt.RightToLeft)
        dialog.resize(480, 320)
        dialog.setStyleSheet("background: white; font-family: 'Segoe UI';")

        layout = QVBoxLayout(dialog)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        title = QLabel(f"✏️ تعديل بيانات قسم ومواعيد الوردية: {dept['name']}")
        title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        layout.addWidget(title)

        grid = QGridLayout()
        grid.setSpacing(8)

        grid.addWidget(QLabel("اسم القسم:"), 0, 0)
        name_in = QLineEdit(dept["name"])
        name_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px;")
        grid.addWidget(name_in, 0, 1)

        grid.addWidget(QLabel("كود القسم:"), 1, 0)
        code_in = QLineEdit(dept["code"] or "")
        code_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px;")
        grid.addWidget(code_in, 1, 1)

        grid.addWidget(QLabel("🕒 وقت الحضور (البداية):"), 2, 0)
        start_in = QTimeEdit()
        start_in.setDisplayFormat("hh:mm AP")
        start_in.setLocale(ARABIC_LOCALE)
        t_s = parse_time_str(dept.get("start_time", "08:00 ص"))
        start_in.setTime(t_s if t_s.isValid() else QTime(8, 0))
        start_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px;")
        grid.addWidget(start_in, 2, 1)

        grid.addWidget(QLabel("🕒 وقت الانصراف (النهاية):"), 3, 0)
        end_in = QTimeEdit()
        end_in.setDisplayFormat("hh:mm AP")
        end_in.setLocale(ARABIC_LOCALE)
        t_e = parse_time_str(dept.get("end_time", "07:00 م"))
        end_in.setTime(t_e if t_e.isValid() else QTime(19, 0))
        end_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px;")
        grid.addWidget(end_in, 3, 1)

        grid.addWidget(QLabel("ملاحظات:"), 4, 0)
        notes_in = QLineEdit(dept["notes"] or "")
        notes_in.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px;")
        grid.addWidget(notes_in, 4, 1)

        layout.addLayout(grid)

        btn_box = QHBoxLayout()
        save_btn = QPushButton("💾 حفظ التعديلات")
        save_btn.setStyleSheet("background: #16a34a; color: white; padding: 7px 16px; font-weight: bold; border-radius: 6px;")
        cancel_btn = QPushButton("إلغاء")
        cancel_btn.setStyleSheet("background: #64748b; color: white; padding: 7px 16px; font-weight: bold; border-radius: 6px;")

        btn_box.addWidget(save_btn)
        btn_box.addWidget(cancel_btn)
        layout.addLayout(btn_box)

        def save_action():
            n = name_in.text().strip()
            if not n:
                QMessageBox.warning(dialog, "تنبيه", "اسم القسم مطلوب.")
                return
            c = code_in.text().strip() or ""
            nt = notes_in.text().strip() or ""
            st_str = format_time_ar(start_in.time())
            et_str = format_time_ar(end_in.time())
            wh = self.calculate_work_hours(start_in.time(), end_in.time())

            res = update_department(dept_id, n, c, nt, st_str, et_str, wh)
            if res.get("success"):
                dialog.accept()
                self.load_departments_table()
                self.refresh_hr_views()
                if hasattr(self, "show_toast"):
                    self.show_toast(f"✅ تم تحديث قسم '{n}' بنجاح")
            else:
                QMessageBox.critical(dialog, "خطأ", f"فشل التعديل: {res.get('error')}")

        save_btn.clicked.connect(save_action)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec()

    def delete_department_handler(self, dept_id, dept_name, emp_count):
        msg = f"هل أنت متأكد من حذف قسم: <b>{dept_name}</b>؟"
        if emp_count > 0:
            msg += f"<br><br><span style='color: #dc2626;'>⚠️ يحتوي هذا القسم على <b>{emp_count}</b> عامل مرتبط به. سيتم إلغاء ربط العمال بهذا القسم دون حذف بياناتهم.</span>"

        reply = QMessageBox.question(
            self, "تأكيد حذف القسم", msg,
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            res = delete_department(dept_id)
            if res.get("success"):
                self.load_departments_table()
                self.refresh_hr_views()
                if hasattr(self, "show_toast"):
                    self.show_toast(f"✅ تم حذف قسم '{dept_name}' بنجاح")
            else:
                QMessageBox.critical(self, "خطأ", f"تعذر حذف القسم: {res.get('error')}")

    def sync_excel_departments_handler(self):
        reply = QMessageBox.question(
            self, "استيراد من ملف الإكسل",
            "هل تريد استيراد الأقسام والعمال والرواتب وحوافز الانتظام من ملف الإكسل المرفق؟",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            res = import_excel_payroll_data()
            if res.get("success"):
                QMessageBox.information(
                    self, "تم الاستيراد بنجاح",
                    f"تم استيراد ومزامنة البيانات بنجاح ✔️\n• إجمالي العاملين: {res['total_employees']}\n• مضاف جديد: {res['new_imported']}\n• محدث: {res['updated']}\n• عدد الأقسام: {res['departments_count']}"
                )
                self.load_departments_table()
                self.refresh_hr_views()
            else:
                QMessageBox.critical(self, "خطأ في الاستيراد", f"فشل الاستيراد: {res.get('error')}")

    def show_department_workers_dialog(self, dept_id, dept_name):
        """نافذة استعراض ونقل وإدارة عمال القسم."""
        workers = get_department_employees(dept_id)
        all_depts = get_departments()

        dialog = QDialog(self)
        dialog.setWindowTitle(f"عمال قسم: {dept_name} ({len(workers)} عامل)")
        dialog.setLayoutDirection(Qt.RightToLeft)
        dialog.resize(780, 520)
        dialog.setStyleSheet("background: white; font-family: 'Segoe UI';")

        layout = QVBoxLayout(dialog)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        header = QLabel(f"🏢 قائمة العاملين بقسم: <b>{dept_name}</b> (إجمالي: {len(workers)} عامل)")
        header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        layout.addWidget(header)

        # Worker Table
        w_table = QTableWidget()
        w_table.setColumnCount(6)
        w_table.setHorizontalHeaderLabels([
            "ID", "الكود", "اسم العامل", "الراتب الأساسي", "حافز الانتظام", "نقل لقسم آخر"
        ])
        w_table.setColumnHidden(0, True)
        w_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        w_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        w_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        w_table.setStyleSheet("""
            QTableWidget { border: 1px solid #cbd5e1; border-radius: 6px; background: white; }
            QHeaderView::section { background: #f1f5f9; color: #1e293b; padding: 8px; font-weight: bold; border: none; }
            QTableWidget::item { padding: 6px; }
        """)

        def load_w_table():
            nonlocal workers
            workers = get_department_employees(dept_id)
            w_table.setRowCount(0)
            for r_idx, w in enumerate(workers):
                w_table.insertRow(r_idx)
                if isinstance(w, dict):
                    eid = w["id"]
                    code = w.get("code", "")
                    name = w["name"]
                    base_sal = w.get("base_salary", 0.0)
                    inc = w.get("attendance_incentive", 500.0)
                else:
                    eid = w[0]
                    name = w[1]
                    code = w[2] if len(w) > 2 else ""
                    base_sal = w[3] if len(w) > 3 else 0.0
                    inc = w[4] if len(w) > 4 else 500.0

                w_table.setItem(r_idx, 0, QTableWidgetItem(str(eid)))
                w_table.setItem(r_idx, 1, QTableWidgetItem(str(code or "-")))
                
                n_item = QTableWidgetItem(name)
                n_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                w_table.setItem(r_idx, 2, n_item)

                sal_item = QTableWidgetItem(f"{float(base_sal):,.2f} ج.م")
                sal_item.setForeground(QColor("#059669"))
                w_table.setItem(r_idx, 3, sal_item)

                inc_item = QTableWidgetItem(f"{float(inc):,.2f} ج.م")
                inc_item.setForeground(QColor("#2563eb"))
                w_table.setItem(r_idx, 4, inc_item)

                # Transfer Button
                tr_btn = QPushButton("🔄 نقل لقسم...")
                tr_btn.setCursor(Qt.PointingHandCursor)
                tr_btn.setStyleSheet("background: #0284c7; color: white; font-weight: bold; border-radius: 4px; padding: 3px 8px; font-size: 11px;")
                tr_btn.clicked.connect(lambda _, worker_id=eid, w_name=name: transfer_single_worker(worker_id, w_name))
                w_table.setCellWidget(r_idx, 5, tr_btn)

        def transfer_single_worker(worker_id, worker_name):
            t_dialog = QDialog(dialog)
            t_dialog.setWindowTitle(f"نقل العامل: {worker_name}")
            t_dialog.setLayoutDirection(Qt.RightToLeft)
            t_dialog.resize(360, 180)
            t_layout = QVBoxLayout(t_dialog)
            t_layout.addWidget(QLabel(f"اختر القسم الجديد للعامل: <b>{worker_name}</b>"))
            
            cbox = QComboBox()
            for row in all_depts:
                did = row[0]
                dname = row[1]
                if did != dept_id:
                    cbox.addItem(f"📁 {dname}", did)
            t_layout.addWidget(cbox)

            b_box = QHBoxLayout()
            ok_b = QPushButton("نقل الآن")
            ok_b.setStyleSheet("background: #16a34a; color: white; padding: 6px 14px; font-weight: bold; border-radius: 4px;")
            canc_b = QPushButton("إلغاء")
            canc_b.setStyleSheet("background: #64748b; color: white; padding: 6px 14px; font-weight: bold; border-radius: 4px;")
            b_box.addWidget(ok_b)
            b_box.addWidget(canc_b)
            t_layout.addLayout(b_box)

            def do_transfer():
                target_did = cbox.currentData()
                conn = get_db_connection()
                c = conn.cursor()
                c.execute("UPDATE employees SET department_id=? WHERE id=?", (target_did, worker_id))
                conn.commit()
                conn.close()
                t_dialog.accept()
                load_w_table()
                self.load_departments_table()
                self.refresh_hr_views()

            ok_b.clicked.connect(do_transfer)
            canc_b.clicked.connect(t_dialog.reject)
            t_dialog.exec()

        load_w_table()
        layout.addWidget(w_table)

        close_btn = QPushButton("إغلاق")
        close_btn.setStyleSheet("background: #64748b; color: white; padding: 7px 16px; font-weight: bold; border-radius: 6px;")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)

        dialog.exec()

    # ══════════════════════════════════════════════════════════════════════
    # 💵 تبويب مسير الرواتب والأجور (Payroll & Salary Sheet)
    # ══════════════════════════════════════════════════════════════════════

    def setup_payroll_tab(self, tab):
        layout = QVBoxLayout(tab)
        layout.setSpacing(10)
        layout.setContentsMargins(12, 12, 12, 12)

        # 1. Top Control Bar
        ctrl_frame = QFrame()
        ctrl_frame.setStyleSheet("background: #f8fafc; border-radius: 8px; padding: 10px; border: 1px solid #e2e8f0;")
        ctrl_layout = QHBoxLayout(ctrl_frame)
        ctrl_layout.setSpacing(10)
        ctrl_layout.setContentsMargins(10, 8, 10, 8)

        # Month
        lbl_m = QLabel("📅 الشهر:")
        lbl_m.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_m.setStyleSheet("color: #334155;")
        ctrl_layout.addWidget(lbl_m)

        self.payroll_month_combo = QComboBox()
        month_names = [
            "1 - يناير", "2 - فبراير", "3 - مارس", "4 - إبريل", "5 - مايو", "6 - يونيو",
            "7 - يوليو", "8 - أغسطس", "9 - سبتمبر", "10 - أكتوبر", "11 - نوفمبر", "12 - ديسمبر"
        ]
        for m_idx, m_name in enumerate(month_names, 1):
            self.payroll_month_combo.addItem(m_name, m_idx)
        cur_m = datetime.now().month
        self.payroll_month_combo.setCurrentIndex(cur_m - 1)
        self.payroll_month_combo.setStyleSheet("font-size: 12px; padding: 5px 8px; border: 1px solid #cbd5e1; border-radius: 6px; background: white;")
        ctrl_layout.addWidget(self.payroll_month_combo)

        # Year
        lbl_y = QLabel("📆 السنة:")
        lbl_y.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_y.setStyleSheet("color: #334155;")
        ctrl_layout.addWidget(lbl_y)

        self.payroll_year_spin = QSpinBox()
        self.payroll_year_spin.setRange(2020, 2035)
        self.payroll_year_spin.setValue(datetime.now().year)
        self.payroll_year_spin.setStyleSheet("font-size: 12px; padding: 5px 8px; border: 1px solid #cbd5e1; border-radius: 6px; background: white;")
        ctrl_layout.addWidget(self.payroll_year_spin)

        # Department Filter
        lbl_d = QLabel("🏢 القسم:")
        lbl_d.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_d.setStyleSheet("color: #334155;")
        ctrl_layout.addWidget(lbl_d)

        self.payroll_dept_filter = QComboBox()
        self.payroll_dept_filter.setStyleSheet("font-size: 12px; padding: 5px 8px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; min-width: 150px;")
        self.populate_dept_filter_combo(self.payroll_dept_filter)
        ctrl_layout.addWidget(self.payroll_dept_filter)

        # Search Query
        lbl_s = QLabel("🔍 بحث:")
        lbl_s.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_s.setStyleSheet("color: #334155;")
        ctrl_layout.addWidget(lbl_s)

        self.payroll_search_input = QLineEdit()
        self.payroll_search_input.setPlaceholderText("ابحث بالاسم أو الكود...")
        self.payroll_search_input.setClearButtonEnabled(True)
        self.payroll_search_input.setStyleSheet("font-size: 12px; padding: 5px 8px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; min-width: 160px;")
        self.payroll_search_input.textChanged.connect(self.filter_payroll_table)
        ctrl_layout.addWidget(self.payroll_search_input)

        ctrl_layout.addStretch()

        # Action Buttons
        calc_btn = QPushButton("🧮 حساب المسير")
        calc_btn.setCursor(Qt.PointingHandCursor)
        calc_btn.setStyleSheet("""
            QPushButton {
                background-color: #2563eb; color: white; padding: 6px 14px;
                font-weight: bold; border-radius: 6px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #1d4ed8; }
        """)
        calc_btn.clicked.connect(self.load_payroll_table)
        ctrl_layout.addWidget(calc_btn)

        export_btn = QPushButton("📥 تصدير إلى Excel")
        export_btn.setCursor(Qt.PointingHandCursor)
        export_btn.setStyleSheet("""
            QPushButton {
                background-color: #16a34a; color: white; padding: 6px 14px;
                font-weight: bold; border-radius: 6px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #15803d; }
        """)
        export_btn.clicked.connect(self.export_payroll_to_excel)
        ctrl_layout.addWidget(export_btn)

        rules_btn = QPushButton("ℹ️ خطة وقواعد الراتب")
        rules_btn.setCursor(Qt.PointingHandCursor)
        rules_btn.setStyleSheet("""
            QPushButton {
                background-color: #475569; color: white; padding: 6px 12px;
                font-weight: bold; border-radius: 6px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #334155; }
        """)
        rules_btn.clicked.connect(self.show_payroll_rules_dialog)
        ctrl_layout.addWidget(rules_btn)

        layout.addWidget(ctrl_frame)

        # 2. KPI Summary Ribbon
        kpi_frame = QFrame()
        kpi_frame.setStyleSheet("background: transparent; border: none;")
        kpi_layout = QHBoxLayout(kpi_frame)
        kpi_layout.setSpacing(10)
        kpi_layout.setContentsMargins(0, 0, 0, 0)

        self.card_payroll_count = self.create_hr_metric_card("👥 عمال الكشف", "0", "#0f172a", "#f1f5f9")
        self.card_payroll_base = self.create_hr_metric_card("💰 إجمالي الأساسي", "0 ج.م", "#2563eb", "#eff6ff")
        self.card_payroll_incentives = self.create_hr_metric_card("🎁 حوافز الانتظام", "0 ج.م", "#16a34a", "#f0fdf4")
        self.card_payroll_deductions = self.create_hr_metric_card("📉 إجمالي الاستقطاعات", "0 ج.م", "#dc2626", "#fef2f2")
        self.card_payroll_net = self.create_hr_metric_card("💵 صافي مسير الرواتب", "0 ج.م", "#7c3aed", "#f5f3ff")

        kpi_layout.addWidget(self.card_payroll_count)
        kpi_layout.addWidget(self.card_payroll_base)
        kpi_layout.addWidget(self.card_payroll_incentives)
        kpi_layout.addWidget(self.card_payroll_deductions)
        kpi_layout.addWidget(self.card_payroll_net)
        layout.addWidget(kpi_frame)

        # 3. Payroll Table
        self.payroll_table = QTableWidget()
        self.payroll_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.payroll_table.setColumnCount(16)
        self.payroll_table.setHorizontalHeaderLabels([
            "ID", "الكود", "اسم العامل", "القسم", "الراتب الأساسي", "أجر اليوم",
            "إجمالي الغياب", "أيام الاحتساب 🎁", "الغياب المخصوم", "حافز الانتظام", "خصم غياب الأساسي", "ساعات السهرات",
            "قيمة السهرات", "مكافآت", "خصومات", "صافي القبض المستحق"
        ])
        self.payroll_table.setColumnHidden(0, True)
        self.payroll_table.setAlternatingRowColors(True)
        self.payroll_table.setStyleSheet("""
            QTableWidget {
                border: 1.5px solid #cbd5e1; border-radius: 8px; background: #ffffff;
                gridline-color: #e2e8f0; alternate-background-color: #f8fafc;
                selection-background-color: #dbeafe; selection-color: #1e3a8a;
            }
            QTableWidget::item { padding: 8px 10px; font-size: 13px; color: #0f172a; border-bottom: 1px solid #e2e8f0; }
            QTableWidget::item:selected { background: #dbeafe; color: #1e3a8a; font-weight: bold; }
            QHeaderView::section {
                background: #0f172a; color: #ffffff; padding: 10px 8px;
                font-weight: bold; font-size: 13px; border: 1px solid #1e293b; border-bottom: 3px solid #2563eb;
            }
        """)
        phdr = self.payroll_table.horizontalHeader()
        phdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        phdr.setSectionResizeMode(2, QHeaderView.Stretch)
        phdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        for c_idx in range(4, 16):
            phdr.setSectionResizeMode(c_idx, QHeaderView.ResizeToContents)
        self.payroll_table.verticalHeader().setDefaultSectionSize(44)
        self.payroll_table.cellDoubleClicked.connect(self._on_payroll_cell_double_clicked)
        layout.addWidget(self.payroll_table)

        # Wire filter changes
        self.payroll_month_combo.currentIndexChanged.connect(self.load_payroll_table)
        self.payroll_year_spin.valueChanged.connect(self.load_payroll_table)
        self.payroll_dept_filter.currentIndexChanged.connect(self.load_payroll_table)

        self.load_payroll_table()

    def load_payroll_table(self):
        """حساب وعرض مسير الرواتب للشهر والسنة والأقسام المحددة مع تفاصيل أيام الاحتساب."""
        if not hasattr(self, "payroll_table"):
            return

        month = self.payroll_month_combo.currentData() or (self.payroll_month_combo.currentIndex() + 1)
        year = self.payroll_year_spin.value()
        dept_id = self.payroll_dept_filter.currentData()

        with self.preserve_table_scroll(self.payroll_table):
            self.payroll_table.setUpdatesEnabled(False)
            self.payroll_table.setRowCount(0)

            summary = calculate_monthly_payroll_summary(year, month, dept_id)
            records = summary.get("records", [])

            # Update KPI Cards
            if hasattr(self, "card_payroll_count"):
                self.card_payroll_count.val_label.setText(str(summary["count"]))
                self.card_payroll_base.val_label.setText(f"{summary['tot_base_salary']:,.0f} ج.م")
                self.card_payroll_incentives.val_label.setText(f"{summary['tot_earned_incentive']:,.0f} ج.م")
                total_ded = summary["tot_absence_base_deduction"] + summary["tot_incentive_deduction"] + summary["tot_penalties"]
                self.card_payroll_deductions.val_label.setText(f"{total_ded:,.0f} ج.م")
                self.card_payroll_net.val_label.setText(f"{summary['tot_net_payout']:,.0f} ج.م")

            for r_idx, rec in enumerate(records):
                self.payroll_table.insertRow(r_idx)

                # 0. ID
                self.payroll_table.setItem(r_idx, 0, QTableWidgetItem(str(rec["employee_id"])))

                # 1. Code
                code_item = QTableWidgetItem(str(rec["code"] or "-"))
                code_item.setTextAlignment(Qt.AlignCenter)
                code_item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                self.payroll_table.setItem(r_idx, 1, code_item)

                # 2. Name
                name_item = QTableWidgetItem(rec["name"])
                name_item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                name_item.setForeground(QColor("#0f172a"))
                name_item.setData(Qt.UserRole, rec)
                self.payroll_table.setItem(r_idx, 2, name_item)

                # 3. Dept
                dept_item = QTableWidgetItem(rec["department_name"])
                dept_item.setTextAlignment(Qt.AlignCenter)
                dept_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                dept_item.setForeground(QColor("#0369a1"))
                dept_item.setData(Qt.UserRole, rec["department_id"])
                self.payroll_table.setItem(r_idx, 3, dept_item)

                # 4. Base Salary
                sal_item = QTableWidgetItem(f"{rec['base_salary']:,.2f} ج.م")
                sal_item.setTextAlignment(Qt.AlignCenter)
                sal_item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                sal_item.setForeground(QColor("#047857"))
                self.payroll_table.setItem(r_idx, 4, sal_item)

                # 5. Daily Wage
                dw_item = QTableWidgetItem(f"{rec['daily_wage']:,.2f} ج.م")
                dw_item.setTextAlignment(Qt.AlignCenter)
                dw_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                dw_item.setForeground(QColor("#475569"))
                self.payroll_table.setItem(r_idx, 5, dw_item)

                # 6. Raw Absent Days (إجمالي الغياب)
                raw_abs = rec.get("raw_absent_days", rec["absent_days"])
                raw_abs_item = QTableWidgetItem(str(raw_abs))
                raw_abs_item.setTextAlignment(Qt.AlignCenter)
                if raw_abs > 0:
                    raw_abs_item.setForeground(QColor("#b91c1c"))
                    raw_abs_item.setBackground(QColor("#fee2e2"))
                    raw_abs_item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                else:
                    raw_abs_item.setFont(QFont("Segoe UI", 10))
                    raw_abs_item.setForeground(QColor("#64748b"))
                self.payroll_table.setItem(r_idx, 6, raw_abs_item)

                # 7. Credited Days (أيام الاحتساب على المصنع)
                cred_days = rec.get("credited_days", 0.0)
                if cred_days > 0:
                    cred_item = QTableWidgetItem(f"🎁 {cred_days:g} يوم")
                    cred_item.setTextAlignment(Qt.AlignCenter)
                    cred_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                    cred_item.setForeground(QColor("#6b21a8"))
                    cred_item.setBackground(QColor("#f5f3ff"))
                    cred_item.setToolTip("أيام احتساب مدفوعة على حساب المصنع - لا تخصم من العامل")
                else:
                    cred_item = QTableWidgetItem("-")
                    cred_item.setTextAlignment(Qt.AlignCenter)
                    cred_item.setForeground(QColor("#94a3b8"))
                self.payroll_table.setItem(r_idx, 7, cred_item)

                # 8. Deducted Absent Days (الغياب المخصوم)
                eff_abs = rec["absent_days"]
                eff_abs_item = QTableWidgetItem(str(eff_abs))
                eff_abs_item.setTextAlignment(Qt.AlignCenter)
                if eff_abs > 0:
                    eff_abs_item.setForeground(QColor("#b45309"))
                    eff_abs_item.setBackground(QColor("#fefce8"))
                    eff_abs_item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                    eff_abs_item.setToolTip(f"أيام الغياب الخاضعة للخصم الفعلي بعد استبعاد الاحتساب ({raw_abs} - {cred_days} = {eff_abs})")
                else:
                    eff_abs_item.setFont(QFont("Segoe UI", 10, QFont.Bold if raw_abs > 0 else QFont.Normal))
                    eff_abs_item.setForeground(QColor("#15803d" if raw_abs > 0 else "#64748b"))
                    eff_abs_item.setBackground(QColor("#f0fdf4" if raw_abs > 0 else "#ffffff"))
                    if raw_abs > 0 and cred_days >= raw_abs:
                        eff_abs_item.setToolTip("تم احتساب جميع أيام الغياب على المصنع بالكامل (0 يوم خصم)")
                self.payroll_table.setItem(r_idx, 8, eff_abs_item)

                # 9. Earned Attendance Incentive
                inc_item = QTableWidgetItem(f"{rec['earned_incentive']:,.2f} ج.م")
                inc_item.setTextAlignment(Qt.AlignCenter)
                if rec['earned_incentive'] > 0:
                    inc_item.setForeground(QColor("#15803d"))
                    inc_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                else:
                    inc_item.setForeground(QColor("#94a3b8"))
                self.payroll_table.setItem(r_idx, 9, inc_item)

                # 10. Absence Base Deduction
                abd_item = QTableWidgetItem(f"-{rec['absence_base_deduction']:,.2f} ج.م" if rec['absence_base_deduction'] > 0 else "0.00 ج.م")
                abd_item.setTextAlignment(Qt.AlignCenter)
                if rec['absence_base_deduction'] > 0:
                    abd_item.setForeground(QColor("#dc2626"))
                    abd_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                else:
                    abd_item.setForeground(QColor("#94a3b8"))
                self.payroll_table.setItem(r_idx, 10, abd_item)

                # 11. Overtime Hours
                ot_h_item = QTableWidgetItem(f"{rec['overtime_hours']:.1f} س")
                ot_h_item.setTextAlignment(Qt.AlignCenter)
                ot_h_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                ot_h_item.setForeground(QColor("#7c3aed") if rec['overtime_hours'] > 0 else QColor("#94a3b8"))
                self.payroll_table.setItem(r_idx, 11, ot_h_item)

                # 12. Overtime Pay
                ot_p_item = QTableWidgetItem(f"+{rec['overtime_pay']:,.2f} ج.م" if rec['overtime_pay'] > 0 else "0.00 ج.م")
                ot_p_item.setTextAlignment(Qt.AlignCenter)
                if rec['overtime_pay'] > 0:
                    ot_p_item.setForeground(QColor("#0284c7"))
                    ot_p_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                else:
                    ot_p_item.setForeground(QColor("#94a3b8"))
                self.payroll_table.setItem(r_idx, 12, ot_p_item)

                # 13. Extra Bonuses
                bon_item = QTableWidgetItem(f"+{rec['extra_bonuses_total']:,.2f} ج.م" if rec['extra_bonuses_total'] > 0 else "0.00 ج.م")
                bon_item.setTextAlignment(Qt.AlignCenter)
                if rec['extra_bonuses_total'] > 0:
                    bon_item.setForeground(QColor("#16a34a"))
                    bon_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                else:
                    bon_item.setForeground(QColor("#94a3b8"))
                self.payroll_table.setItem(r_idx, 13, bon_item)

                # 14. Penalties / Deductions
                pen_item = QTableWidgetItem(f"-{rec['penalties_total']:,.2f} ج.م" if rec['penalties_total'] > 0 else "0.00 ج.م")
                pen_item.setTextAlignment(Qt.AlignCenter)
                if rec['penalties_total'] > 0:
                    pen_item.setForeground(QColor("#dc2626"))
                    pen_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                else:
                    pen_item.setForeground(QColor("#94a3b8"))
                self.payroll_table.setItem(r_idx, 14, pen_item)

                # 15. Net Payout (صافي القبض المستحق - بارز ومميز)
                net_item = QTableWidgetItem(f"{rec['net_payout']:,.2f} ج.م")
                net_item.setTextAlignment(Qt.AlignCenter)
                net_item.setFont(QFont("Segoe UI", 12, QFont.Bold))
                net_item.setBackground(QColor("#fef3c7"))
                net_item.setForeground(QColor("#92400e"))
                self.payroll_table.setItem(r_idx, 15, net_item)

            self.payroll_table.setUpdatesEnabled(True)

        if hasattr(self, "payroll_search_input") and self.payroll_search_input.text():
            self.filter_payroll_table(self.payroll_search_input.text())

    def filter_payroll_table(self, text=""):
        q = text.strip().lower()
        for r in range(self.payroll_table.rowCount()):
            c_it = self.payroll_table.item(r, 1)
            n_it = self.payroll_table.item(r, 2)
            d_it = self.payroll_table.item(r, 3)

            t_code = c_it.text().lower() if c_it else ""
            t_name = n_it.text().lower() if n_it else ""
            t_dept = d_it.text().lower() if d_it else ""

            match = (not q or q in t_code or q in t_name or q in t_dept)
            self.payroll_table.setRowHidden(r, not match)

    def _on_payroll_cell_double_clicked(self, row, col):
        name_item = self.payroll_table.item(row, 2)
        if not name_item:
            return
        rec = name_item.data(Qt.UserRole)
        if rec:
            self.show_employee_payslip_dialog(rec)

    def show_employee_payslip_dialog(self, rec):
        """نافذة مفردات مرتب تفصيلية للعامل (Pay Slip) مع تفاصيل أيام الاحتساب."""
        dialog = QDialog(self)
        dialog.setWindowTitle(f"مفردات مرتب: {rec['name']} - شهر {rec['month']}/{rec['year']}")
        dialog.setLayoutDirection(Qt.RightToLeft)
        dialog.resize(540, 620)
        dialog.setStyleSheet("background: white; font-family: 'Segoe UI';")

        layout = QVBoxLayout(dialog)
        layout.setSpacing(12)
        layout.setContentsMargins(18, 18, 18, 18)

        # Header Slip Frame
        h_frame = QFrame()
        h_frame.setStyleSheet("background: #0f172a; border-radius: 8px; padding: 12px;")
        h_layout = QVBoxLayout(h_frame)
        
        t1 = QLabel(f"📄 كشف مفردات مرتب للعامل: {rec['name']}")
        t1.setFont(QFont("Segoe UI", 12, QFont.Bold))
        t1.setStyleSheet("color: white;")
        h_layout.addWidget(t1)

        t2 = QLabel(f"الكود: {rec['code']} | القسم: {rec['department_name']} | شهر: {rec['month']} / {rec['year']}")
        t2.setFont(QFont("Segoe UI", 10))
        t2.setStyleSheet("color: #94a3b8;")
        h_layout.addWidget(t2)
        layout.addWidget(h_frame)

        # Details Grid Frame
        grid_frame = QFrame()
        grid_frame.setStyleSheet("background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px;")
        grid = QGridLayout(grid_frame)
        grid.setSpacing(8)

        def add_row(title, val_str, row_i, color="#0f172a", is_bold=False):
            lbl_t = QLabel(title)
            lbl_t.setFont(QFont("Segoe UI", 10, QFont.Bold if is_bold else QFont.Normal))
            lbl_t.setStyleSheet("color: #475569;")
            lbl_v = QLabel(val_str)
            lbl_v.setFont(QFont("Segoe UI", 10, QFont.Bold if is_bold else QFont.Normal))
            lbl_v.setStyleSheet(f"color: {color};")
            lbl_v.setAlignment(Qt.AlignLeft)
            grid.addWidget(lbl_t, row_i, 0)
            grid.addWidget(lbl_v, row_i, 1)

        cur_row = 0
        add_row("الراتب الأساسي الشهري:", f"{rec['base_salary']:,.2f} ج.م", cur_row, "#0f172a", True); cur_row += 1
        daily_div_val = rec.get('daily_divisor', 30.0) if hasattr(rec, 'get') else 30.0
        add_row("أجر اليوم (الأساسي ÷ قاسم الأيام):", f"{rec['daily_wage']:,.2f} ج.م (أجر الساعة: {rec.get('hourly_rate', 0.0):,.2f} ج.م)", cur_row, "#0284c7"); cur_row += 1
        add_row("أيام الحضور الفعلية:", f"{rec['present_days']} يوم", cur_row, "#16a34a"); cur_row += 1

        raw_abs = rec.get('raw_absent_days', rec['absent_days'])
        cred_days = rec.get('credited_days', 0.0)
        eff_abs = rec['absent_days']

        add_row("إجمالي أيام الغياب المسجلة:", f"{raw_abs} يوم", cur_row, "#dc2626" if raw_abs > 0 else "#475569"); cur_row += 1

        if cred_days > 0:
            add_row("أيام الاحتساب (مدفوعة على المصنع):", f"🎁 {cred_days:g} يوم (لا تُخصم من العامل)", cur_row, "#7c3aed", True); cur_row += 1
            add_row("أيام الغياب الخاضعة للخصم:", f"{eff_abs:g} يوم (بعد استبعاد أيام الاحتساب)", cur_row, "#b45309" if eff_abs > 0 else "#15803d", True); cur_row += 1
        else:
            add_row("أيام الغياب الخاضعة للخصم:", f"{eff_abs:g} يوم", cur_row, "#dc2626" if eff_abs > 0 else "#475569"); cur_row += 1

        add_row("حافز الانتظام المستحق:", f"+ {rec['earned_incentive']:,.2f} ج.م (من أصل {rec['max_incentive']:,.0f})", cur_row, "#16a34a", True); cur_row += 1

        if rec['absence_base_deduction'] > 0:
            add_row("خصم غياب من الراتب الأساسي:", f"- {rec['absence_base_deduction']:,.2f} ج.م ({rec['absent_days_from_base']} يوم)", cur_row, "#dc2626"); cur_row += 1
        if rec['incentive_deduction'] > 0:
            add_row("خصم غياب من حافز الانتظام:", f"- {rec['incentive_deduction']:,.2f} ج.م", cur_row, "#dc2626"); cur_row += 1
        
        ot_desc = f"{rec['overtime_hours']:.1f} ساعة"
        norm_h = rec.get('normal_ot_hours', 0.0)
        fri_h = rec.get('friday_ot_hours', 0.0)
        norm_m = rec.get('overtime_normal_mult', 1.5)
        fri_m = rec.get('overtime_friday_mult', 2.0)
        if norm_h > 0 or fri_h > 0:
            ot_desc += f" (عادي: {norm_h:.1f}س × {norm_m} | جمعة: {fri_h:.1f}س × {fri_m})"
        add_row("ساعات السهرات والإضافي:", ot_desc, cur_row, "#0f172a"); cur_row += 1
        add_row("قيمة أجر السهرات المستحقة:", f"+ {rec['overtime_pay']:,.2f} ج.م", cur_row, "#0284c7", True); cur_row += 1

        if rec['extra_bonuses_total'] > 0:
            add_row("مكافآت وحوافز إضافية:", f"+ {rec['extra_bonuses_total']:,.2f} ج.م", cur_row, "#16a34a"); cur_row += 1
        if rec['penalties_total'] > 0:
            add_row("جزاءات واستقطاعات أخرى:", f"- {rec['penalties_total']:,.2f} ج.م", cur_row, "#dc2626"); cur_row += 1
        if rec.get('monthly_advances_total', 0.0) > 0:
            add_row("السلف الشهرية المخصومة:", f"- {rec['monthly_advances_total']:,.2f} ج.م", cur_row, "#dc2626", True); cur_row += 1
        if rec.get('loan_installments_total', 0.0) > 0:
            add_row("قسط سلفة طويلة الأجل:", f"- {rec['loan_installments_total']:,.2f} ج.م", cur_row, "#dc2626", True); cur_row += 1

        layout.addWidget(grid_frame)

        # Net Payout Banner
        net_frame = QFrame()
        net_frame.setStyleSheet("background: #f5f3ff; border: 1.5px solid #c4b5fd; border-radius: 8px; padding: 12px;")
        net_layout = QHBoxLayout(net_frame)
        nlbl1 = QLabel("💵 صافي الراتب النهائي المستحق للقبض:")
        nlbl1.setFont(QFont("Segoe UI", 11, QFont.Bold))
        nlbl1.setStyleSheet("color: #6b21a8;")
        net_layout.addWidget(nlbl1)

        nlbl2 = QLabel(f"{rec['net_payout']:,.2f} ج.م")
        nlbl2.setFont(QFont("Segoe UI", 14, QFont.Bold))
        nlbl2.setStyleSheet("color: #7c3aed;")
        net_layout.addWidget(nlbl2)
        layout.addWidget(net_frame)

        btn_box = QHBoxLayout()
        cred_btn = QPushButton("✨ إدارة أيام الاحتساب لهذا العامل")
        cred_btn.setCursor(Qt.PointingHandCursor)
        cred_btn.setStyleSheet("background: #7c3aed; color: white; padding: 8px 16px; font-weight: bold; border-radius: 6px; font-size: 11px;")
        cred_btn.clicked.connect(lambda: [dialog.accept(), self.open_add_credited_leave_dialog(rec['employee_id'], callback=self.load_payroll_table)])
        btn_box.addWidget(cred_btn)

        btn_box.addStretch()

        btn_close = QPushButton("إغلاق")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setStyleSheet("background: #64748b; color: white; padding: 8px 18px; font-weight: bold; border-radius: 6px;")
        btn_close.clicked.connect(dialog.accept)
        btn_box.addWidget(btn_close)

        layout.addLayout(btn_box)

        dialog.exec()

    def show_payroll_rules_dialog(self):
        """نافذة توضيح خطة وقواعد احتساب حافز الانتظام والغياب والإضافي وأيام الاحتساب."""
        dialog = QDialog(self)
        dialog.setWindowTitle("خطة وقواعد الراتب وحافز الانتظام والإضافي")
        dialog.setLayoutDirection(Qt.RightToLeft)
        dialog.resize(640, 560)
        dialog.setStyleSheet("background: white; font-family: 'Segoe UI';")

        layout = QVBoxLayout(dialog)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        title = QLabel("📜 خطة احتساب الراتب، حافز الانتظام، أيام الاحتساب وقواعد الخصم")
        title.setFont(QFont("Segoe UI", 13, QFont.Bold))
        title.setStyleSheet("color: #1e3a8a;")
        layout.addWidget(title)

        info_browser = QTextBrowser()
        info_browser.setStyleSheet("border: 1px solid #cbd5e1; border-radius: 8px; background: #f8fafc; padding: 10px; font-size: 13px; color: #1e293b;")
        info_browser.setHtml("""
        <div style='font-family: Segoe UI; line-height: 1.6;'>
            <h3 style='color: #2563eb;'>📌 المبادئ الأساسية لاحتساب الأجور والساعات:</h3>
            <ul>
                <li><b>أجر اليوم:</b> يُحسب بقسمة الراتب الأساسي على 30 يوماً (أو القاسم المحدد في الإعدادات): <code>الأساسي ÷ 30</code>.</li>
                <li><b>أجر الساعة:</b> يُحسب بقسمة أجر اليوم على عدد ساعات وردية العامل: <code>أجر اليوم ÷ ساعات الوردية</code>.</li>
                <li><b>حافز الانتظام الكامل:</b> 500 جنيه شهرياً للعامل المنتظم (قابل للتعديل).</li>
            </ul>

            <h3 style='color: #7c3aed;'>🎁 نظام أيام الاحتساب (الإجازات المدفوعة على حساب المصنع):</h3>
            <ul>
                <li>يمكن للإدارة منح العامل <b>أيام احتساب</b> تعفيه من خصومات الغياب.</li>
                <li><b>معادلة الخصم:</b> <code>الغياب الصافي الخاضع للخصم = إجمالي أيام الغياب - أيام الاحتساب</code>.</li>
                <li>مثال: إذا غاب العامل 5 أيام وحصل على 3 أيام احتساب، يُخصم منه <b>2 يوم فقط</b> و 3 أيام على المصنع.</li>
            </ul>

            <h3 style='color: #0891b2;'>🌙 قواعد احتساب ساعات السهرات والإضافي:</h3>
            <ul>
                <li><b>في الأيام العادية:</b> تُحسب ساعة الإضافي بمعدل <b>1.5 × أجر الساعة</b> (قابلة للتعديل من الإعدادات).</li>
                <li><b>في أيام الجمعة والعطلات:</b> تُحسب ساعة الإضافي بمعدل <b>2.0 × أجر الساعة</b> (قابلة للتعديل من الإعدادات).</li>
            </ul>

            <h3 style='color: #059669;'>🎯 خطة خصم أيام الغياب الصافية (بعد استبعاد الاحتساب):</h3>
            <table border='1' cellspacing='0' cellpadding='6' style='border-collapse: collapse; width: 100%; border-color: #cbd5e1;'>
                <tr style='background-color: #f1f5f9;'>
                    <th align='center'>أيام الغياب الصافية</th>
                    <th align='center'>الخصم من حافز الانتظام</th>
                    <th align='center'>حافز الانتظام المستحق</th>
                    <th align='center'>الخصم من الراتب الأساسي</th>
                </tr>
                <tr>
                    <td align='center'><b>0 يوم</b></td>
                    <td align='center' style='color: #16a34a;'>0 ج.م</td>
                    <td align='center' style='color: #16a34a;'><b>500 ج.م</b> (كامل)</td>
                    <td align='center' style='color: #16a34a;'><b>0 ج.م</b> (لا خصم)</td>
                </tr>
                <tr>
                    <td align='center'><b>1 يوم</b></td>
                    <td align='center' style='color: #dc2626;'>100 ج.م</td>
                    <td align='center' style='color: #2563eb;'><b>400 ج.م</b></td>
                    <td align='center' style='color: #16a34a;'><b>0 ج.م</b> (لا يُخصم من الأساسي)</td>
                </tr>
                <tr>
                    <td align='center'><b>2 يوم</b></td>
                    <td align='center' style='color: #dc2626;'>200 ج.م</td>
                    <td align='center' style='color: #2563eb;'><b>300 ج.م</b></td>
                    <td align='center' style='color: #dc2626;'><b>أجر يوم واحد</b> (الأساسي ÷ 30)</td>
                </tr>
                <tr>
                    <td align='center'><b>3 أيام</b></td>
                    <td align='center' style='color: #dc2626;'>300 ج.م</td>
                    <td align='center' style='color: #2563eb;'><b>200 ج.م</b></td>
                    <td align='center' style='color: #dc2626;'><b>أجر يومين</b> (2 × أجر اليوم)</td>
                </tr>
                <tr>
                    <td align='center'><b>4 أيام فأكثر</b></td>
                    <td align='center' style='color: #dc2626;'>500 ج.م (كامل)</td>
                    <td align='center' style='color: #dc2626;'><b>0 ج.م</b></td>
                    <td align='center' style='color: #dc2626;'><b>أجر (أيام الغياب - 1)</b> من الأساسي</td>
                </tr>
            </table>
        </div>
        """)
        layout.addWidget(info_browser)

        btn_ok = QPushButton("موافق")
        btn_ok.setStyleSheet("background: #2563eb; color: white; padding: 8px 18px; font-weight: bold; border-radius: 6px;")
        btn_ok.clicked.connect(dialog.accept)
        layout.addWidget(btn_ok)

        dialog.exec()

    def export_payroll_to_excel(self):
        """تصدير مسير الرواتب والأجور الشامل إلى ملف Excel منسق متضمناً أيام الاحتساب."""
        month = self.payroll_month_combo.currentData() or (self.payroll_month_combo.currentIndex() + 1)
        year = self.payroll_year_spin.value()
        dept_id = self.payroll_dept_filter.currentData()
        dept_name = self.payroll_dept_filter.currentText()

        summary = calculate_monthly_payroll_summary(year, month, dept_id)
        records = summary.get("records", [])

        if not records:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات مسير رواتب لتصديرها.")
            return

        default_filename = f"مسير_رواتب_شهر_{month:02d}_{year}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "حفظ مسير الرواتب", default_filename, "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = f"مسير {month}-{year}"
            ws.views.sheetView[0].rightToLeft = True

            # 1. Title Header
            ws.merge_cells("A1:O1")
            title_cell = ws["A1"]
            title_cell.value = f"كشف مسير الرواتب والأجور الشامل - شهر {month:02d} / {year} ({dept_name})"
            title_cell.font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 36

            # 2. Table Headers
            headers = [
                "كود العامل", "اسم العامل", "القسم", "الراتب الأساسي", "أجر اليوم",
                "إجمالي الغياب", "أيام الاحتساب (على المصنع)", "الغياب المخصوم", "حافز الانتظام المستحق", "خصم غياب الأساسي", "ساعات السهرات",
                "قيمة السهرات", "مكافآت وحوافز", "جزاءات واستقطاعات", "صافي القبض المستحق"
            ]
            ws.append([]) # Empty Row 2
            ws.append(headers) # Row 3
            ws.row_dimensions[3].height = 26

            header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style="thin", color="CBD5E1"),
                right=Side(style="thin", color="CBD5E1"),
                top=Side(style="thin", color="CBD5E1"),
                bottom=Side(style="thin", color="CBD5E1")
            )

            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=3, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            # 3. Data Rows
            alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            start_row = 4
            for r_i, rec in enumerate(records, start_row):
                row_data = [
                    rec["code"],
                    rec["name"],
                    rec["department_name"],
                    rec["base_salary"],
                    rec["daily_wage"],
                    rec.get("raw_absent_days", rec["absent_days"]),
                    rec.get("credited_days", 0.0),
                    rec["absent_days"],
                    rec["earned_incentive"],
                    rec["absence_base_deduction"],
                    rec["overtime_hours"],
                    rec["overtime_pay"],
                    rec["extra_bonuses_total"],
                    rec["penalties_total"],
                    rec["net_payout"],
                ]
                ws.append(row_data)
                ws.row_dimensions[r_i].height = 22

                is_alt = (r_i % 2 == 0)
                for col_idx in range(1, len(row_data) + 1):
                    cell = ws.cell(row=r_i, column=col_idx)
                    cell.border = thin_border
                    if is_alt:
                        cell.fill = alt_fill

                    # Format types
                    if col_idx == 2:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        cell.font = Font(name="Calibri", size=11, bold=True)
                    elif col_idx == 7 and rec.get("credited_days", 0.0) > 0:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.font = Font(name="Calibri", size=11, bold=True, color="6B21A8")
                        cell.fill = PatternFill(start_color="F5F3FF", end_color="F5F3FF", fill_type="solid")
                    elif col_idx in (4, 5, 9, 10, 12, 13, 14, 15):
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if col_idx == 15:
                            cell.font = Font(name="Calibri", size=11, bold=True, color="6B21A8")
                            cell.fill = PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid")
                    else:
                        cell.alignment = Alignment(horizontal="center", vertical="center")

            # 4. Total Summary Row
            total_row_idx = len(records) + start_row
            total_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            total_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
            total_border = Border(
                left=Side(style="thin", color="94A3B8"),
                right=Side(style="thin", color="94A3B8"),
                top=Side(style="medium", color="64748B"),
                bottom=Side(style="double", color="64748B")
            )

            ws.append([
                "الإجمالي",
                f"عدد العمال: {summary['count']}",
                "-",
                summary["tot_base_salary"],
                "-",
                summary.get("tot_raw_absent_days", summary["tot_absent_days"]),
                summary.get("tot_credited_days", 0.0),
                summary["tot_absent_days"],
                summary["tot_earned_incentive"],
                summary["tot_absence_base_deduction"],
                "-",
                summary["tot_overtime_pay"],
                summary["tot_bonuses"],
                summary["tot_penalties"],
                summary["tot_net_payout"],
            ])
            ws.row_dimensions[total_row_idx].height = 26

            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=total_row_idx, column=col_idx)
                cell.fill = total_fill
                cell.font = total_font
                cell.border = total_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx in (4, 9, 10, 12, 13, 14, 15):
                    cell.number_format = "#,##0.00"

            # 5. Adjust Column Widths
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.row == 1:
                        continue
                    val_str = str(cell.value or "")
                    if len(val_str) > max_len:
                        max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            ws.column_dimensions["B"].width = 28 # Employee Name Column

            wb.save(file_path)
            QMessageBox.information(self, "نجاح التصدير", f"تم تصدير مسير الرواتب بنجاح إلى:\n{file_path} ✔️")

        except Exception as e:
            QMessageBox.critical(self, "خطأ في التصدير", f"حدث خطأ أثناء التصدير:\n{str(e)}")

    def _on_dept_cell_double_clicked(self, row, col):
        id_item = self.dept_table.item(row, 0)
        name_item = self.dept_table.item(row, 2)
        if id_item and name_item:
            self.show_department_workers_dialog(int(id_item.text()), name_item.text())

    # ═════════════════════════════════════════════════════════════════════════
    # ── تبويب الإعدادات وقواعد الرواتب وحوافز الانتظام ─────────────────────────
    # ═════════════════════════════════════════════════════════════════════════

    def setup_payroll_settings_tab(self, tab):
        """بناء وتجهيز تبويب إعدادات دورة الشهر والزيادات وقواعد الغياب ومواعيد الأقسام."""
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(12, 12, 12, 12)
        tab_layout.setSpacing(12)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: #f8fafc; border: none; }")

        container = QWidget()
        container.setStyleSheet("background: #f8fafc;")
        container_layout = QVBoxLayout(container)
        container_layout.setSpacing(16)
        container_layout.setContentsMargins(8, 8, 8, 8)

        # ── Card 1: Payroll Month Cycle ────────────────────────────────────
        card_cycle = QFrame()
        card_cycle.setStyleSheet("""
            QFrame {
                background: white; border-radius: 10px; border: 1.5px solid #cbd5e1; padding: 16px;
            }
        """)
        c1_layout = QVBoxLayout(card_cycle)
        c1_layout.setSpacing(12)

        c1_title = QLabel("📅 دورة الشهر المالي وتحديد بداية ونهاية الشهر")
        c1_title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        c1_title.setStyleSheet("color: #0f172a; border: none;")
        c1_desc = QLabel("حدد يوم بداية ونهاية الشهر المالي لحساب كشوف الحضور ومسير الرواتب والاستقطاعات تلقائياً:")
        c1_desc.setStyleSheet("color: #64748b; font-size: 11px; border: none;")
        c1_layout.addWidget(c1_title)
        c1_layout.addWidget(c1_desc)

        cfg = get_payroll_config()

        self.cycle_type_std = QRadioButton("🔵 شهر تقويمي قياسي (من أول الشهر 1 حتى نهايته 30 / 31)")
        self.cycle_type_std.setStyleSheet("font-size: 12px; font-weight: bold; color: #1e293b; border: none;")
        self.cycle_type_custom = QRadioButton("🟣 دورة مخصصة (مثال: من يوم 29 في الشهر السابق حتى يوم 28 في الشهر الحالي)")
        self.cycle_type_custom.setStyleSheet("font-size: 12px; font-weight: bold; color: #1e293b; border: none;")

        if cfg.get("cycle_type") == "custom_day":
            self.cycle_type_custom.setChecked(True)
        else:
            self.cycle_type_std.setChecked(True)

        c1_layout.addWidget(self.cycle_type_std)
        c1_layout.addWidget(self.cycle_type_custom)

        custom_days_box = QHBoxLayout()
        custom_days_box.setSpacing(16)

        custom_days_box.addWidget(QLabel("يوم بداية الدورة المالية:"))
        self.cycle_start_spin = QSpinBox()
        self.cycle_start_spin.setRange(1, 31)
        self.cycle_start_spin.setValue(int(cfg.get("start_day", 29)))
        self.cycle_start_spin.setStyleSheet("padding: 6px; font-size: 12px; font-weight: bold; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 80px;")
        custom_days_box.addWidget(self.cycle_start_spin)

        custom_days_box.addWidget(QLabel("يوم نهاية الدورة المالية:"))
        self.cycle_end_spin = QSpinBox()
        self.cycle_end_spin.setRange(1, 31)
        self.cycle_end_spin.setValue(int(cfg.get("end_day", 28)))
        self.cycle_end_spin.setStyleSheet("padding: 6px; font-size: 12px; font-weight: bold; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 80px;")
        custom_days_box.addWidget(self.cycle_end_spin)

        custom_days_box.addStretch()
        c1_layout.addLayout(custom_days_box)

        self.cycle_preview_lbl = QLabel("")
        self.cycle_preview_lbl.setStyleSheet("color: #2563eb; font-weight: bold; font-size: 12px; border: none; padding: 4px;")
        c1_layout.addWidget(self.cycle_preview_lbl)

        btn_save_cycle = QPushButton("💾 حفظ إعدادات دورة الشهر المالي")
        btn_save_cycle.setCursor(Qt.PointingHandCursor)
        btn_save_cycle.setStyleSheet("""
            QPushButton {
                background-color: #2563eb; color: white; padding: 8px 18px;
                font-weight: bold; border-radius: 6px; font-size: 12px; border: none;
            }
            QPushButton:hover { background-color: #1d4ed8; }
        """)
        btn_save_cycle.clicked.connect(self.save_cycle_settings_handler)
        c1_layout.addWidget(btn_save_cycle, 0, Qt.AlignRight)

        container_layout.addWidget(card_cycle)

        # ── Card 2: Bulk Salary Increment Tool ─────────────────────────────
        card_bulk = QFrame()
        card_bulk.setStyleSheet("""
            QFrame {
                background: white; border-radius: 10px; border: 1.5px solid #cbd5e1; padding: 16px;
            }
        """)
        c2_layout = QVBoxLayout(card_bulk)
        c2_layout.setSpacing(12)

        c2_title = QLabel("📈 زيادة جماعية للرواتب (% نسبة مئوية أو مبلغ ثابت)")
        c2_title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        c2_title.setStyleSheet("color: #0f172a; border: none;")
        c2_desc = QLabel("إمكانية إضافة زيادة على الرواتب بنسبة مئوية أو بمبلغ ثابت لجميع العاملين أو لقسم محدد:")
        c2_desc.setStyleSheet("color: #64748b; font-size: 11px; border: none;")
        c2_layout.addWidget(c2_title)
        c2_layout.addWidget(c2_desc)

        bulk_grid = QGridLayout()
        bulk_grid.setSpacing(10)

        bulk_grid.addWidget(QLabel("نوع الزيادة:"), 0, 0)
        self.bulk_type_combo = QComboBox()
        self.bulk_type_combo.addItems(["نسبة مئوية (%)", "مبلغ ثابت (ج.م)"])
        self.bulk_type_combo.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px;")
        bulk_grid.addWidget(self.bulk_type_combo, 0, 1)

        bulk_grid.addWidget(QLabel("قيمة الزيادة:"), 0, 2)
        self.bulk_val_spin = QDoubleSpinBox()
        self.bulk_val_spin.setRange(0.1, 100000.0)
        self.bulk_val_spin.setValue(10.0)
        self.bulk_val_spin.setDecimals(2)
        self.bulk_val_spin.setStyleSheet("padding: 6px; font-weight: bold; border: 1.5px solid #cbd5e1; border-radius: 6px;")
        bulk_grid.addWidget(self.bulk_val_spin, 0, 3)

        bulk_grid.addWidget(QLabel("نطاق التطبيق:"), 1, 0)
        self.bulk_target_combo = QComboBox()
        self.bulk_target_combo.addItems(["👥 جميع العاملين بالمصنع", "🏢 قسم محدد فقط"])
        self.bulk_target_combo.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px;")
        bulk_grid.addWidget(self.bulk_target_combo, 1, 1)

        self.bulk_dept_combo = QComboBox()
        self.bulk_dept_combo.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px;")
        self.populate_dept_combo(self.bulk_dept_combo)
        self.bulk_dept_combo.setEnabled(False)
        bulk_grid.addWidget(self.bulk_dept_combo, 1, 2, 1, 2)

        self.bulk_target_combo.currentIndexChanged.connect(
            lambda idx: self.bulk_dept_combo.setEnabled(idx == 1)
        )

        bulk_grid.addWidget(QLabel("تطبيق الزيادة على:"), 2, 0)
        self.bulk_apply_to_combo = QComboBox()
        self.bulk_apply_to_combo.addItems([
            "💵 الراتب الأساسي فقط",
            "🎁 حافز الانتظام فقط",
            "🌟 الراتب الأساسي + حافز الانتظام معاً"
        ])
        self.bulk_apply_to_combo.setStyleSheet("padding: 6px; border: 1.5px solid #cbd5e1; border-radius: 6px;")
        bulk_grid.addWidget(self.bulk_apply_to_combo, 2, 1, 1, 3)

        c2_layout.addLayout(bulk_grid)

        bulk_btn_box = QHBoxLayout()
        bulk_btn_box.addStretch()

        btn_preview_inc = QPushButton("🔍 معاينة تأثير الزيادة والتكلفة")
        btn_preview_inc.setCursor(Qt.PointingHandCursor)
        btn_preview_inc.setStyleSheet("""
            QPushButton {
                background-color: #0284c7; color: white; padding: 8px 18px;
                font-weight: bold; border-radius: 6px; font-size: 12px; border: none;
            }
            QPushButton:hover { background-color: #0369a1; }
        """)
        btn_preview_inc.clicked.connect(self.preview_bulk_salary_increment_handler)
        bulk_btn_box.addWidget(btn_preview_inc)

        btn_apply_inc = QPushButton("🚀 تطبيق الزيادة فورياً")
        btn_apply_inc.setCursor(Qt.PointingHandCursor)
        btn_apply_inc.setStyleSheet("""
            QPushButton {
                background-color: #16a34a; color: white; padding: 8px 20px;
                font-weight: bold; border-radius: 6px; font-size: 12px; border: none;
            }
            QPushButton:hover { background-color: #15803d; }
        """)
        btn_apply_inc.clicked.connect(self.apply_bulk_salary_increment_handler)
        bulk_btn_box.addWidget(btn_apply_inc)

        c2_layout.addLayout(bulk_btn_box)
        container_layout.addWidget(card_bulk)

        # ── Card 3: Configurable Absence & Incentive Rules ─────────────────
        card_rules = QFrame()
        card_rules.setStyleSheet("""
            QFrame {
                background: white; border-radius: 10px; border: 1.5px solid #cbd5e1; padding: 16px;
            }
        """)
        c3_layout = QVBoxLayout(card_rules)
        c3_layout.setSpacing(12)

        c3_title = QLabel("🎯 تخصيص وضبط خطة وقواعد الغياب وحافز الانتظام")
        c3_title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        c3_title.setStyleSheet("color: #0f172a; border: none;")
        c3_desc = QLabel("حدد شروط ودرجات الخصم من حافز الانتظام والراتب الأساسي بدقة لكل شريحة غياب:")
        c3_desc.setStyleSheet("color: #64748b; font-size: 11px; border: none;")
        c3_layout.addWidget(c3_title)
        c3_layout.addWidget(c3_desc)

        base_rules_row = QHBoxLayout()
        base_rules_row.setSpacing(16)

        base_rules_row.addWidget(QLabel("🎁 حافز الانتظام الشهري الافتراضي:"))
        self.default_inc_spin = QDoubleSpinBox()
        self.default_inc_spin.setRange(0.0, 50000.0)
        self.default_inc_spin.setValue(cfg.get("default_incentive", 500.0))
        self.default_inc_spin.setDecimals(2)
        self.default_inc_spin.setSuffix(" ج.م")
        self.default_inc_spin.setStyleSheet("padding: 6px; font-weight: bold; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 120px;")
        base_rules_row.addWidget(self.default_inc_spin)

        base_rules_row.addWidget(QLabel("➗ قاسم أجر اليوم (عدد أيام قسمة الراتب):"))
        self.daily_divisor_spin = QDoubleSpinBox()
        self.daily_divisor_spin.setRange(1.0, 365.0)
        self.daily_divisor_spin.setValue(cfg.get("daily_divisor", 30.0))
        self.daily_divisor_spin.setDecimals(1)
        self.daily_divisor_spin.setSuffix(" يوم")
        self.daily_divisor_spin.setStyleSheet("padding: 6px; font-weight: bold; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 100px;")
        base_rules_row.addWidget(self.daily_divisor_spin)

        base_rules_row.addStretch()
        c3_layout.addLayout(base_rules_row)

        base_rules_row2 = QHBoxLayout()
        base_rules_row2.setSpacing(16)

        base_rules_row2.addWidget(QLabel("🌙 معامل ساعة الإضافي (الأيام العادية):"))
        self.ot_normal_mult_spin = QDoubleSpinBox()
        self.ot_normal_mult_spin.setRange(0.5, 10.0)
        self.ot_normal_mult_spin.setValue(float(cfg.get("overtime_normal_multiplier", 1.5)))
        self.ot_normal_mult_spin.setDecimals(2)
        self.ot_normal_mult_spin.setSingleStep(0.25)
        self.ot_normal_mult_spin.setSuffix(" ×")
        self.ot_normal_mult_spin.setStyleSheet("padding: 6px; font-weight: bold; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 100px; color: #7c3aed;")
        base_rules_row2.addWidget(self.ot_normal_mult_spin)

        base_rules_row2.addWidget(QLabel("🕌 معامل ساعة الإضافي (أيام الجمعة):"))
        self.ot_friday_mult_spin = QDoubleSpinBox()
        self.ot_friday_mult_spin.setRange(0.5, 10.0)
        self.ot_friday_mult_spin.setValue(float(cfg.get("overtime_friday_multiplier", 2.0)))
        self.ot_friday_mult_spin.setDecimals(2)
        self.ot_friday_mult_spin.setSingleStep(0.25)
        self.ot_friday_mult_spin.setSuffix(" ×")
        self.ot_friday_mult_spin.setStyleSheet("padding: 6px; font-weight: bold; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 100px; color: #0284c7;")
        base_rules_row2.addWidget(self.ot_friday_mult_spin)

        base_rules_row2.addStretch()
        c3_layout.addLayout(base_rules_row2)

        # Absence Tiers Table
        self.rules_table = QTableWidget()
        self.rules_table.setColumnCount(4)
        self.rules_table.setHorizontalHeaderLabels([
            "تدرج أيام الغياب", "🎁 مبلغ خصم حافز الانتظام (ج.م)", "📉 أيام خصم الأساسي", "📝 القاعدة والتوضيح"
        ])
        rhdr = self.rules_table.horizontalHeader()
        rhdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        rhdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        rhdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        rhdr.setSectionResizeMode(3, QHeaderView.Stretch)
        self.rules_table.verticalHeader().setVisible(False)
        self.rules_table.verticalHeader().setDefaultSectionSize(42)
        self.rules_table.setFixedHeight(250)
        self.rules_table.setStyleSheet("""
            QTableWidget { border: 1.5px solid #cbd5e1; border-radius: 8px; background: white; gridline-color: #e2e8f0; }
            QTableWidget::item { padding: 6px 10px; font-size: 13px; color: #0f172a; }
            QHeaderView::section { background-color: #0f172a; color: white; padding: 9px 8px; font-weight: bold; font-size: 12px; }
        """)

        self.populate_absence_rules_table(cfg.get("absence_rules", {}))
        c3_layout.addWidget(self.rules_table)

        rules_btn_box = QHBoxLayout()
        btn_reset_rules = QPushButton("🔄 استعادة القواعد الافتراضية")
        btn_reset_rules.setCursor(Qt.PointingHandCursor)
        btn_reset_rules.setStyleSheet("""
            QPushButton {
                background-color: #64748b; color: white; padding: 8px 16px;
                font-weight: bold; border-radius: 6px; font-size: 12px; border: none;
            }
            QPushButton:hover { background-color: #475569; }
        """)
        btn_reset_rules.clicked.connect(self.reset_default_payroll_rules_handler)
        rules_btn_box.addWidget(btn_reset_rules)

        rules_btn_box.addStretch()

        btn_save_rules = QPushButton("💾 حفظ قواعد الرواتب والغياب")
        btn_save_rules.setCursor(Qt.PointingHandCursor)
        btn_save_rules.setStyleSheet("""
            QPushButton {
                background-color: #16a34a; color: white; padding: 8px 22px;
                font-weight: bold; border-radius: 6px; font-size: 12px; border: none;
            }
            QPushButton:hover { background-color: #15803d; }
        """)
        btn_save_rules.clicked.connect(self.save_payroll_rules_handler)
        rules_btn_box.addWidget(btn_save_rules)

        c3_layout.addLayout(rules_btn_box)
        container_layout.addWidget(card_rules)

        # ── Card 4: Departments Shift Schedule Manager ─────────────────────
        card_shifts = QFrame()
        card_shifts.setStyleSheet("""
            QFrame {
                background: white; border-radius: 10px; border: 1.5px solid #cbd5e1; padding: 16px;
            }
        """)
        c4_layout = QVBoxLayout(card_shifts)
        c4_layout.setSpacing(12)

        c4_title = QLabel("🕒 مواعيد وساعات عمل الأقسام (ورديات الحضور والانصراف)")
        c4_title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        c4_title.setStyleSheet("color: #0f172a; border: none;")
        c4_desc = QLabel("تعديل مواعيد الحضور والانصراف وساعات العمل لجميع أقسام المصنع مع إمكانية الحفظ الفردي أو الجماعي:")
        c4_desc.setStyleSheet("color: #64748b; font-size: 11px; border: none;")
        c4_layout.addWidget(c4_title)
        c4_layout.addWidget(c4_desc)

        self.dept_shifts_table = QTableWidget()
        headers_shifts = ["ID", "اسم القسم", "🕒 وقت الحضور (البداية)", "🕑 وقت الانصراف (النهاية)", "⏱️ ساعات العمل اليومية", "💾 حفظ القسم"]
        self.dept_shifts_table.setColumnCount(len(headers_shifts))
        self.dept_shifts_table.setHorizontalHeaderLabels(headers_shifts)
        shdr = self.dept_shifts_table.horizontalHeader()
        shdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        shdr.setSectionResizeMode(1, QHeaderView.Stretch)
        shdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        shdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        shdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        shdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.dept_shifts_table.verticalHeader().setVisible(False)
        self.dept_shifts_table.verticalHeader().setDefaultSectionSize(44)
        self.dept_shifts_table.setFixedHeight(300)
        self.dept_shifts_table.setStyleSheet("""
            QTableWidget { border: 1.5px solid #cbd5e1; border-radius: 8px; background: white; gridline-color: #e2e8f0; }
            QTableWidget::item { padding: 6px 10px; font-size: 13px; color: #0f172a; }
            QHeaderView::section { background-color: #0f172a; color: white; padding: 9px 8px; font-weight: bold; font-size: 12px; }
        """)

        self.load_dept_shifts_batch_table()
        c4_layout.addWidget(self.dept_shifts_table)

        shifts_btn_box = QHBoxLayout()
        btn_refresh_shifts = QPushButton("🔄 إعادة تحميل مواعيد الأقسام")
        btn_refresh_shifts.setCursor(Qt.PointingHandCursor)
        btn_refresh_shifts.setStyleSheet("background: #e2e8f0; color: #1e293b; padding: 7px 16px; font-weight: bold; border-radius: 6px; border: none;")
        btn_refresh_shifts.clicked.connect(self.load_dept_shifts_batch_table)
        shifts_btn_box.addWidget(btn_refresh_shifts)

        shifts_btn_box.addStretch()

        btn_save_all_shifts = QPushButton("💾 حفظ مواعيد جميع الأقسام دفعة واحدة")
        btn_save_all_shifts.setCursor(Qt.PointingHandCursor)
        btn_save_all_shifts.setStyleSheet("""
            QPushButton {
                background-color: #16a34a; color: white; padding: 7px 22px;
                font-weight: bold; border-radius: 6px; font-size: 12px; border: none;
            }
            QPushButton:hover { background-color: #15803d; }
        """)
        btn_save_all_shifts.clicked.connect(self.save_all_dept_shifts_handler)
        shifts_btn_box.addWidget(btn_save_all_shifts)

        c4_layout.addLayout(shifts_btn_box)
        container_layout.addWidget(card_shifts)

        scroll.setWidget(container)
        tab_layout.addWidget(scroll)

        self.update_cycle_preview_label()
        self.cycle_start_spin.valueChanged.connect(self.update_cycle_preview_label)
        self.cycle_end_spin.valueChanged.connect(self.update_cycle_preview_label)
        self.cycle_type_std.toggled.connect(self.update_cycle_preview_label)
        self.cycle_type_custom.toggled.connect(self.update_cycle_preview_label)

    def update_cycle_preview_label(self):
        """تحديث النص التوضيحي لمعاينة دورة الشهر الحالية."""
        if not hasattr(self, "cycle_preview_lbl"):
            return
        now = datetime.now()
        y = now.year
        m = now.month
        is_custom = self.cycle_type_custom.isChecked()
        if not is_custom:
            days_in_m = calendar.monthrange(y, m)[1]
            self.cycle_preview_lbl.setText(f"ℹ️ معاينة دورة شهر {m} الحالية: من 01-{m:02d}-{y} إلى {days_in_m:02d}-{m:02d}-{y} (شهر تقويمي كامل)")
        else:
            s_day = self.cycle_start_spin.value()
            e_day = self.cycle_end_spin.value()
            prev_m = 12 if m == 1 else m - 1
            prev_y = y - 1 if m == 1 else y
            self.cycle_preview_lbl.setText(
                f"ℹ️ معاينة دورة شهر {m} الحالية: من {s_day:02d}-{prev_m:02d}-{prev_y} إلى {e_day:02d}-{m:02d}-{y}"
            )

    def populate_dept_combo(self, combo):
        """تعبئة القائمة المنسدلة بالأقسام المسجلة."""
        combo.clear()
        depts = get_departments()
        for d in depts:
            combo.addItem(f"🏢 {d[1]}", d[0])

    def populate_absence_rules_table(self, rules: dict):
        """تعبئة جدول قواعد الغياب وحوافز الانتظام."""
        self.rules_table.setRowCount(5)

        tier_data = [
            ("tier_0", "0 يوم غياب", rules.get("tier_0", {}).get("deduct_incentive", 0.0), rules.get("tier_0", {}).get("deduct_base_days", 0.0), "يستحق حافز الانتظام كاملاً وبدون خصم من الأساسي"),
            ("tier_1", "1 يوم غياب", rules.get("tier_1", {}).get("deduct_incentive", 100.0), rules.get("tier_1", {}).get("deduct_base_days", 0.0), "يُخصم من حافز الانتظام فقط ولا يُخصم من الأساسي"),
            ("tier_2", "2 يوم غياب", rules.get("tier_2", {}).get("deduct_incentive", 200.0), rules.get("tier_2", {}).get("deduct_base_days", 1.0), "يُخصم من حافز الانتظام ويُخصم أجر يوم واحد من الأساسي"),
            ("tier_3", "3 أيام غياب", rules.get("tier_3", {}).get("deduct_incentive", 300.0), rules.get("tier_3", {}).get("deduct_base_days", 2.0), "يُخصم من حافز الانتظام ويُخصم أجر يومين من الأساسي"),
            ("tier_4", "4+ أيام غياب", rules.get("tier_4", {}).get("deduct_incentive", 500.0), "أيام الغياب - 1", "يصبح الحافز 0 ويُخصم (أيام الغياب - 1) من الأساسي"),
        ]

        for r_idx, (key, title, inc_ded, base_ded, note) in enumerate(tier_data):
            self.rules_table.setRowHeight(r_idx, 40)

            # Col 0: Title
            t_item = QTableWidgetItem(title)
            t_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
            t_item.setTextAlignment(Qt.AlignCenter)
            t_item.setFlags(t_item.flags() & ~Qt.ItemIsEditable)
            self.rules_table.setItem(r_idx, 0, t_item)

            # Col 1: Incentive Deduction Spinbox
            inc_spin = QDoubleSpinBox()
            inc_spin.setRange(0.0, 50000.0)
            inc_spin.setDecimals(2)
            inc_spin.setValue(float(inc_ded))
            inc_spin.setSuffix(" ج.م")
            inc_spin.setStyleSheet("padding: 4px 8px; font-weight: bold; font-size: 12px; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 120px; min-height: 28px;")
            self.rules_table.setCellWidget(r_idx, 1, inc_spin)

            # Col 2: Base deduction days
            if isinstance(base_ded, (int, float)):
                base_spin = QDoubleSpinBox()
                base_spin.setRange(0.0, 31.0)
                base_spin.setDecimals(1)
                base_spin.setValue(float(base_ded))
                base_spin.setSuffix(" يوم")
                base_spin.setStyleSheet("padding: 4px 8px; font-weight: bold; font-size: 12px; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 110px; min-height: 28px;")
                self.rules_table.setCellWidget(r_idx, 2, base_spin)
            else:
                b_item = QTableWidgetItem(str(base_ded))
                b_item.setTextAlignment(Qt.AlignCenter)
                b_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                b_item.setForeground(QColor("#dc2626"))
                b_item.setFlags(b_item.flags() & ~Qt.ItemIsEditable)
                self.rules_table.setItem(r_idx, 2, b_item)

            # Col 3: Note
            n_item = QTableWidgetItem(note)
            n_item.setFont(QFont("Segoe UI", 9))
            n_item.setFlags(n_item.flags() & ~Qt.ItemIsEditable)
            self.rules_table.setItem(r_idx, 3, n_item)

    def load_dept_shifts_batch_table(self):
        """تحميل وعرض بيانات مواعيد وساعات عمل الأقسام."""
        if not hasattr(self, "dept_shifts_table") or self.dept_shifts_table is None:
            return

        depts = get_departments()
        self.dept_shifts_table.setRowCount(len(depts))

        for r_idx, d in enumerate(depts):
            dept_id = d[0]
            dept_name = d[1]
            s_time_str = d[4] if len(d) > 4 and d[4] else "08:00 ص"
            e_time_str = d[5] if len(d) > 5 and d[5] else "07:00 م"
            w_hours = float(d[6]) if len(d) > 6 and d[6] is not None else 11.0

            self.dept_shifts_table.setRowHeight(r_idx, 42)

            # 0. ID
            id_item = QTableWidgetItem(str(dept_id))
            id_item.setTextAlignment(Qt.AlignCenter)
            id_item.setFlags(id_item.flags() & ~Qt.ItemIsEditable)
            self.dept_shifts_table.setItem(r_idx, 0, id_item)

            # 1. Name
            name_item = QTableWidgetItem(dept_name)
            name_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.dept_shifts_table.setItem(r_idx, 1, name_item)

            # 2. Check-in Time
            s_edit = QTimeEdit()
            s_edit.setDisplayFormat("hh:mm AP")
            s_edit.setLocale(ARABIC_LOCALE)
            t_s = parse_time_str(s_time_str)
            s_edit.setTime(t_s if t_s.isValid() else QTime(8, 0))
            s_edit.setStyleSheet("padding: 4px 8px; font-weight: bold; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 120px; min-height: 28px;")
            self.dept_shifts_table.setCellWidget(r_idx, 2, s_edit)

            # 3. Check-out Time
            e_edit = QTimeEdit()
            e_edit.setDisplayFormat("hh:mm AP")
            e_edit.setLocale(ARABIC_LOCALE)
            t_e = parse_time_str(e_time_str)
            e_edit.setTime(t_e if t_e.isValid() else QTime(19, 0))
            e_edit.setStyleSheet("padding: 4px 8px; font-weight: bold; border: 1.5px solid #cbd5e1; border-radius: 6px; min-width: 120px; min-height: 28px;")
            self.dept_shifts_table.setCellWidget(r_idx, 3, e_edit)

            # 4. Work Hours Item
            h_item = QTableWidgetItem(f"{w_hours:.1f} س")
            h_item.setTextAlignment(Qt.AlignCenter)
            h_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
            h_item.setForeground(QColor("#2563eb"))
            h_item.setFlags(h_item.flags() & ~Qt.ItemIsEditable)
            self.dept_shifts_table.setItem(r_idx, 4, h_item)

            # Auto calculate hours on time change
            def make_update_h(row, s_w, e_w):
                def update():
                    wh = self.calculate_work_hours(s_w.time(), e_w.time())
                    item = self.dept_shifts_table.item(row, 4)
                    if item:
                        item.setText(f"{wh:.1f} س")
                return update

            updater = make_update_h(r_idx, s_edit, e_edit)
            s_edit.timeChanged.connect(updater)
            e_edit.timeChanged.connect(updater)

            # 5. Save Button for this dept
            save_b = QPushButton("💾 حفظ")
            save_b.setCursor(Qt.PointingHandCursor)
            save_b.setStyleSheet("""
                QPushButton {
                    background-color: #2563eb; color: white; padding: 5px 14px;
                    font-weight: bold; border-radius: 5px; font-size: 11px; border: none;
                }
                QPushButton:hover { background-color: #1d4ed8; }
            """)

            def make_save_dept(did, dname, s_w, e_w):
                def save():
                    st = format_time_ar(s_w.time())
                    et = format_time_ar(e_w.time())
                    wh = self.calculate_work_hours(s_w.time(), e_w.time())
                    cur_d = get_department_by_id(did)
                    code = cur_d.get("code", "") if cur_d else ""
                    notes = cur_d.get("notes", "") if cur_d else ""
                    res = update_department(did, dname, code, notes, st, et, wh)
                    if res.get("success"):
                        self.load_departments_table()
                        self.refresh_hr_views()
                        if hasattr(self, "show_toast"):
                            self.show_toast(f"✓ تم حفظ مواعيد قسم '{dname}' ({st} إلى {et} - {wh:.1f} س)")
                    else:
                        QMessageBox.critical(self, "خطأ", f"تعذر حفظ القسم: {res.get('error')}")
                return save

            save_b.clicked.connect(make_save_dept(dept_id, dept_name, s_edit, e_edit))
            self.dept_shifts_table.setCellWidget(r_idx, 5, save_b)

    def save_all_dept_shifts_handler(self):
        """حفظ مواعيد وساعات عمل جميع الأقسام دفعة واحدة."""
        if not hasattr(self, "dept_shifts_table") or self.dept_shifts_table is None:
            return

        saved_count = 0
        for r in range(self.dept_shifts_table.rowCount()):
            id_item = self.dept_shifts_table.item(r, 0)
            name_item = self.dept_shifts_table.item(r, 1)
            s_edit = self.dept_shifts_table.cellWidget(r, 2)
            e_edit = self.dept_shifts_table.cellWidget(r, 3)

            if id_item and name_item and s_edit and e_edit:
                did = int(id_item.text())
                dname = name_item.text()
                st = format_time_ar(s_edit.time())
                et = format_time_ar(e_edit.time())
                wh = self.calculate_work_hours(s_edit.time(), e_edit.time())

                cur_d = get_department_by_id(did)
                code = cur_d.get("code", "") if cur_d else ""
                notes = cur_d.get("notes", "") if cur_d else ""
                update_department(did, dname, code, notes, st, et, wh)
                saved_count += 1

        self.load_departments_table()
        self.refresh_hr_views()
        if hasattr(self, "show_toast"):
            self.show_toast(f"✓ تم حفظ مواعيد جميع الأقسام ({saved_count} قسم) دفعة واحدة بنجاح!")

    def save_cycle_settings_handler(self):
        """حفظ إعدادات دورة الشهر المالي."""
        is_custom = self.cycle_type_custom.isChecked()
        cycle_type = "custom_day" if is_custom else "calendar"
        start_day = self.cycle_start_spin.value()
        end_day = self.cycle_end_spin.value()

        cfg = {
            "cycle_type": cycle_type,
            "start_day": start_day,
            "end_day": end_day,
        }
        if save_payroll_config(cfg):
            self.refresh_hr_views()
            if hasattr(self, "show_toast"):
                self.show_toast("✓ تم حفظ إعدادات دورة الشهر المالي بنجاح")
            else:
                QMessageBox.information(self, "نجاح", "تم حفظ إعدادات دورة الشهر المالي بنجاح.")
        else:
            QMessageBox.critical(self, "خطأ", "تعذر حفظ إعدادات دورة الشهر.")

    def preview_bulk_salary_increment_handler(self):
        """معاينة تكلفة الزيادة الجماعية قبل تطبيقها."""
        inc_type = "percentage" if self.bulk_type_combo.currentIndex() == 0 else "amount"
        value = self.bulk_val_spin.value()
        target_type = "all" if self.bulk_target_combo.currentIndex() == 0 else "department"
        dept_id = self.bulk_dept_combo.currentData() if target_type == "department" else None
        dept_name = self.bulk_dept_combo.currentText() if target_type == "department" else "جميع العاملين بالمصنع"
        apply_to_idx = self.bulk_apply_to_combo.currentIndex()
        apply_to = "base_salary" if apply_to_idx == 0 else ("attendance_incentive" if apply_to_idx == 1 else "both")

        conn = get_db_connection()
        c = conn.cursor()
        where_clause = "WHERE is_active=1"
        params = []
        if target_type == "department" and dept_id:
            where_clause += " AND department_id=?"
            params.append(dept_id)

        c.execute(f"SELECT COUNT(*), COALESCE(SUM(base_salary), 0.0), COALESCE(SUM(attendance_incentive), 0.0) FROM employees {where_clause}", params)
        row = c.fetchone()
        conn.close()

        count, cur_base, cur_inc = row[0] or 0, row[1] or 0.0, row[2] or 0.0
        if count == 0:
            QMessageBox.warning(self, "تنبيه", "لا يوجد عمال مطابقين للنطاق المحدد.")
            return

        new_base = cur_base
        new_inc = cur_inc

        if apply_to in ("base_salary", "both"):
            if inc_type == "percentage":
                new_base = cur_base * (1.0 + (value / 100.0))
            else:
                new_base = cur_base + (value * count)

        if apply_to in ("attendance_incentive", "both"):
            if inc_type == "percentage":
                new_inc = cur_inc * (1.0 + (value / 100.0))
            else:
                new_inc = cur_inc + (value * count)

        base_diff = new_base - cur_base
        inc_diff = new_inc - cur_inc
        total_monthly_diff = base_diff + inc_diff

        type_str = f"{value:.2f}%" if inc_type == "percentage" else f"{value:,.2f} ج.م لكل عامل"
        apply_to_str = self.bulk_apply_to_combo.currentText()

        msg = f"""
        <h3 style='color: #0f172a;'>📊 معاينة تأثير الزيادة الجماعية للرواتب</h3>
        <hr>
        <p><b>النطاق المستهدف:</b> {dept_name}</p>
        <p><b>عدد العمال المشمولين:</b> <span style='color: #2563eb; font-weight: bold;'>{count} عامل</span></p>
        <p><b>نوع وقيمة الزيادة:</b> <span style='color: #16a34a; font-weight: bold;'>{type_str}</span></p>
        <p><b>البند المشمول:</b> {apply_to_str}</p>
        <br>
        <table border='1' cellspacing='0' cellpadding='6' style='border-collapse: collapse; width: 100%; border-color: #cbd5e1;'>
            <tr style='background: #f1f5f9;'>
                <th>البيان</th><th>الوضع الحالي</th><th>بعد الزيادة المقترحة</th><th>الفرق (الزيادة)</th>
            </tr>
            <tr>
                <td><b>إجمالي الأساسي</b></td>
                <td>{cur_base:,.2f} ج.م</td>
                <td style='color: #059669;'><b>{new_base:,.2f} ج.م</b></td>
                <td style='color: #2563eb;'>+{base_diff:,.2f} ج.م</td>
            </tr>
            <tr>
                <td><b>إجمالي حوافز الانتظام</b></td>
                <td>{cur_inc:,.2f} ج.م</td>
                <td style='color: #059669;'><b>{new_inc:,.2f} ج.م</b></td>
                <td style='color: #2563eb;'>+{inc_diff:,.2f} ج.م</td>
            </tr>
            <tr style='background: #eff6ff; font-weight: bold;'>
                <td><b>التكلفة الشهرية الإجمالية</b></td>
                <td>{(cur_base + cur_inc):,.2f} ج.م</td>
                <td style='color: #16a34a;'>{(new_base + new_inc):,.2f} ج.م</td>
                <td style='color: #dc2626;'>+{total_monthly_diff:,.2f} ج.م شهرياً</td>
            </tr>
        </table>
        """
        dlg = QMessageBox(self)
        dlg.setWindowTitle("معاينة الزيادة والتكلفة")
        dlg.setTextFormat(Qt.RichText)
        dlg.setText(msg)
        dlg.setStandardButtons(QMessageBox.Ok)
        dlg.exec()

    def apply_bulk_salary_increment_handler(self):
        """تطبيق الزيادة الجماعية للرواتب فورياً بعد التأكيد."""
        inc_type = "percentage" if self.bulk_type_combo.currentIndex() == 0 else "amount"
        value = self.bulk_val_spin.value()
        target_type = "all" if self.bulk_target_combo.currentIndex() == 0 else "department"
        dept_id = self.bulk_dept_combo.currentData() if target_type == "department" else None
        dept_name = self.bulk_dept_combo.currentText() if target_type == "department" else "جميع العاملين بالمصنع"
        apply_to_idx = self.bulk_apply_to_combo.currentIndex()
        apply_to = "base_salary" if apply_to_idx == 0 else ("attendance_incentive" if apply_to_idx == 1 else "both")

        type_str = f"{value:.2f}%" if inc_type == "percentage" else f"{value:,.2f} ج.م لكل عامل"
        confirm_msg = f"هل أنت متأكد من تطبيق زيادة قدرها <b>{type_str}</b> على <b>{dept_name}</b>؟<br><br>⚠️ سيتم تحديث الرواتب الأساسية وحوافز الانتظام وأجر اليوم لكافة العمال المشمولين فورياً."

        reply = QMessageBox.question(
            self, "تأكيد تطبيق زيادة الرواتب", confirm_msg,
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            res = apply_bulk_salary_increment(target_type, dept_id, inc_type, value, apply_to)
            if res.get("success"):
                self.load_departments_table()
                self.refresh_hr_views()
                if hasattr(self, "load_payroll_table"):
                    self.load_payroll_table()
                QMessageBox.information(
                    self, "نجاح تطبيق الزيادة",
                    f"🎉 تم تطبيق الزيادة بنجاح على <b>{res['affected_count']}</b> عامل!<br><br>"
                    f"• إجمالي الرواتب الأساسية الجديد: <b>{res['new_total_base']:,.2f} ج.م</b> (زيادة +{res['base_increase']:,.2f} ج.م)<br>"
                    f"• إجمالي حوافز الانتظام الجديد: <b>{res['new_total_inc']:,.2f} ج.م</b> (زيادة +{res['inc_increase']:,.2f} ج.م)"
                )
            else:
                QMessageBox.critical(self, "خطأ", f"تعذر تطبيق الزيادة: {res.get('error')}")

    def save_payroll_rules_handler(self):
        """حفظ قواعد الغياب وحافز الانتظام ومضاعفات الإضافي."""
        default_inc = self.default_inc_spin.value()
        daily_div = self.daily_divisor_spin.value()
        ot_normal_m = self.ot_normal_mult_spin.value() if hasattr(self, "ot_normal_mult_spin") else 1.5
        ot_friday_m = self.ot_friday_mult_spin.value() if hasattr(self, "ot_friday_mult_spin") else 2.0

        # Read from table
        t0_inc = self.rules_table.cellWidget(0, 1).value() if self.rules_table.cellWidget(0, 1) else 0.0
        t0_base = self.rules_table.cellWidget(0, 2).value() if self.rules_table.cellWidget(0, 2) else 0.0

        t1_inc = self.rules_table.cellWidget(1, 1).value() if self.rules_table.cellWidget(1, 1) else 100.0
        t1_base = self.rules_table.cellWidget(1, 2).value() if self.rules_table.cellWidget(1, 2) else 0.0

        t2_inc = self.rules_table.cellWidget(2, 1).value() if self.rules_table.cellWidget(2, 1) else 200.0
        t2_base = self.rules_table.cellWidget(2, 2).value() if self.rules_table.cellWidget(2, 2) else 1.0

        t3_inc = self.rules_table.cellWidget(3, 1).value() if self.rules_table.cellWidget(3, 1) else 300.0
        t3_base = self.rules_table.cellWidget(3, 2).value() if self.rules_table.cellWidget(3, 2) else 2.0

        t4_inc = self.rules_table.cellWidget(4, 1).value() if self.rules_table.cellWidget(4, 1) else 500.0

        cfg = {
            "default_incentive": default_inc,
            "daily_divisor": daily_div,
            "overtime_normal_multiplier": ot_normal_m,
            "overtime_friday_multiplier": ot_friday_m,
            "absence_rules": {
                "tier_0": {"deduct_incentive": t0_inc, "deduct_base_days": t0_base},
                "tier_1": {"deduct_incentive": t1_inc, "deduct_base_days": t1_base},
                "tier_2": {"deduct_incentive": t2_inc, "deduct_base_days": t2_base},
                "tier_3": {"deduct_incentive": t3_inc, "deduct_base_days": t3_base},
                "tier_4": {"deduct_incentive": t4_inc, "deduct_base_rule": "days_minus_1", "deduct_base_fixed_days": 0.0},
            }
        }

        if save_payroll_config(cfg):
            if hasattr(self, "load_payroll_table"):
                self.load_payroll_table()
            if hasattr(self, "show_toast"):
                self.show_toast("✓ تم حفظ وتطبيق قواعد الرواتب والغياب ومضاعفات الإضافي بنجاح")
            else:
                QMessageBox.information(self, "نجاح", "تم حفظ وتطبيق قواعد الرواتب والغياب ومضاعفات الإضافي بنجاح.")
        else:
            QMessageBox.critical(self, "خطأ", "تعذر حفظ قواعد الرواتب.")

    def reset_default_payroll_rules_handler(self):
        """استعادة القواعد الافتراضية للغياب وحافز الانتظام والإضافي."""
        reply = QMessageBox.question(
            self, "تأكيد استعادة الافتراضيات",
            "هل تريد استعادة قواعد الرواتب والغياب والإضافي القياسية المصنعية؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            default_cfg = {
                "default_incentive": 500.0,
                "daily_divisor": 30.0,
                "overtime_normal_multiplier": 1.5,
                "overtime_friday_multiplier": 2.0,
                "absence_rules": {
                    "tier_0": {"deduct_incentive": 0.0, "deduct_base_days": 0.0},
                    "tier_1": {"deduct_incentive": 100.0, "deduct_base_days": 0.0},
                    "tier_2": {"deduct_incentive": 200.0, "deduct_base_days": 1.0},
                    "tier_3": {"deduct_incentive": 300.0, "deduct_base_days": 2.0},
                    "tier_4": {"deduct_incentive": 500.0, "deduct_base_rule": "days_minus_1", "deduct_base_fixed_days": 0.0},
                }
            }
            self.default_inc_spin.setValue(500.0)
            self.daily_divisor_spin.setValue(30.0)
            if hasattr(self, "ot_normal_mult_spin"):
                self.ot_normal_mult_spin.setValue(1.5)
            if hasattr(self, "ot_friday_mult_spin"):
                self.ot_friday_mult_spin.setValue(2.0)
            self.populate_absence_rules_table(default_cfg["absence_rules"])
            save_payroll_config(default_cfg)
            if hasattr(self, "load_payroll_table"):
                self.load_payroll_table()
            if hasattr(self, "show_toast"):
                self.show_toast("✓ تم استعادة القواعد القياسية للرواتب والغياب والإضافي")

    def refresh_hr_views(self):
        """تحديث جميع القوائم المنسدلة والجداول في شؤون العاملين."""
        for combo in getattr(self, "hr_employee_combos", []):
            if combo:
                self.populate_employee_combo(combo)
        if hasattr(self, "att_dept_filter"):
            self.populate_dept_filter_combo(self.att_dept_filter)
        if hasattr(self, "payroll_dept_filter"):
            self.populate_dept_filter_combo(self.payroll_dept_filter)
        if hasattr(self, "bulk_dept_combo"):
            self.populate_dept_combo(self.bulk_dept_combo)
        if hasattr(self, "load_attendance"):
            self.load_attendance()
        for fv in getattr(self, "hr_financial_views", []):
            self.load_financial(fv["table"], fv["table_widget"], fv["sum_label"])
        if hasattr(self, "load_departments_table"):
            self.load_departments_table()
        if hasattr(self, "load_dept_shifts_batch_table"):
            self.load_dept_shifts_batch_table()
        if hasattr(self, "load_payroll_table"):
            self.load_payroll_table()
        if hasattr(self, "load_monthly_advances_table"):
            self.load_monthly_advances_table()
        if hasattr(self, "load_long_term_loans_table"):
            self.load_long_term_loans_table()

