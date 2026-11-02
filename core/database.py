import sqlite3
import os as _os
from core.config import DB_NAME, _BASE_DIR


def get_db_connection():
    """Open a SQLite connection with performance pragmas pre-applied."""
    conn = sqlite3.connect(DB_NAME)
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    conn.execute("PRAGMA temp_store = MEMORY;")
    conn.execute("PRAGMA cache_size = -10000;")
    return conn


def init_database():
    """تهيئة وتجهيز جداول قاعدة بيانات نظام شؤون العاملين."""
    _os.makedirs(_os.path.dirname(DB_NAME), exist_ok=True)
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("PRAGMA foreign_keys = ON;")
    cursor.execute("PRAGMA journal_mode = WAL;")
    cursor.execute("PRAGMA synchronous = NORMAL;")
    cursor.execute("PRAGMA cache_size = -10000;")
    cursor.execute("PRAGMA temp_store = MEMORY;")

    # ── جداول شؤون العاملين والأقسام ───────────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS departments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            code TEXT DEFAULT '',
            start_time TEXT DEFAULT '08:00 ص',
            end_time TEXT DEFAULT '07:00 م',
            default_work_hours REAL DEFAULT 11.0,
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Ensure departments columns migration
    cursor.execute("PRAGMA table_info(departments)")
    dept_cols = {col[1] for col in cursor.fetchall()}
    if "start_time" not in dept_cols:
        cursor.execute("ALTER TABLE departments ADD COLUMN start_time TEXT DEFAULT '08:00 ص'")
    if "end_time" not in dept_cols:
        cursor.execute("ALTER TABLE departments ADD COLUMN end_time TEXT DEFAULT '07:00 م'")
    if "default_work_hours" not in dept_cols:
        cursor.execute("ALTER TABLE departments ADD COLUMN default_work_hours REAL DEFAULT 11.0")

    # Set specific shift times requested by user for key departments if default
    cursor.execute("""
        UPDATE departments 
        SET start_time='08:00 ص', end_time='06:00 م', default_work_hours=10.0 
        WHERE name LIKE '%التجميع%' AND (default_work_hours = 11.0 OR default_work_hours IS NULL)
    """)
    cursor.execute("""
        UPDATE departments 
        SET start_time='08:00 ص', end_time='07:00 م', default_work_hours=11.0 
        WHERE name LIKE '%الحقن%'
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            code TEXT DEFAULT '',
            department_id INTEGER,
            base_salary REAL DEFAULT 0.0,
            attendance_incentive REAL DEFAULT 500.0,
            daily_wage REAL DEFAULT 0.0,
            work_hours REAL DEFAULT 11.0,
            hire_date TEXT DEFAULT '',
            birth_date TEXT DEFAULT '',
            national_id TEXT DEFAULT '',
            insurance_status TEXT DEFAULT 'غير مؤمن',
            insurance_no TEXT DEFAULT '',
            marital_status TEXT DEFAULT 'أعزب',
            phone TEXT DEFAULT '',
            has_relatives TEXT DEFAULT 'لا',
            relative_name TEXT DEFAULT '',
            relative_relation TEXT DEFAULT '',
            address TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            is_active INTEGER DEFAULT 1,
            FOREIGN KEY (department_id) REFERENCES departments(id) ON DELETE SET NULL
        )
    """)

    # Ensure employees columns migration
    cursor.execute("PRAGMA table_info(employees)")
    emp_cols = {col[1] for col in cursor.fetchall()}
    if "code" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN code TEXT DEFAULT ''")
    if "department_id" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN department_id INTEGER REFERENCES departments(id) ON DELETE SET NULL")
    if "base_salary" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN base_salary REAL DEFAULT 0.0")
    if "attendance_incentive" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN attendance_incentive REAL DEFAULT 500.0")
    if "daily_wage" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN daily_wage REAL DEFAULT 0.0")
    if "work_hours" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN work_hours REAL DEFAULT 11.0")
    if "hire_date" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN hire_date TEXT DEFAULT ''")
    if "birth_date" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN birth_date TEXT DEFAULT ''")
    if "national_id" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN national_id TEXT DEFAULT ''")
    if "insurance_status" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN insurance_status TEXT DEFAULT 'غير مؤمن'")
    if "insurance_no" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN insurance_no TEXT DEFAULT ''")
    if "marital_status" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN marital_status TEXT DEFAULT 'أعزب'")
    if "phone" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN phone TEXT DEFAULT ''")
    if "has_relatives" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN has_relatives TEXT DEFAULT 'لا'")
    if "relative_name" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN relative_name TEXT DEFAULT ''")
    if "relative_relation" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN relative_relation TEXT DEFAULT ''")
    if "address" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN address TEXT DEFAULT ''")
    if "notes" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN notes TEXT DEFAULT ''")
    if "is_active" not in emp_cols:
        cursor.execute("ALTER TABLE employees ADD COLUMN is_active INTEGER DEFAULT 1")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            employee_id INTEGER,
            status TEXT DEFAULT 'حضور',
            check_in_time TEXT DEFAULT '08:00 ص',
            check_out_time TEXT DEFAULT '07:00 م',
            work_hours REAL DEFAULT 11.0,
            notes TEXT DEFAULT '',
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
        )
    """)

    # Ensure attendance columns
    cursor.execute("PRAGMA table_info(attendance)")
    att_cols = {col[1] for col in cursor.fetchall()}
    if "check_in_time" not in att_cols:
        cursor.execute("ALTER TABLE attendance ADD COLUMN check_in_time TEXT DEFAULT '08:00 ص'")
    if "check_out_time" not in att_cols:
        cursor.execute("ALTER TABLE attendance ADD COLUMN check_out_time TEXT DEFAULT '07:00 م'")
    if "work_hours" not in att_cols:
        cursor.execute("ALTER TABLE attendance ADD COLUMN work_hours REAL DEFAULT 11.0")
    if "notes" not in att_cols:
        cursor.execute("ALTER TABLE attendance ADD COLUMN notes TEXT DEFAULT ''")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS deductions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            employee_id INTEGER,
            amount REAL,
            reason TEXT,
            notes TEXT,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS incentives (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            employee_id INTEGER,
            amount REAL,
            reason TEXT,
            notes TEXT,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS overtime (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            employee_id INTEGER,
            start_time TEXT,
            end_time TEXT,
            notes TEXT,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS system_settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            description TEXT DEFAULT '',
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Initialize default settings if not exists
    default_settings = [
        ("payroll_cycle_type", "custom_day", "نوع دورة الشهر المالي: calendar (تقويمي) أو custom_day (مخصص)"),
        ("payroll_cycle_start_day", "29", "يوم بداية الشهر المالي في الدورة المخصصة"),
        ("payroll_cycle_end_day", "28", "يوم نهاية الشهر المالي في الدورة المخصصة"),
        ("default_attendance_incentive", "500.0", "قيمة حافز الانتظام الشهري الافتراضي"),
        ("daily_wage_divisor", "30.0", "عدد أيام قسمة الراتب الأساسي لحساب أجر اليوم"),
        ("absence_tier_0_incentive_ded", "0.0", "خصم الحافز لغياب 0 يوم"),
        ("absence_tier_0_base_ded_days", "0.0", "أيام خصم الأساسي لغياب 0 يوم"),
        ("absence_tier_1_incentive_ded", "100.0", "خصم الحافز لغياب 1 يوم"),
        ("absence_tier_1_base_ded_days", "0.0", "أيام خصم الأساسي لغياب 1 يوم"),
        ("absence_tier_2_incentive_ded", "200.0", "خصم الحافز لغياب 2 يوم"),
        ("absence_tier_2_base_ded_days", "1.0", "أيام خصم الأساسي لغياب 2 يوم"),
        ("absence_tier_3_incentive_ded", "300.0", "خصم الحافز لغياب 3 أيام"),
        ("absence_tier_3_base_ded_days", "2.0", "أيام خصم الأساسي لغياب 3 أيام"),
        ("absence_tier_4_incentive_ded", "500.0", "خصم الحافز لغياب 4+ أيام"),
        ("absence_tier_4_base_rule", "days_minus_1", "قاعدة خصم الأساسي لغياب 4+ أيام (days_minus_1 أو fixed_days)"),
        ("absence_tier_4_base_ded_days", "0.0", "أيام خصم الأساسي الثابتة لغياب 4+ أيام في حال fixed_days"),
        ("overtime_normal_multiplier", "1.5", "معامل ضرب ساعة الإضافي في الأيام العادية (1.5)"),
        ("overtime_friday_multiplier", "2.0", "معامل ضرب ساعة الإضافي في أيام الجمعة والعطلات (2.0)"),
    ]

    for s_key, s_val, s_desc in default_settings:
        cursor.execute("""
            INSERT OR IGNORE INTO system_settings (key, value, description)
            VALUES (?, ?, ?)
        """, (s_key, s_val, s_desc))

    # Normalize AM/PM in attendance and overtime to Arabic ص / م
    cursor.execute("UPDATE attendance SET check_in_time = REPLACE(REPLACE(check_in_time, 'AM', 'ص'), 'PM', 'م') WHERE check_in_time LIKE '%AM%' OR check_in_time LIKE '%PM%'")
    cursor.execute("UPDATE attendance SET check_out_time = REPLACE(REPLACE(check_out_time, 'AM', 'ص'), 'PM', 'م') WHERE check_out_time LIKE '%AM%' OR check_out_time LIKE '%PM%'")
    cursor.execute("UPDATE overtime SET start_time = REPLACE(REPLACE(start_time, 'AM', 'ص'), 'PM', 'م') WHERE start_time LIKE '%AM%' OR start_time LIKE '%PM%'")
    cursor.execute("UPDATE overtime SET end_time = REPLACE(REPLACE(end_time, 'AM', 'ص'), 'PM', 'م') WHERE end_time LIKE '%AM%' OR end_time LIKE '%PM%'")
    cursor.execute("UPDATE attendance SET status = 'غياب' WHERE status = 'إجازة'")

    # جدول الإعدادات
    cursor.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")

    # ── جداول السلف الشهرية والسلف طويلة الأجل ────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS monthly_advances (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            employee_id INTEGER NOT NULL,
            amount REAL NOT NULL,
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS long_term_loans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            employee_id INTEGER NOT NULL,
            total_amount REAL NOT NULL,
            installment_amount REAL NOT NULL,
            start_date TEXT NOT NULL,
            remaining_amount REAL NOT NULL,
            status TEXT DEFAULT 'نشطة',
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS loan_installments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loan_id INTEGER NOT NULL,
            employee_id INTEGER NOT NULL,
            amount REAL NOT NULL,
            date TEXT NOT NULL,
            payroll_month INTEGER DEFAULT 0,
            payroll_year INTEGER DEFAULT 0,
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (loan_id) REFERENCES long_term_loans(id) ON DELETE CASCADE,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS credited_leaves (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            employee_id INTEGER NOT NULL,
            days REAL NOT NULL DEFAULT 1.0,
            reason TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
        )
    """)

    # الفهارس لسرعة الاستعلام
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_attendance_date ON attendance(date);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_attendance_emp ON attendance(employee_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_deductions_date ON deductions(date);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_incentives_date ON incentives(date);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_overtime_date ON overtime(date);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_employees_dept ON employees(department_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_employees_code ON employees(code);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_monthly_advances_date ON monthly_advances(date);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_monthly_advances_emp ON monthly_advances(employee_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_long_term_loans_emp ON long_term_loans(employee_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_loan_installments_loan ON loan_installments(loan_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_credited_leaves_date ON credited_leaves(date);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_credited_leaves_emp ON credited_leaves(employee_id);")

    conn.commit()
    conn.close()


# ── دوال إدارة الأقسام ───────────────────────────────────────────────────

def get_departments():
    """الحصول على قائمة بجميع الأقسام مرتبة حسب الاسم."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("""
        SELECT id, name, code, notes, 
               COALESCE(start_time, '08:00 ص'), 
               COALESCE(end_time, '07:00 م'), 
               COALESCE(default_work_hours, 11.0)
        FROM departments 
        ORDER BY name ASC
    """)
    rows = c.fetchall()
    conn.close()
    return rows


def get_department_by_id(dept_id: int):
    """الحصول على بيانات قسم محدد."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("""
        SELECT id, name, code, notes,
               COALESCE(start_time, '08:00 ص'), 
               COALESCE(end_time, '07:00 م'), 
               COALESCE(default_work_hours, 11.0)
        FROM departments WHERE id=?
    """, (dept_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return {
            "id": row[0],
            "name": row[1],
            "code": row[2] or "",
            "notes": row[3] or "",
            "start_time": row[4] or "08:00 ص",
            "end_time": row[5] or "07:00 م",
            "default_work_hours": row[6] or 11.0
        }
    return None


def add_department(name: str, code: str = "", notes: str = "",
                   start_time: str = "08:00 ص", end_time: str = "07:00 م",
                   default_work_hours: float = 11.0) -> dict:
    """إضافة قسم جديد إلى قاعدة البيانات مع تحديد مواعيد العمل."""
    clean_name = name.strip()
    if not clean_name:
        return {"success": False, "error": "اسم القسم لا يمكن أن يكون فارغاً."}
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            INSERT INTO departments (name, code, notes, start_time, end_time, default_work_hours) 
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            clean_name,
            (code or "").strip(),
            (notes or "").strip(),
            (start_time or "08:00 ص").strip(),
            (end_time or "07:00 م").strip(),
            float(default_work_hours or 11.0)
        ))
        dept_id = c.lastrowid
        conn.commit()
        conn.close()
        return {"success": True, "id": dept_id}
    except Exception as e:
        return {"success": False, "error": str(e)}


def update_department(dept_id: int, name: str, code: str = "", notes: str = "",
                      start_time: str = "08:00 ص", end_time: str = "07:00 م",
                      default_work_hours: float = 11.0) -> dict:
    """تعديل بيانات قسم قائم بما في ذلك مواعيد العمل."""
    clean_name = name.strip()
    if not clean_name:
        return {"success": False, "error": "اسم القسم لا يمكن أن يكون فارغاً."}
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            UPDATE departments 
            SET name=?, code=?, notes=?, start_time=?, end_time=?, default_work_hours=? 
            WHERE id=?
        """, (
            clean_name,
            (code or "").strip(),
            (notes or "").strip(),
            (start_time or "08:00 ص").strip(),
            (end_time or "07:00 م").strip(),
            float(default_work_hours or 11.0),
            dept_id
        ))
        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def delete_department(dept_id: int) -> dict:
    """حذف قسم وإلغاء ارتباط العمال به."""
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("UPDATE employees SET department_id=NULL WHERE department_id=?", (dept_id,))
        c.execute("DELETE FROM departments WHERE id=?", (dept_id,))
        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_all_departments_with_stats():
    """
    استرجاع الأقسام مع إحصائيات عدد العمال المرتبطين وإجمالي الرواتب الأساسية ومواعيد العمل.
    """
    conn = get_db_connection()
    c = conn.cursor()
    query = """
        SELECT 
            d.id, 
            d.name, 
            COALESCE(d.code, ''), 
            COALESCE(d.notes, ''),
            COUNT(e.id) AS worker_count,
            COALESCE(SUM(e.base_salary), 0.0) AS total_salary,
            COALESCE(d.start_time, '08:00 ص'),
            COALESCE(d.end_time, '07:00 م'),
            COALESCE(d.default_work_hours, 11.0)
        FROM departments d
        LEFT JOIN employees e ON d.id = e.department_id AND e.is_active = 1
        GROUP BY d.id, d.name, d.code, d.notes, d.start_time, d.end_time, d.default_work_hours
        ORDER BY d.name ASC
    """
    c.execute(query)
    results = []
    for row in c.fetchall():
        results.append({
            "id": row[0],
            "name": row[1],
            "code": row[2],
            "notes": row[3],
            "worker_count": row[4],
            "employee_count": row[4],
            "total_salary": row[5],
            "total_base_salary": row[5],
            "start_time": row[6] or "08:00 ص",
            "end_time": row[7] or "07:00 م",
            "default_work_hours": row[8] or 11.0
        })
    conn.close()
    return results


# ── دوال الإعدادات العامة وقواعد الرواتب ────────────────────────────────────

def get_setting(key: str, default: str = None) -> str:
    """الحصول على قيمة إعداد من جدول الإعدادات العامة."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT value FROM system_settings WHERE key=?", (key,))
    row = c.fetchone()
    conn.close()
    if row and row[0] is not None:
        return row[0]
    return default


def set_setting(key: str, value: str, description: str = "") -> bool:
    """تعيين أو تحديث قيمة إعداد في جدول الإعدادات العامة."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("""
        INSERT INTO system_settings (key, value, description, updated_at)
        VALUES (?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(key) DO UPDATE SET 
            value=excluded.value, 
            description=COALESCE(NULLIF(excluded.description, ''), system_settings.description),
            updated_at=CURRENT_TIMESTAMP
    """, (key, str(value), description))
    conn.commit()
    conn.close()
    return True


def get_payroll_config() -> dict:
    """استرجاع كافة إعدادات وقواعد الرواتب وحوافز الانتظام ودورة الشهر."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT key, value FROM system_settings")
    settings = dict(c.fetchall())
    conn.close()

    cycle_type = settings.get("payroll_cycle_type", "calendar")
    start_day = int(settings.get("payroll_cycle_start_day", "1"))
    end_day = int(settings.get("payroll_cycle_end_day", "30"))
    default_inc = float(settings.get("default_attendance_incentive", "500.0"))
    daily_divisor = float(settings.get("daily_wage_divisor", "30.0"))

    tier_0_inc = float(settings.get("absence_tier_0_incentive_ded", "0.0"))
    tier_0_base = float(settings.get("absence_tier_0_base_ded_days", "0.0"))

    tier_1_inc = float(settings.get("absence_tier_1_incentive_ded", "100.0"))
    tier_1_base = float(settings.get("absence_tier_1_base_ded_days", "0.0"))

    tier_2_inc = float(settings.get("absence_tier_2_incentive_ded", "200.0"))
    tier_2_base = float(settings.get("absence_tier_2_base_ded_days", "1.0"))

    tier_3_inc = float(settings.get("absence_tier_3_incentive_ded", "300.0"))
    tier_3_base = float(settings.get("absence_tier_3_base_ded_days", "2.0"))

    tier_4_inc = float(settings.get("absence_tier_4_incentive_ded", "500.0"))
    tier_4_rule = settings.get("absence_tier_4_base_rule", "days_minus_1")
    tier_4_base_fixed = float(settings.get("absence_tier_4_base_ded_days", "0.0"))

    ot_normal_mult = float(settings.get("overtime_normal_multiplier", "1.5"))
    ot_friday_mult = float(settings.get("overtime_friday_multiplier", "2.0"))

    return {
        "cycle_type": cycle_type,
        "start_day": start_day,
        "end_day": end_day,
        "default_incentive": default_inc,
        "daily_divisor": daily_divisor,
        "overtime_normal_multiplier": ot_normal_mult,
        "overtime_friday_multiplier": ot_friday_mult,
        "absence_rules": {
            "tier_0": {"name": "0 يوم غياب", "deduct_incentive": tier_0_inc, "deduct_base_days": tier_0_base},
            "tier_1": {"name": "1 يوم غياب", "deduct_incentive": tier_1_inc, "deduct_base_days": tier_1_base},
            "tier_2": {"name": "2 يوم غياب", "deduct_incentive": tier_2_inc, "deduct_base_days": tier_2_base},
            "tier_3": {"name": "3 أيام غياب", "deduct_incentive": tier_3_inc, "deduct_base_days": tier_3_base},
            "tier_4": {
                "name": "4+ أيام غياب",
                "deduct_incentive": tier_4_inc,
                "deduct_base_rule": tier_4_rule,
                "deduct_base_fixed_days": tier_4_base_fixed
            }
        }
    }


def save_payroll_config(cfg: dict) -> bool:
    """حفظ وتحديث إعدادات وقواعد الرواتب وحوافز الانتظام ومضاعفات الإضافي."""
    try:
        if "cycle_type" in cfg:
            set_setting("payroll_cycle_type", cfg["cycle_type"])
        if "start_day" in cfg:
            set_setting("payroll_cycle_start_day", str(cfg["start_day"]))
        if "end_day" in cfg:
            set_setting("payroll_cycle_end_day", str(cfg["end_day"]))
        if "default_incentive" in cfg:
            set_setting("default_attendance_incentive", str(cfg["default_incentive"]))
        if "daily_divisor" in cfg:
            set_setting("daily_wage_divisor", str(cfg["daily_divisor"]))
        if "overtime_normal_multiplier" in cfg:
            set_setting("overtime_normal_multiplier", str(cfg["overtime_normal_multiplier"]))
        if "overtime_friday_multiplier" in cfg:
            set_setting("overtime_friday_multiplier", str(cfg["overtime_friday_multiplier"]))

        rules = cfg.get("absence_rules", {})
        if "tier_0" in rules:
            set_setting("absence_tier_0_incentive_ded", str(rules["tier_0"].get("deduct_incentive", 0.0)))
            set_setting("absence_tier_0_base_ded_days", str(rules["tier_0"].get("deduct_base_days", 0.0)))
        if "tier_1" in rules:
            set_setting("absence_tier_1_incentive_ded", str(rules["tier_1"].get("deduct_incentive", 100.0)))
            set_setting("absence_tier_1_base_ded_days", str(rules["tier_1"].get("deduct_base_days", 0.0)))
        if "tier_2" in rules:
            set_setting("absence_tier_2_incentive_ded", str(rules["tier_2"].get("deduct_incentive", 200.0)))
            set_setting("absence_tier_2_base_ded_days", str(rules["tier_2"].get("deduct_base_days", 1.0)))
        if "tier_3" in rules:
            set_setting("absence_tier_3_incentive_ded", str(rules["tier_3"].get("deduct_incentive", 300.0)))
            set_setting("absence_tier_3_base_ded_days", str(rules["tier_3"].get("deduct_base_days", 2.0)))
        if "tier_4" in rules:
            set_setting("absence_tier_4_incentive_ded", str(rules["tier_4"].get("deduct_incentive", 500.0)))
            set_setting("absence_tier_4_base_rule", str(rules["tier_4"].get("deduct_base_rule", "days_minus_1")))
            set_setting("absence_tier_4_base_ded_days", str(rules["tier_4"].get("deduct_base_fixed_days", 0.0)))

        return True
    except Exception:
        return False


def get_payroll_cycle_dates(year: int, month: int) -> tuple:
    """
    حساب تاريخ بداية ونهاية الشهر المالي بناءً على إعدادات دورة الرواتب.
    يرجع: (start_date_str, end_date_str, description_str)
    """
    import calendar
    from datetime import date, timedelta

    cfg = get_payroll_config()
    cycle_type = cfg.get("cycle_type", "calendar")
    start_day = cfg.get("start_day", 1)
    end_day = cfg.get("end_day", 30)

    if cycle_type == "calendar" or start_day == 1:
        days_in_month = calendar.monthrange(year, month)[1]
        start_date_str = f"{year:04d}-{month:02d}-01"
        end_date_str = f"{year:04d}-{month:02d}-{days_in_month:02d}"
        desc = f"من 01-{month:02d}-{year} إلى {days_in_month:02d}-{month:02d}-{year}"
        return start_date_str, end_date_str, desc
    else:
        # Custom cycle: Starts on start_day of previous month, ends on end_day of current month
        if month == 1:
            prev_year = year - 1
            prev_month = 12
        else:
            prev_year = year
            prev_month = month - 1

        prev_max_days = calendar.monthrange(prev_year, prev_month)[1]
        actual_start_day = min(start_day, prev_max_days)

        curr_max_days = calendar.monthrange(year, month)[1]
        actual_end_day = min(end_day, curr_max_days)

        start_date_str = f"{prev_year:04d}-{prev_month:02d}-{actual_start_day:02d}"
        end_date_str = f"{year:04d}-{month:02d}-{actual_end_day:02d}"
        desc = f"من {actual_start_day:02d}-{prev_month:02d}-{prev_year} إلى {actual_end_day:02d}-{month:02d}-{year}"
        return start_date_str, end_date_str, desc


def apply_bulk_salary_increment(target_type: str = "all", dept_id: int = None,
                                inc_type: str = "percentage", value: float = 0.0,
                                apply_to: str = "base_salary") -> dict:
    """
    تطبيق زيادة جماعية للرواتب على مستوى المصنع ككل أو قسم محدد.
    - target_type: 'all' (الجميع) أو 'department' (قسم محدد)
    - inc_type: 'percentage' (نسبة مئوية %) أو 'amount' (مبلغ ثابت ج.م)
    - value: قيمة الزيادة (مثلاً 10.0 لـ 10%، أو 500.0 لمبلغ 500 ج.م)
    - apply_to: 'base_salary' (الراتب الأساسي)، 'attendance_incentive' (حافز الانتظام)، أو 'both' (كلاهما)
    """
    if value <= 0:
        return {"success": False, "error": "قيمة الزيادة يجب أن تكون أكبر من صفر."}

    conn = get_db_connection()
    c = conn.cursor()

    where_clause = "WHERE is_active=1"
    params = []
    if target_type == "department" and dept_id and dept_id > 0:
        where_clause += " AND department_id=?"
        params.append(dept_id)

    c.execute(f"SELECT id, name, base_salary, attendance_incentive FROM employees {where_clause}", params)
    employees = c.fetchall()

    if not employees:
        conn.close()
        return {"success": False, "error": "لا يوجد عمال مطابقين لنطاق الزيادة المحدد."}

    cfg = get_payroll_config()
    daily_divisor = cfg.get("daily_divisor", 30.0)

    old_total_base = sum(e[2] or 0.0 for e in employees)
    old_total_inc = sum(e[3] or 0.0 for e in employees)
    affected_count = len(employees)

    for emp_id, name, old_base, old_inc in employees:
        new_base = old_base or 0.0
        new_inc = old_inc or 0.0

        if apply_to in ("base_salary", "both"):
            if inc_type == "percentage":
                new_base = round(new_base * (1.0 + (value / 100.0)), 2)
            else:
                new_base = round(new_base + value, 2)

        if apply_to in ("attendance_incentive", "both"):
            if inc_type == "percentage":
                new_inc = round(new_inc * (1.0 + (value / 100.0)), 2)
            else:
                new_inc = round(new_inc + value, 2)

        new_daily_wage = round(new_base / daily_divisor, 2)

        c.execute("""
            UPDATE employees 
            SET base_salary=?, attendance_incentive=?, daily_wage=? 
            WHERE id=?
        """, (new_base, new_inc, new_daily_wage, emp_id))

    conn.commit()

    c.execute(f"SELECT SUM(base_salary), SUM(attendance_incentive) FROM employees {where_clause}", params)
    new_totals = c.fetchone()
    new_total_base = new_totals[0] or 0.0
    new_total_inc = new_totals[1] or 0.0

    conn.close()

    return {
        "success": True,
        "affected_count": affected_count,
        "old_total_base": round(old_total_base, 2),
        "new_total_base": round(new_total_base, 2),
        "base_increase": round(new_total_base - old_total_base, 2),
        "old_total_inc": round(old_total_inc, 2),
        "new_total_inc": round(new_total_inc, 2),
        "inc_increase": round(new_total_inc - old_total_inc, 2),
    }


def get_department_employees(dept_id: int):
    """استرجاع جميع العمال التابعين لقسم محدد."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("""
        SELECT id, name, code, base_salary, attendance_incentive, notes 
        FROM employees 
        WHERE department_id=? AND is_active=1
        ORDER BY name ASC
    """, (dept_id,))
    rows = c.fetchall()
    conn.close()
    results = []
    for r in rows:
        results.append({
            "id": r[0],
            "name": r[1],
            "code": r[2] or "",
            "base_salary": r[3] or 0.0,
            "attendance_incentive": r[4] or 500.0,
            "notes": r[5] or ""
        })
    return results


def transfer_department_employees(from_dept_id: int, to_dept_id: int):
    """نقل جميع عمال قسم إلى قسم آخر."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("UPDATE employees SET department_id=? WHERE department_id=?", (to_dept_id, from_dept_id))
    conn.commit()
    conn.close()
    return True


# ── استيراد وتحديث بيانات الإكسل ────────────────────────────────────────

def import_excel_payroll_data(file_path: str = None) -> dict:
    """
    استيراد وتحديث الأقسام، العمال، الرواتب، وحوافز الانتظام من شيت الإكسل.
    """
    import openpyxl
    if file_path is None:
        file_path = _os.path.join(_BASE_DIR, "المحركات- اغسطس  2026) ) - - Copy.xlsm")
    
    if not _os.path.exists(file_path):
        return {"success": False, "error": f"ملف الإكسل غير موجود: {file_path}"}

    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        if "الاسماء" not in wb.sheetnames:
            return {"success": False, "error": "شيت 'الاسماء' غير موجود داخل ملف الإكسل."}
        
        ws = wb["الاسماء"]
        conn = get_db_connection()
        c = conn.cursor()

        # 1. إدراج الأقسام من الشيت
        departments_in_sheet = set()
        employees_to_import = []

        for r in range(3, ws.max_row + 1):
            code_val = ws.cell(r, 2).value
            dept_val = ws.cell(r, 3).value
            name_val = ws.cell(r, 4).value
            salary_val = ws.cell(r, 5).value
            inc_val = ws.cell(r, 14).value

            if not name_val or not str(name_val).strip():
                continue
            name_str = str(name_val).strip()
            if name_str.startswith("#") or "اجمالى" in name_str or "العدد" in name_str:
                continue
            
            dept_str = str(dept_val).strip() if dept_val and not str(dept_val).startswith("#") else "عام"
            code_str = str(int(code_val)) if isinstance(code_val, (int, float)) else str(code_val or "").strip()
            if code_str.startswith("#"):
                code_str = ""

            try:
                base_sal = float(salary_val) if salary_val and not str(salary_val).startswith("#") else 0.0
            except Exception:
                base_sal = 0.0

            try:
                att_inc = float(inc_val) if inc_val and not str(inc_val).startswith("#") else (500.0 if base_sal > 0 else 0.0)
            except Exception:
                att_inc = 500.0 if base_sal > 0 else 0.0

            departments_in_sheet.add(dept_str)
            employees_to_import.append({
                "code": code_str,
                "name": name_str,
                "dept": dept_str,
                "base_salary": base_sal,
                "attendance_incentive": att_inc,
                "daily_wage": round(base_sal / 30.0, 2)
            })

        # Save all departments
        dept_id_map = {}
        for dept_name in departments_in_sheet:
            c.execute("SELECT id FROM departments WHERE name=?", (dept_name,))
            row = c.fetchone()
            if row:
                dept_id_map[dept_name] = row[0]
            else:
                c.execute("INSERT INTO departments (name) VALUES (?)", (dept_name,))
                dept_id_map[dept_name] = c.lastrowid

        # Insert or update employees
        imported_count = 0
        updated_count = 0
        for emp in employees_to_import:
            dept_id = dept_id_map.get(emp["dept"])
            c.execute("SELECT id, base_salary, attendance_incentive FROM employees WHERE name=?", (emp["name"],))
            existing = c.fetchone()
            if existing:
                emp_id = existing[0]
                c.execute("""
                    UPDATE employees 
                    SET code=?, department_id=?, base_salary=?, attendance_incentive=?, daily_wage=?
                    WHERE id=?
                """, (emp["code"], dept_id, emp["base_salary"], emp["attendance_incentive"], emp["daily_wage"], emp_id))
                updated_count += 1
            else:
                c.execute("""
                    INSERT INTO employees (name, code, department_id, base_salary, attendance_incentive, daily_wage)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (emp["name"], emp["code"], dept_id, emp["base_salary"], emp["attendance_incentive"], emp["daily_wage"]))
                imported_count += 1

        conn.commit()
        conn.close()

        return {
            "success": True,
            "total_employees": len(employees_to_import),
            "new_imported": imported_count,
            "updated": updated_count,
            "departments_count": len(departments_in_sheet)
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_db_setting(key: str, default: str = "") -> str:
    """قراءة قيمة من جدول settings."""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
        c.execute("SELECT value FROM settings WHERE key=?", (key,))
        row = c.fetchone()
        conn.close()
        return row[0] if row and row[0] is not None else default
    except Exception:
        return default


def set_db_setting(key: str, value: str):
    """حفظ قيمة في جدول settings."""
    try:
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
        c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, str(value)))
        conn.commit()
        conn.close()
    except Exception:
        pass


def is_app_password_enabled() -> bool:
    return get_db_setting("app_password_enabled", "0") == "1"


def get_app_password_hash() -> str:
    return get_db_setting("app_password_hash", "")


def get_app_password_hint() -> str:
    return get_db_setting("app_password_hint", "")


def set_app_password(password: str, hint: str = "") -> bool:
    from core.utils import hash_password
    set_db_setting("app_password_hash", hash_password(password))
    set_db_setting("app_password_hint", hint or "")
    set_db_setting("app_password_enabled", "1")
    return True


def disable_app_password() -> bool:
    set_db_setting("app_password_enabled", "0")
    return True


# ══════════════════════════════════════════════════════════════════════════════
# ── دوال إدارة السلف الشهرية (Monthly Advances) ───────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def get_monthly_advances(start_date: str = None, end_date: str = None,
                         employee_id: int = None, department_id: int = None):
    """جلب سجلات السلف الشهرية مع بيانات العامل والقسم."""
    conn = get_db_connection()
    c = conn.cursor()
    query = """
        SELECT a.id, a.date, a.employee_id, e.name, COALESCE(e.code, ''), 
               COALESCE(d.name, 'عام'), a.amount, a.notes, a.created_at, e.department_id
        FROM monthly_advances a
        JOIN employees e ON a.employee_id = e.id
        LEFT JOIN departments d ON e.department_id = d.id
        WHERE 1=1
    """
    params = []
    if start_date and end_date:
        query += " AND a.date BETWEEN ? AND ?"
        params.extend([start_date, end_date])
    elif start_date:
        query += " AND a.date >= ?"
        params.append(start_date)
    elif end_date:
        query += " AND a.date <= ?"
        params.append(end_date)

    if employee_id and employee_id > 0:
        query += " AND a.employee_id = ?"
        params.append(employee_id)

    if department_id and department_id > 0:
        query += " AND e.department_id = ?"
        params.append(department_id)

    query += " ORDER BY a.date DESC, a.id DESC"
    c.execute(query, tuple(params))
    rows = c.fetchall()
    conn.close()
    return rows


def add_monthly_advance(employee_id: int, amount: float, date: str, notes: str = "") -> dict:
    """إضافة سلفة شهرية جديدة لعامل."""
    if not employee_id or amount <= 0:
        return {"success": False, "error": "يرجى اختيار عامل وإدخال مبلغ صحيح للسلفة."}
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            INSERT INTO monthly_advances (employee_id, amount, date, notes)
            VALUES (?, ?, ?, ?)
        """, (employee_id, float(amount), date.strip(), (notes or "").strip()))
        adv_id = c.lastrowid
        conn.commit()
        conn.close()
        return {"success": True, "id": adv_id}
    except Exception as e:
        return {"success": False, "error": str(e)}


def update_monthly_advance(advance_id: int, employee_id: int, amount: float, date: str, notes: str = "") -> dict:
    """تعديل بيانات سلفة شهرية."""
    if not advance_id or amount <= 0:
        return {"success": False, "error": "بيانات السلفة غير صالحة."}
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            UPDATE monthly_advances
            SET employee_id = ?, amount = ?, date = ?, notes = ?
            WHERE id = ?
        """, (employee_id, float(amount), date.strip(), (notes or "").strip(), advance_id))
        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def delete_monthly_advance(advance_id: int) -> dict:
    """حذف سلفة شهرية."""
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("DELETE FROM monthly_advances WHERE id = ?", (advance_id,))
        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_monthly_advances_total_for_emp(employee_id: int, start_date: str, end_date: str) -> float:
    """حساب إجمالي السلف الشهرية لعامل محدد خلال فترة دورة الراتب."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("""
        SELECT COALESCE(SUM(amount), 0.0) 
        FROM monthly_advances 
        WHERE employee_id = ? AND date BETWEEN ? AND ?
    """, (employee_id, start_date, end_date))
    val = c.fetchone()[0] or 0.0
    conn.close()
    return float(val)


# ══════════════════════════════════════════════════════════════════════════════
# ── دوال إدارة السلف طويلة الأجل والأقساط (Long-Term Loans) ───────────────────
# ══════════════════════════════════════════════════════════════════════════════

def get_long_term_loans(status: str = None, employee_id: int = None, department_id: int = None):
    """
    جلب جميع السلف طويلة الأجل مع تفاصيل العمال وإحصائيات المسدد والمتبقي.
    """
    conn = get_db_connection()
    c = conn.cursor()
    query = """
        SELECT l.id, l.date, l.employee_id, e.name, COALESCE(e.code, ''),
               COALESCE(d.name, 'عام'), l.total_amount, l.installment_amount,
               l.start_date, l.remaining_amount, l.status, l.notes,
               (l.total_amount - l.remaining_amount) AS paid_amount,
               l.created_at, e.department_id
        FROM long_term_loans l
        JOIN employees e ON l.employee_id = e.id
        LEFT JOIN departments d ON e.department_id = d.id
        WHERE 1=1
    """
    params = []
    if status and status != "الكل":
        query += " AND l.status = ?"
        params.append(status)

    if employee_id and employee_id > 0:
        query += " AND l.employee_id = ?"
        params.append(employee_id)

    if department_id and department_id > 0:
        query += " AND e.department_id = ?"
        params.append(department_id)

    query += " ORDER BY l.status ASC, l.date DESC, l.id DESC"
    c.execute(query, tuple(params))
    rows = c.fetchall()
    conn.close()
    return rows


def get_long_term_loan_by_id(loan_id: int):
    """جلب تفاصيل سلفة طويلة الأجل محددة."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("""
        SELECT l.id, l.date, l.employee_id, e.name, COALESCE(e.code, ''),
               COALESCE(d.name, 'عام'), l.total_amount, l.installment_amount,
               l.start_date, l.remaining_amount, l.status, l.notes,
               (l.total_amount - l.remaining_amount) AS paid_amount
        FROM long_term_loans l
        JOIN employees e ON l.employee_id = e.id
        LEFT JOIN departments d ON e.department_id = d.id
        WHERE l.id = ?
    """, (loan_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return {
            "id": row[0],
            "date": row[1],
            "employee_id": row[2],
            "emp_name": row[3],
            "emp_code": row[4],
            "dept_name": row[5],
            "total_amount": row[6],
            "installment_amount": row[7],
            "start_date": row[8],
            "remaining_amount": row[9],
            "status": row[10],
            "notes": row[11] or "",
            "paid_amount": row[12]
        }
    return None


def add_long_term_loan(employee_id: int, total_amount: float, installment_amount: float,
                       start_date: str, date: str = None, notes: str = "") -> dict:
    """إضافة سلفة طويلة الأجل جديدة لعامل."""
    if not employee_id or total_amount <= 0 or installment_amount <= 0:
        return {"success": False, "error": "يرجى تحديد العامل وإجمالي مبلغ السلفة وقيمة القسط الشهري."}
    
    grant_date = date or start_date
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            INSERT INTO long_term_loans (employee_id, total_amount, installment_amount, 
                                         start_date, date, remaining_amount, status, notes)
            VALUES (?, ?, ?, ?, ?, ?, 'نشطة', ?)
        """, (employee_id, float(total_amount), float(installment_amount),
              start_date.strip(), grant_date.strip(), float(total_amount), (notes or "").strip()))
        loan_id = c.lastrowid
        conn.commit()
        conn.close()
        return {"success": True, "id": loan_id}
    except Exception as e:
        return {"success": False, "error": str(e)}


def update_long_term_loan(loan_id: int, installment_amount: float, status: str, notes: str = "") -> dict:
    """تعديل بيانات سلفة طويلة الأجل (القسط أو الحالة أو الملاحظات)."""
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            UPDATE long_term_loans
            SET installment_amount = ?, status = ?, notes = ?
            WHERE id = ?
        """, (float(installment_amount), status.strip(), (notes or "").strip(), loan_id))
        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def delete_long_term_loan(loan_id: int) -> dict:
    """حذف سلفة طويلة الأجل وجميع أقساطها."""
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("DELETE FROM loan_installments WHERE loan_id = ?", (loan_id,))
        c.execute("DELETE FROM long_term_loans WHERE id = ?", (loan_id,))
        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def pay_loan_installment(loan_id: int, amount: float, date: str,
                         payroll_month: int = 0, payroll_year: int = 0, notes: str = "") -> dict:
    """تسجيل سداد قسط أو دفعة من سلفة طويلة الأجل وتحديث الرصيد المتبقي."""
    if amount <= 0:
        return {"success": False, "error": "قيمة الدفعة يجب أن تكون أكبر من صفر."}
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT employee_id, remaining_amount, total_amount FROM long_term_loans WHERE id = ?", (loan_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return {"success": False, "error": "السلفة غير موجودة."}

        emp_id, rem, tot = row
        new_rem = max(0.0, rem - float(amount))
        new_status = "مسددة بالكامل" if new_rem <= 0 else "نشطة"

        c.execute("""
            INSERT INTO loan_installments (loan_id, employee_id, amount, date, payroll_month, payroll_year, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (loan_id, emp_id, float(amount), date.strip(), payroll_month, payroll_year, (notes or "").strip()))

        c.execute("""
            UPDATE long_term_loans
            SET remaining_amount = ?, status = ?
            WHERE id = ?
        """, (new_rem, new_status, loan_id))

        conn.commit()
        conn.close()
        return {"success": True, "remaining": new_rem, "status": new_status}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_loan_installments(loan_id: int = None, employee_id: int = None):
    """جلب سجل الأقساط المسددة لسلفة معينة أو لعامل معين."""
    conn = get_db_connection()
    c = conn.cursor()
    query = """
        SELECT i.id, i.loan_id, i.employee_id, e.name, i.amount, i.date,
               i.payroll_month, i.payroll_year, i.notes, i.created_at
        FROM loan_installments i
        JOIN employees e ON i.employee_id = e.id
        WHERE 1=1
    """
    params = []
    if loan_id and loan_id > 0:
        query += " AND i.loan_id = ?"
        params.append(loan_id)
    if employee_id and employee_id > 0:
        query += " AND i.employee_id = ?"
        params.append(employee_id)

    query += " ORDER BY i.date DESC, i.id DESC"
    c.execute(query, tuple(params))
    rows = c.fetchall()
    conn.close()
    return rows


def delete_loan_installment(installment_id: int) -> dict:
    """حذف قسط واستعادة قيمته إلى المتبقي من السلفة."""
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT loan_id, amount FROM loan_installments WHERE id = ?", (installment_id,))
        row = c.fetchone()
        if not row:
            conn.close()
            return {"success": False, "error": "سجل القسط غير موجود."}
        loan_id, amt = row
        c.execute("DELETE FROM loan_installments WHERE id = ?", (installment_id,))

        # Update loan remaining amount
        c.execute("SELECT total_amount FROM long_term_loans WHERE id = ?", (loan_id,))
        tot_row = c.fetchone()
        if tot_row:
            tot = tot_row[0]
            c.execute("SELECT COALESCE(SUM(amount), 0.0) FROM loan_installments WHERE loan_id = ?", (loan_id,))
            paid = c.fetchone()[0] or 0.0
            new_rem = max(0.0, tot - paid)
            new_status = "مسددة بالكامل" if new_rem <= 0 else "نشطة"
            c.execute("UPDATE long_term_loans SET remaining_amount = ?, status = ? WHERE id = ?", (new_rem, new_status, loan_id))

        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_active_long_term_loan_installment_for_emp(employee_id: int, cycle_end_date: str) -> float:
    """
    حساب إجمالي الأقساط المستحقة للسلف النشطة لعامل في هذا الشهر.
    يتم الخصم إذا كان تاريخ بدء السلفة <= تاريخ نهاية دورة الراتب وباقي عليها مبلغ.
    """
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("""
        SELECT id, installment_amount, remaining_amount, start_date
        FROM long_term_loans
        WHERE employee_id = ? AND status = 'نشطة' AND remaining_amount > 0
    """, (employee_id,))
    rows = c.fetchall()
    conn.close()

    total_installment = 0.0
    for lid, inst_amt, rem_amt, s_date in rows:
        # Check if loan start date is reached on or before cycle end date
        if s_date <= cycle_end_date or s_date[:7] <= cycle_end_date[:7]:
            due = min(inst_amt, rem_amt)
            total_installment += due

    return float(total_installment)


# ══════════════════════════════════════════════════════════════════════════════
# ── دوال إدارة أيام الاحتساب (Credited / Paid Absence Days) ───────────────────
# ══════════════════════════════════════════════════════════════════════════════

def get_credited_leaves(start_date: str = None, end_date: str = None,
                        employee_id: int = None, department_id: int = None):
    """جلب سجلات أيام الاحتساب (الإجازات المدفوعة على المصنع) مع بيانات العامل والقسم."""
    conn = get_db_connection()
    c = conn.cursor()
    query = """
        SELECT cl.id, cl.date, cl.employee_id, e.name, COALESCE(e.code, ''),
               COALESCE(d.name, 'عام'), cl.days, cl.reason, cl.notes, cl.created_at, e.department_id
        FROM credited_leaves cl
        JOIN employees e ON cl.employee_id = e.id
        LEFT JOIN departments d ON e.department_id = d.id
        WHERE 1=1
    """
    params = []
    if start_date and end_date:
        query += " AND cl.date BETWEEN ? AND ?"
        params.extend([start_date, end_date])
    elif start_date:
        query += " AND cl.date >= ?"
        params.append(start_date)
    elif end_date:
        query += " AND cl.date <= ?"
        params.append(end_date)

    if employee_id and employee_id > 0:
        query += " AND cl.employee_id = ?"
        params.append(employee_id)

    if department_id and department_id > 0:
        query += " AND e.department_id = ?"
        params.append(department_id)

    query += " ORDER BY cl.date DESC, cl.id DESC"
    c.execute(query, tuple(params))
    rows = c.fetchall()
    conn.close()
    return rows


def add_credited_leave(employee_id: int, days: float, date: str, reason: str = "", notes: str = "") -> dict:
    """إضافة أيام احتساب مدفوعة على حساب المصنع لعامل."""
    if not employee_id or days <= 0:
        return {"success": False, "error": "يرجى اختيار عامل وإدخال عدد أيام احتساب صحيح أكبر من الصفر."}
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            INSERT INTO credited_leaves (employee_id, days, date, reason, notes)
            VALUES (?, ?, ?, ?, ?)
        """, (employee_id, float(days), date.strip(), (reason or "").strip(), (notes or "").strip()))
        leave_id = c.lastrowid
        conn.commit()
        conn.close()
        return {"success": True, "id": leave_id}
    except Exception as e:
        return {"success": False, "error": str(e)}


def update_credited_leave(leave_id: int, employee_id: int, days: float, date: str, reason: str = "", notes: str = "") -> dict:
    """تعديل بيانات سجل أيام الاحتساب."""
    if not leave_id or days <= 0:
        return {"success": False, "error": "بيانات سجل الاحتساب غير صالحة."}
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            UPDATE credited_leaves
            SET employee_id = ?, days = ?, date = ?, reason = ?, notes = ?
            WHERE id = ?
        """, (employee_id, float(days), date.strip(), (reason or "").strip(), (notes or "").strip(), leave_id))
        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def delete_credited_leave(leave_id: int) -> dict:
    """حذف سجل أيام احتساب."""
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("DELETE FROM credited_leaves WHERE id = ?", (leave_id,))
        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def get_employee_credited_days(employee_id: int, start_date: str, end_date: str) -> float:
    """حساب إجمالي أيام الاحتساب (المدفوعة على المصنع) لعامل محدد خلال فترة دورة الراتب."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("""
        SELECT COALESCE(SUM(days), 0.0) 
        FROM credited_leaves 
        WHERE employee_id = ? AND date BETWEEN ? AND ?
    """, (employee_id, start_date, end_date))
    val = c.fetchone()[0] or 0.0
    conn.close()
    return float(val)







