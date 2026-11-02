from PySide6.QtCore import *
from PySide6.QtWidgets import *
from PySide6.QtGui import *
import sys
import os as _os
import sqlite3
from datetime import datetime
from core.config import *
from core.config import _BASE_DIR
from core.database import init_database
from gui.widgets import DashboardCard, BaseSectionWidget
from gui.sections.table_helpers import TableHelpersMixin
from gui.sections.dashboard import DashboardMixin
from gui.sections.hr import HRMixin
from gui.sections.loans import LoansMixin


class MainWindow(QMainWindow,
                 TableHelpersMixin,
                 DashboardMixin,
                 HRMixin,
                 LoansMixin):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("نظام إدارة شؤون العاملين - مصنع حوا")
        icon_path = get_asset_path("app_icon.ico")
        if not _os.path.exists(icon_path):
            icon_path = get_asset_path("app_icon.png")
        if _os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.setLayoutDirection(Qt.RightToLeft)
        self.toast_label = None

        self.setStyleSheet("""
            QMainWindow { background: #f0f4f8; font-family: 'Segoe UI', 'Tahoma', sans-serif; }
            QWidget { font-family: 'Segoe UI', 'Tahoma', sans-serif; }
            QLabel { color: #1e293b; }

            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox, QDateEdit, QTimeEdit {
                border: 1.5px solid #cbd5e1; border-radius: 6px; padding: 6px 10px;
                background: #ffffff; color: #0f172a; font-size: 12px;
                selection-background-color: #eff6ff; selection-color: #1d4ed8;
            }
            QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus,
            QComboBox:focus, QDateEdit:focus, QTimeEdit:focus {
                border: 2px solid #2563eb; background: #ffffff;
            }
            QLineEdit:disabled, QSpinBox:disabled, QComboBox:disabled, QDateEdit:disabled {
                background: #f1f5f9; color: #94a3b8; border-color: #e2e8f0;
            }
            QComboBox::drop-down { border: none; width: 24px; background: transparent; }
            QComboBox::down-arrow {
                border-left: 4px solid transparent; border-right: 4px solid transparent;
                border-top: 5px solid #64748b; width: 0; height: 0; margin-left: 4px;
            }
            QComboBox QAbstractItemView {
                border: 1.5px solid #cbd5e1; background: #ffffff; border-radius: 6px;
                padding: 4px; selection-background-color: #eff6ff;
                selection-color: #1d4ed8; color: #1e293b; outline: none;
            }
            QScrollBar:vertical {
                background: #f1f5f9; width: 8px; margin: 0px; border-radius: 4px;
            }
            QScrollBar::handle:vertical { background: #cbd5e1; min-height: 25px; border-radius: 4px; }
            QScrollBar::handle:vertical:hover { background: #94a3b8; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
            QScrollBar:horizontal { background: #f1f5f9; height: 8px; margin: 0px; border-radius: 4px; }
            QScrollBar::handle:horizontal { background: #cbd5e1; min-width: 25px; border-radius: 4px; }
            QScrollBar::handle:horizontal:hover { background: #94a3b8; }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; }
            QDialog { background-color: #ffffff; }
            QDialog QLabel { color: #0f172a; font-weight: 500; }
            QDialogButtonBox QPushButton {
                background-color: #2563eb; color: #ffffff; border-radius: 6px;
                padding: 6px 16px; font-weight: bold; border: none;
            }
            QDialogButtonBox QPushButton:hover { background-color: #1d4ed8; }
            QToolTip {
                background-color: #0f172a; color: #ffffff; border: 1px solid #334155;
                border-radius: 6px; padding: 6px 10px; font-size: 11px;
            }
            QTabWidget::pane { border: 1px solid #cbd5e1; border-radius: 8px; background: #ffffff; top: -1px; }
            QTabBar::tab {
                background: #f8fafc; color: #475569; padding: 9px 22px;
                border-top-left-radius: 6px; border-top-right-radius: 6px;
                font-weight: bold; font-size: 13px; margin-right: 4px;
                border: 1px solid #e2e8f0; border-bottom: none;
            }
            QTabBar::tab:selected { background: #2563eb; color: #ffffff; border-color: #2563eb; }
            QTabBar::tab:hover:!selected { background: #e2e8f0; color: #0f172a; }
        """)

        init_database()
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        # Index 0: Dashboard
        self.init_dashboard()
        # Index 1: HR Section
        self.init_all_sections()

        self.stacked_widget.currentChanged.connect(self.on_page_changed)
        self._previous_page_index = 0

        # Auto refresh every 5 seconds
        self.auto_timer = QTimer(self)
        self.auto_timer.timeout.connect(self.auto_refresh_ui)
        self.auto_timer.start(5000)

        # ── فحص التحديثات تلقائياً عند فتح التطبيق ──
        QTimer.singleShot(3000, self._check_for_app_updates)

    def _check_for_app_updates(self):
        """فحص التحديثات من GitHub بعد 3 ثواني من فتح التطبيق."""
        try:
            from gui.update_dialog import start_update_check
            start_update_check(self)
        except Exception as e:
            print(f"Update check skipped: {e}")

    def show_toast(self, message, is_error=False):
        if not hasattr(self, "toast_label") or self.toast_label is None:
            self.toast_label = QLabel(self)
            self.toast_label.setAlignment(Qt.AlignCenter)
            self.toast_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
            self.toast_timer = QTimer(self)
            self.toast_timer.setSingleShot(True)
            self.toast_timer.timeout.connect(lambda: self.toast_label.hide())
        bg_color = "#dc2626" if is_error else "#16a34a"
        self.toast_label.setText(message)
        self.toast_label.setStyleSheet(f"""
            QLabel {{
                background-color: {bg_color};
                color: #ffffff; padding: 10px 24px;
                border-radius: 8px; font-size: 13px;
                font-weight: bold; border: 2px solid #ffffff;
            }}
        """)
        self.toast_label.adjustSize()
        x = max(20, (self.width() - self.toast_label.width()) // 2)
        self.toast_label.move(x, 15)
        self.toast_label.raise_()
        self.toast_label.show()
        self.toast_timer.start(3500)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "toast_label") and self.toast_label and self.toast_label.isVisible():
            x = max(20, (self.width() - self.toast_label.width()) // 2)
            self.toast_label.move(x, 15)

    def lock_application(self):
        from gui.widgets import AppLockDialog
        from core.database import is_app_password_enabled
        if not is_app_password_enabled():
            self.show_toast("🔓 حماية كلمة المرور غير مفعّلة.", is_error=True)
            return
        login_dlg = AppLockDialog(parent=self, is_startup=False)
        if login_dlg.exec() == AppLockDialog.Accepted:
            if getattr(login_dlg, "is_master_recovery", False):
                self.notify_master_recovery_used()
            self.show_toast("✓ تم فك القفل بنجاح!")

    def notify_master_recovery_used(self):
        QMessageBox.information(
            self, "🔑 استرداد كلمة المرور",
            "تم الدخول بنجاح باستخدام كلمة مرور الاسترداد الاحتياطية."
        )

    def on_page_changed(self, index):
        prev_idx = getattr(self, "_previous_page_index", 0)
        self._previous_page_index = index
        self.refresh_dashboard_data()
        self.setFocus()

    def auto_refresh_ui(self):
        if QApplication.activeModalWidget() is not None:
            return
        self.refresh_dashboard_data()

    def init_all_sections(self):
        back = lambda: self.stacked_widget.setCurrentIndex(0)

        # ── Index 1: شؤون العاملين ──────────────────────────────
        self.sec_hr = BaseSectionWidget(
            "شؤون العاملين", back, icon_str="👥",
            primary_color="#2563eb", tint_color="#eff6ff", border_color="#bfdbfe",
            category="الحضور، الانصراف وشؤون الأفراد"
        )
        self.setup_hr_section(self.sec_hr)

        # ── دمج قسمي السلف الشهرية والسلف طويلة الأجل في تبويبات شؤون العاملين ──
        if hasattr(self, "hr_tabs") and self.hr_tabs is not None:
            self.tab_monthly_advances = self.setup_monthly_advances_tab()
            self.tab_long_term_loans = self.setup_long_term_loans_tab()

            payroll_tab_idx = -1
            for i in range(self.hr_tabs.count()):
                if "مسير" in self.hr_tabs.tabText(i) or "الرواتب" in self.hr_tabs.tabText(i):
                    payroll_tab_idx = i
                    break

            if payroll_tab_idx >= 0:
                self.hr_tabs.insertTab(payroll_tab_idx, self.tab_monthly_advances, "💰 السلف الشهرية")
                self.hr_tabs.insertTab(payroll_tab_idx + 1, self.tab_long_term_loans, "🏦 السلف طويلة الأجل")
            else:
                self.hr_tabs.addTab(self.tab_monthly_advances, "💰 السلف الشهرية")
                self.hr_tabs.addTab(self.tab_long_term_loans, "🏦 السلف طويلة الأجل")

            self.hr_tabs.currentChanged.connect(self._on_hr_tab_changed)

            # التحميل المبدئي لبيانات السلف ومواعيد الأقسام
            QTimer.singleShot(100, lambda: (
                self.populate_adv_emp_combo(),
                self.populate_adv_dept_filter(),
                self.load_monthly_advances_table(),
                self.populate_loan_emp_combo(),
                self.populate_loan_dept_filter(),
                self.load_long_term_loans_table(),
                self.load_dept_shifts_batch_table()
            ))

        self.stacked_widget.addWidget(self.sec_hr)  # index 1

        self.refresh_map = {
            1: lambda: self._on_hr_section_activated(),
        }

    def _on_hr_tab_changed(self, index: int):
        if not hasattr(self, "hr_tabs") or self.hr_tabs is None:
            return
        curr_widget = self.hr_tabs.widget(index)
        tab_text = self.hr_tabs.tabText(index)
        if hasattr(self, "tab_monthly_advances") and curr_widget == self.tab_monthly_advances:
            self.populate_adv_emp_combo()
            self.populate_adv_dept_filter()
            self.load_monthly_advances_table()
        elif hasattr(self, "tab_long_term_loans") and curr_widget == self.tab_long_term_loans:
            self.populate_loan_emp_combo()
            self.populate_loan_dept_filter()
            self.load_long_term_loans_table()
        elif "الإعدادات" in tab_text or "قواعد" in tab_text:
            self.load_dept_shifts_batch_table()

    def _on_hr_section_activated(self):
        if hasattr(self, "populate_adv_emp_combo"):
            self.populate_adv_emp_combo()
        if hasattr(self, "populate_adv_dept_filter"):
            self.populate_adv_dept_filter()
        if hasattr(self, "populate_loan_emp_combo"):
            self.populate_loan_emp_combo()
        if hasattr(self, "populate_loan_dept_filter"):
            self.populate_loan_dept_filter()
        if hasattr(self, "load_dept_shifts_batch_table"):
            self.load_dept_shifts_batch_table()

    def style_attendance_combo(self, combo, status):
        """تنسيق وتجميل قائمة حالة الحضور والغياب في جدول الحضور والانصراف."""
        if not combo:
            return

        combo.setFont(QFont("Segoe UI", 11, QFont.Bold))
        status_clean = str(status).strip()

        if "حضور" in status_clean:
            combo.setStyleSheet("""
                QComboBox {
                    background-color: #dcfce7;
                    color: #15803d;
                    border: 1.5px solid #86efac;
                    border-radius: 6px;
                    padding-top: 4px;
                    padding-bottom: 4px;
                    padding-right: 14px;
                    padding-left: 24px;
                    font-weight: bold;
                    font-size: 13px;
                    min-width: 115px;
                    min-height: 28px;
                }
                QComboBox:hover {
                    background-color: #bbf7d0;
                    border-color: #4ade80;
                }
                QComboBox::drop-down {
                    subcontrol-origin: padding;
                    subcontrol-position: left center;
                    width: 22px;
                    border: none;
                    background: transparent;
                }
                QComboBox::down-arrow {
                    image: none;
                    border-left: 4px solid transparent;
                    border-right: 4px solid transparent;
                    border-top: 5px solid #15803d;
                    width: 0px;
                    height: 0px;
                    margin-right: 4px;
                }
                QComboBox QAbstractItemView {
                    background-color: #ffffff;
                    color: #0f172a;
                    border: 1.5px solid #cbd5e1;
                    border-radius: 6px;
                    padding: 4px;
                    selection-background-color: #dcfce7;
                    selection-color: #15803d;
                    font-size: 13px;
                    font-weight: bold;
                    outline: none;
                }
                QComboBox QAbstractItemView::item {
                    min-height: 28px;
                    padding: 4px 10px;
                    border-radius: 4px;
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
                    padding-top: 4px;
                    padding-bottom: 4px;
                    padding-right: 14px;
                    padding-left: 24px;
                    font-weight: bold;
                    font-size: 13px;
                    min-width: 115px;
                    min-height: 28px;
                }
                QComboBox:hover {
                    background-color: #fecaca;
                    border-color: #f87171;
                }
                QComboBox::drop-down {
                    subcontrol-origin: padding;
                    subcontrol-position: left center;
                    width: 22px;
                    border: none;
                    background: transparent;
                }
                QComboBox::down-arrow {
                    image: none;
                    border-left: 4px solid transparent;
                    border-right: 4px solid transparent;
                    border-top: 5px solid #b91c1c;
                    width: 0px;
                    height: 0px;
                    margin-right: 4px;
                }
                QComboBox QAbstractItemView {
                    background-color: #ffffff;
                    color: #0f172a;
                    border: 1.5px solid #cbd5e1;
                    border-radius: 6px;
                    padding: 4px;
                    selection-background-color: #fee2e2;
                    selection-color: #b91c1c;
                    font-size: 13px;
                    font-weight: bold;
                    outline: none;
                }
                QComboBox QAbstractItemView::item {
                    min-height: 28px;
                    padding: 4px 10px;
                    border-radius: 4px;
                }
            """)

    def load_dept_shifts_batch_table(self):
        """تحميل وتنسيق جدول مواعيد وساعات عمل الأقسام بتصميم أنيق ودقيق وعناوين صحيحة."""
        if not hasattr(self, "dept_shifts_table") or self.dept_shifts_table is None:
            return

        headers = ["ID", "اسم القسم", "🕒 وقت الحضور (البداية)", "🕑 وقت الانصراف (النهاية)", "⏱️ ساعات العمل اليومية", "💾 حفظ التعديل"]
        self.dept_shifts_table.setColumnCount(len(headers))
        self.dept_shifts_table.setHorizontalHeaderLabels(headers)
        self.dept_shifts_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.dept_shifts_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.dept_shifts_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.dept_shifts_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.dept_shifts_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.dept_shifts_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.dept_shifts_table.verticalHeader().setVisible(False)
        self.dept_shifts_table.setAlternatingRowColors(True)

        self.dept_shifts_table.setStyleSheet("""
            QTableWidget {
                border: 1.5px solid #cbd5e1;
                border-radius: 8px;
                background: #ffffff;
                gridline-color: #e2e8f0;
                alternate-background-color: #f8fafc;
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
            QTableWidget::item {
                padding: 6px 12px;
                color: #0f172a;
                font-size: 13px;
            }
        """)

        from core.database import get_departments, update_department
        from core.utils import parse_time_str, format_time_ar, ARABIC_LOCALE
        depts = get_departments()
        self.dept_shifts_table.setRowCount(len(depts))

        for r_idx, d in enumerate(depts):
            dept_id = d[0]
            dept_name = d[1]
            dept_code = d[2] or ""
            dept_notes = d[3] or ""
            s_time_str = d[4] if len(d) > 4 and d[4] else "08:00 ص"
            e_time_str = d[5] if len(d) > 5 and d[5] else "07:00 م"
            w_hours = float(d[6]) if len(d) > 6 and d[6] is not None else 11.0

            self.dept_shifts_table.setRowHeight(r_idx, 44)

            # Col 0: ID
            id_item = QTableWidgetItem(str(dept_id))
            id_item.setTextAlignment(Qt.AlignCenter)
            id_item.setFlags(id_item.flags() & ~Qt.ItemIsEditable)
            self.dept_shifts_table.setItem(r_idx, 0, id_item)

            # Col 1: اسم القسم
            name_item = QTableWidgetItem(dept_name)
            name_item.setFont(QFont("Segoe UI", 11, QFont.Bold))
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.dept_shifts_table.setItem(r_idx, 1, name_item)

            # Col 2: وقت الحضور (البداية)
            s_edit = QTimeEdit()
            s_edit.setDisplayFormat("hh:mm AP")
            s_edit.setLocale(ARABIC_LOCALE)
            t_s = parse_time_str(s_time_str)
            if t_s.isValid():
                s_edit.setTime(t_s)
            else:
                s_edit.setTime(QTime(8, 0))
            s_edit.setStyleSheet("""
                QTimeEdit {
                    padding: 5px 10px;
                    border: 1.5px solid #cbd5e1;
                    border-radius: 6px;
                    font-size: 13px;
                    font-weight: bold;
                    min-width: 125px;
                    min-height: 30px;
                    background: #ffffff;
                    color: #0f172a;
                }
                QTimeEdit:focus {
                    border: 2px solid #2563eb;
                }
            """)
            self.dept_shifts_table.setCellWidget(r_idx, 2, s_edit)

            # Col 3: وقت الانصراف (النهاية)
            e_edit = QTimeEdit()
            e_edit.setDisplayFormat("hh:mm AP")
            e_edit.setLocale(ARABIC_LOCALE)
            t_e = parse_time_str(e_time_str)
            if t_e.isValid():
                e_edit.setTime(t_e)
            else:
                e_edit.setTime(QTime(19, 0))
            e_edit.setStyleSheet("""
                QTimeEdit {
                    padding: 5px 10px;
                    border: 1.5px solid #cbd5e1;
                    border-radius: 6px;
                    font-size: 13px;
                    font-weight: bold;
                    min-width: 125px;
                    min-height: 30px;
                    background: #ffffff;
                    color: #0f172a;
                }
                QTimeEdit:focus {
                    border: 2px solid #2563eb;
                }
            """)
            self.dept_shifts_table.setCellWidget(r_idx, 3, e_edit)

            # Col 4: ساعات العمل اليومية
            h_item = QTableWidgetItem(f"{w_hours:.1f} ساعة")
            h_item.setTextAlignment(Qt.AlignCenter)
            h_item.setFont(QFont("Segoe UI", 11, QFont.Bold))
            h_item.setForeground(QColor("#2563eb"))
            h_item.setFlags(h_item.flags() & ~Qt.ItemIsEditable)
            self.dept_shifts_table.setItem(r_idx, 4, h_item)

            # دالة تحديث الساعات تلقائياً عند تغيير وقت الحضور أو الانصراف
            def make_update_h(row, s_widget, e_widget):
                def updater():
                    t1 = s_widget.time()
                    t2 = e_widget.time()
                    m1 = t1.hour() * 60 + t1.minute()
                    m2 = t2.hour() * 60 + t2.minute()
                    dur_m = m2 - m1
                    if dur_m < 0:
                        dur_m += 24 * 60
                    dur_h = round(dur_m / 60.0, 2)
                    item = self.dept_shifts_table.item(row, 4)
                    if item:
                        item.setText(f"{dur_h:.1f} ساعة")
                return updater

            s_edit.timeChanged.connect(make_update_h(r_idx, s_edit, e_edit))
            e_edit.timeChanged.connect(make_update_h(r_idx, s_edit, e_edit))

            # Col 5: زر الحفظ
            save_btn = QPushButton("💾 حفظ")
            save_btn.setCursor(Qt.PointingHandCursor)
            save_btn.setStyleSheet("""
                QPushButton {
                    background-color: #16a34a;
                    color: #ffffff;
                    padding: 6px 18px;
                    font-weight: bold;
                    border-radius: 6px;
                    font-size: 12px;
                    border: none;
                    min-height: 28px;
                }
                QPushButton:hover {
                    background-color: #15803d;
                }
                QPushButton:pressed {
                    background-color: #166534;
                }
            """)

            def make_save_dept(did, dname, dcode, dnotes, s_widget, e_widget):
                def saver():
                    t1 = s_widget.time()
                    t2 = e_widget.time()
                    s_str = format_time_ar(t1)
                    e_str = format_time_ar(t2)
                    m1 = t1.hour() * 60 + t1.minute()
                    m2 = t2.hour() * 60 + t2.minute()
                    dur_m = m2 - m1
                    if dur_m < 0:
                        dur_m += 24 * 60
                    dur_h = round(dur_m / 60.0, 2)

                    res = update_department(did, dname, dcode, dnotes, s_str, e_str, dur_h)
                    if res.get("success"):
                        self.show_toast(f"✓ تم حفظ مواعيد قسم '{dname}' ({s_str} إلى {e_str} - {dur_h:.1f} س) بنجاح")
                    else:
                        QMessageBox.critical(self, "خطأ", f"تعذر حفظ مواعيد القسم: {res.get('error')}")
                return saver

            save_btn.clicked.connect(make_save_dept(dept_id, dept_name, dept_code, dept_notes, s_edit, e_edit))
            self.dept_shifts_table.setCellWidget(r_idx, 5, save_btn)

    def show_employee_payslip_dialog(self, employee_id: int, year: int, month: int):
        """عرض مفردات مرتب تفصيلية للعامل تشمل السلف الشهرية وأقساط السلف طويلة الأجل."""
        from core.payroll import calculate_employee_payroll
        rec = calculate_employee_payroll(employee_id, year, month)
        if not rec:
            QMessageBox.warning(self, "تنبيه", "لم يتم العثور على بيانات راتب هذا العامل.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f"مفردات مرتب: {rec['name']} - شهر {month}/{year}")
        dlg.setLayoutDirection(Qt.RightToLeft)
        dlg.resize(520, 620)
        dlg.setStyleSheet("background: white; font-family: 'Segoe UI';")

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # ترويسة
        header = QFrame()
        header.setStyleSheet("background: #0f172a; border-radius: 8px; padding: 12px;")
        h_layout = QVBoxLayout(header)
        h_layout.setSpacing(4)
        lbl_h1 = QLabel(f"📄 كشف مفردات مرتب للعامل: {rec['name']}")
        lbl_h1.setFont(QFont("Segoe UI", 12, QFont.Bold))
        lbl_h1.setStyleSheet("color: white;")
        code_str = f"الكود: {rec['code']}" if rec['code'] else ""
        lbl_h2 = QLabel(f"{code_str} | القسم: {rec['department_name']} | شهر: {month} / {year}")
        lbl_h2.setFont(QFont("Segoe UI", 10))
        lbl_h2.setStyleSheet("color: #94a3b8;")
        h_layout.addWidget(lbl_h1)
        h_layout.addWidget(lbl_h2)
        layout.addWidget(header)

        # جدول البنود
        body = QFrame()
        body.setStyleSheet("background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px;")
        b_layout = QFormLayout(body)
        b_layout.setSpacing(8)

        def add_row(lbl_text, val_text, color="#0f172a", is_bold=False):
            lbl = QLabel(lbl_text)
            lbl.setFont(QFont("Segoe UI", 10, QFont.Bold if is_bold else QFont.Normal))
            val = QLabel(val_text)
            val.setFont(QFont("Segoe UI", 10, QFont.Bold if is_bold else QFont.Normal))
            val.setStyleSheet(f"color: {color};")
            b_layout.addRow(lbl, val)

        add_row("الراتب الأساسي الشهري:", f"{rec['base_salary']:,.2f} ج.م", "#0f172a", True)
        add_row("أجر اليوم التقريبي (الأساسي ÷ 30):", f"{rec['daily_wage']:,.2f} ج.م", "#0284c7")
        add_row("أيام الحضور الفعلية:", f"{rec['present_days']} يوم", "#16a34a")

        raw_abs = rec.get('raw_absent_days', rec['absent_days'])
        cred_days = rec.get('credited_days', 0.0)
        eff_abs = rec['absent_days']

        add_row("إجمالي أيام الغياب المسجلة:", f"{raw_abs} يوم", "#dc2626" if raw_abs > 0 else "#475569")
        if cred_days > 0:
            add_row("أيام الاحتساب (مدفوعة على المصنع):", f"🎁 {cred_days:g} يوم (لا تُخصم من العامل)", "#7c3aed", True)
            add_row("أيام الغياب الخاضعة للخصم:", f"{eff_abs:g} يوم (بعد استبعاد أيام الاحتساب)", "#b45309" if eff_abs > 0 else "#15803d", True)
        else:
            add_row("أيام الغياب الخاضعة للخصم:", f"{eff_abs:g} يوم", "#dc2626" if eff_abs > 0 else "#475569")

        add_row("حافز الانتظام المستحق:", f"+ {rec['earned_incentive']:,.2f} ج.م (من أصل {rec['max_incentive']:,.0f})", "#16a34a")

        if rec['absence_base_deduction'] > 0:
            add_row("خصم غياب من الراتب الأساسي:", f"- {rec['absence_base_deduction']:,.2f} ج.م ({rec['absent_days_from_base']} يوم)", "#dc2626")
        if rec['incentive_deduction'] > 0:
            add_row("خصم غياب من حافز الانتظام:", f"- {rec['incentive_deduction']:,.2f} ج.م", "#dc2626")

        if rec['overtime_hours'] > 0:
            add_row("ساعات السهرات والإضافي:", f"{rec['overtime_hours']:.1f} ساعة", "#7c3aed")
            add_row("قيمة أجر السهرات المستحقة:", f"+ {rec['overtime_pay']:,.2f} ج.م", "#7c3aed")

        if rec['extra_bonuses_total'] > 0:
            add_row("مكافآت وحوافز إضافية:", f"+ {rec['extra_bonuses_total']:,.2f} ج.م", "#16a34a")
        if rec['penalties_total'] > 0:
            add_row("جزاءات واستقطاعات أخرى:", f"- {rec['penalties_total']:,.2f} ج.م", "#dc2626")

        # بند السلف الشهرية
        if rec.get('monthly_advances_total', 0.0) > 0:
            add_row("💰 السلف الشهرية المخصومة:", f"- {rec['monthly_advances_total']:,.2f} ج.م", "#dc2626", True)

        # بند أقساط السلف طويلة الأجل
        if rec.get('loan_installments_total', 0.0) > 0:
            add_row("🏦 قسط سلفة طويلة الأجل:", f"- {rec['loan_installments_total']:,.2f} ج.م", "#dc2626", True)

        layout.addWidget(body)

        # الصافي
        net_frame = QFrame()
        net_frame.setStyleSheet("background: #f5f3ff; border: 1.5px solid #c4b5fd; border-radius: 8px; padding: 12px;")
        n_layout = QHBoxLayout(net_frame)
        lbl_net_t = QLabel("💵 صافي الراتب النهائي المستحق للقبض:")
        lbl_net_t.setFont(QFont("Segoe UI", 11, QFont.Bold))
        lbl_net_t.setStyleSheet("color: #6b21a8;")

        lbl_net_v = QLabel(f"{rec['net_payout']:,.2f} ج.م")
        lbl_net_v.setFont(QFont("Segoe UI", 14, QFont.Bold))
        lbl_net_v.setStyleSheet("color: #7c3aed;")

        n_layout.addWidget(lbl_net_t)
        n_layout.addStretch()
        n_layout.addWidget(lbl_net_v)
        layout.addWidget(net_frame)

        btn_close = QPushButton("إغلاق")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setStyleSheet("background: #64748b; color: white; padding: 8px 16px; font-weight: bold; border-radius: 6px;")
        btn_close.clicked.connect(dlg.accept)
        layout.addWidget(btn_close, alignment=Qt.AlignLeft)

        dlg.exec()
