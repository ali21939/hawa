# -*- coding: utf-8 -*-
from PySide6.QtCore import Qt, QDate, QTimer
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QComboBox, QDateEdit, QDoubleSpinBox, QFrame, QMessageBox,
    QDialog, QFormLayout, QDialogButtonBox, QProgressBar,
    QFileDialog, QScrollArea, QSplitter
)
from PySide6.QtGui import QFont, QColor, QIcon
from datetime import datetime, date
import sqlite3
import os

from core.database import (
    get_db_connection,
    get_monthly_advances, add_monthly_advance, update_monthly_advance, delete_monthly_advance,
    get_long_term_loans, get_long_term_loan_by_id, add_long_term_loan, update_long_term_loan,
    delete_long_term_loan, pay_loan_installment, get_loan_installments, delete_loan_installment,
    get_departments
)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import openpyxl.utils


class LoansMixin:
    """
    Mixin class implementing management tabs for:
    1. Monthly Advances (السلف الشهرية - تخصم بالكامل من راتب نفس الشهر).
    2. Long-Term Loans (السلف طويلة الأجل والأقساط الشهرية).
    """

    # ══════════════════════════════════════════════════════════════════════════
    # ── 1. تبويب السلف الشهرية (Monthly Advances) ──────────────────────────────
    # ══════════════════════════════════════════════════════════════════════════

    def setup_monthly_advances_tab(self) -> QWidget:
        tab = QWidget()
        main_layout = QVBoxLayout(tab)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        # ── 1. البطاقات الإحصائية العلوية (Metrics) ─────────────────────────
        metrics_frame = QFrame()
        metrics_frame.setStyleSheet("background: #f8fafc; border-radius: 8px; padding: 4px; border: 1px solid #e2e8f0;")
        m_layout = QHBoxLayout(metrics_frame)
        m_layout.setContentsMargins(6, 4, 6, 4)
        m_layout.setSpacing(12)

        self.adv_card_total = self._create_loan_metric_card("💰 إجمالي سلف الفترة", "0.00 ج.م", "#2563eb", "#eff6ff")
        self.adv_card_count = self._create_loan_metric_card("📋 عدد السلف المسجلة", "0", "#059669", "#f0fdf4")
        self.adv_card_emps = self._create_loan_metric_card("👥 العمال المستفيدين", "0", "#7c3aed", "#f5f3ff")

        m_layout.addWidget(self.adv_card_total)
        m_layout.addWidget(self.adv_card_count)
        m_layout.addWidget(self.adv_card_emps)
        main_layout.addWidget(metrics_frame)

        # ── 2. نموذج إضافة سلفة شهرية جديدة ──────────────────────────────────
        form_frame = QFrame()
        form_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1.5px solid #cbd5e1;
                border-radius: 8px;
                padding: 12px;
            }
        """)
        f_layout = QVBoxLayout(form_frame)
        f_layout.setContentsMargins(10, 10, 10, 10)
        f_layout.setSpacing(12)

        form_title = QLabel("➕ تسجيل سلفة شهرية جديدة (تُخصم من راتب نفس الشهر تلقائياً)")
        form_title.setFont(QFont("Segoe UI", 11, QFont.Bold))
        form_title.setStyleSheet("color: #1e3a8a; border: none; margin-bottom: 2px;")
        f_layout.addWidget(form_title)

        inputs_layout = QHBoxLayout()
        inputs_layout.setSpacing(14)

        # التاريخ
        lbl_date = QLabel("📅 التاريخ:")
        lbl_date.setStyleSheet("font-weight: bold; border: none; font-size: 13px; color: #1e293b;")
        self.adv_date_picker = QDateEdit()
        self.adv_date_picker.setCalendarPopup(True)
        self.adv_date_picker.setDate(QDate.currentDate())
        self.adv_date_picker.setDisplayFormat("yyyy-MM-dd")
        self.adv_date_picker.setMinimumWidth(125)
        self.adv_date_picker.setMinimumHeight(34)

        # العامل
        lbl_emp = QLabel("👤 العامل:")
        lbl_emp.setStyleSheet("font-weight: bold; border: none; font-size: 13px; color: #1e293b;")
        self.adv_emp_combo = QComboBox()
        self.adv_emp_combo.setMinimumWidth(240)
        self.adv_emp_combo.setMinimumHeight(34)

        # المبلغ
        lbl_amt = QLabel("💵 المبلغ (ج.م):")
        lbl_amt.setStyleSheet("font-weight: bold; border: none; font-size: 13px; color: #1e293b;")
        self.adv_amount_spin = QDoubleSpinBox()
        self.adv_amount_spin.setRange(1.0, 500000.0)
        self.adv_amount_spin.setDecimals(2)
        self.adv_amount_spin.setValue(500.0)
        self.adv_amount_spin.setMinimumWidth(140)
        self.adv_amount_spin.setMinimumHeight(34)
        self.adv_amount_spin.setSuffix(" ج.م")

        # الملاحظات
        lbl_notes = QLabel("📝 ملاحظات:")
        lbl_notes.setStyleSheet("font-weight: bold; border: none; font-size: 13px; color: #1e293b;")
        self.adv_notes_input = QLineEdit()
        self.adv_notes_input.setPlaceholderText("سبب السلفة أو توضيح اختياري...")
        self.adv_notes_input.setMinimumHeight(34)

        # زر الحفظ
        self.adv_save_btn = QPushButton("💾 تسجيل السلفة")
        self.adv_save_btn.setCursor(Qt.PointingHandCursor)
        self.adv_save_btn.setMinimumHeight(36)
        self.adv_save_btn.setMinimumWidth(140)
        self.adv_save_btn.setStyleSheet("""
            QPushButton {
                background-color: #2563eb; color: #ffffff;
                font-weight: bold; padding: 8px 18px;
                border-radius: 6px; font-size: 13px; border: none;
            }
            QPushButton:hover { background-color: #1d4ed8; }
            QPushButton:pressed { background-color: #1e40af; }
        """)
        self.adv_save_btn.clicked.connect(self.add_monthly_advance_handler)

        inputs_layout.addWidget(lbl_date)
        inputs_layout.addWidget(self.adv_date_picker)
        inputs_layout.addWidget(lbl_emp)
        inputs_layout.addWidget(self.adv_emp_combo)
        inputs_layout.addWidget(lbl_amt)
        inputs_layout.addWidget(self.adv_amount_spin)
        inputs_layout.addWidget(lbl_notes)
        inputs_layout.addWidget(self.adv_notes_input)
        inputs_layout.addWidget(self.adv_save_btn)
        f_layout.addLayout(inputs_layout)
        main_layout.addWidget(form_frame)

        # ── 3. شريط التصفية والبحث والتصدير ──────────────────────────────────
        filter_bar = QHBoxLayout()
        filter_bar.setSpacing(8)

        lbl_f_from = QLabel("من:")
        lbl_f_from.setStyleSheet("font-weight: bold;")
        self.adv_filter_from = QDateEdit()
        self.adv_filter_from.setCalendarPopup(True)
        self.adv_filter_from.setDisplayFormat("yyyy-MM-dd")
        curr_d = QDate.currentDate()
        self.adv_filter_from.setDate(QDate(curr_d.year(), curr_d.month(), 1))
        self.adv_filter_from.setFixedWidth(110)

        lbl_f_to = QLabel("إلى:")
        lbl_f_to.setStyleSheet("font-weight: bold;")
        self.adv_filter_to = QDateEdit()
        self.adv_filter_to.setCalendarPopup(True)
        self.adv_filter_to.setDisplayFormat("yyyy-MM-dd")
        self.adv_filter_to.setDate(QDate(curr_d.year(), curr_d.month(), curr_d.daysInMonth()))
        self.adv_filter_to.setFixedWidth(110)

        btn_this_month = QPushButton("📅 هذا الشهر")
        btn_this_month.setCursor(Qt.PointingHandCursor)
        btn_this_month.setStyleSheet("background: #f1f5f9; color: #334155; font-weight: bold; padding: 5px 10px; border-radius: 4px; border: 1px solid #cbd5e1;")
        btn_this_month.clicked.connect(self._adv_set_this_month)

        lbl_dept = QLabel("🏢 القسم:")
        lbl_dept.setStyleSheet("font-weight: bold;")
        self.adv_filter_dept = QComboBox()
        self.adv_filter_dept.setFixedWidth(140)

        lbl_search = QLabel("🔍 بحث:")
        lbl_search.setStyleSheet("font-weight: bold;")
        self.adv_search_input = QLineEdit()
        self.adv_search_input.setPlaceholderText("ابحث باسم العامل أو الكود...")
        self.adv_search_input.setMinimumWidth(160)
        self.adv_search_input.textChanged.connect(self.filter_monthly_advances_table)

        btn_refresh = QPushButton("🔄 تحديث")
        btn_refresh.setCursor(Qt.PointingHandCursor)
        btn_refresh.setStyleSheet("background: #f8fafc; color: #334155; font-weight: bold; padding: 6px 12px; border-radius: 6px; border: 1px solid #cbd5e1;")
        btn_refresh.clicked.connect(self.load_monthly_advances_table)

        btn_export = QPushButton("📥 تصدير Excel")
        btn_export.setCursor(Qt.PointingHandCursor)
        btn_export.setStyleSheet("background: #15803d; color: white; font-weight: bold; padding: 6px 14px; border-radius: 6px; border: none;")
        btn_export.clicked.connect(self.export_monthly_advances_excel)

        self.adv_filter_from.dateChanged.connect(self.load_monthly_advances_table)
        self.adv_filter_to.dateChanged.connect(self.load_monthly_advances_table)
        self.adv_filter_dept.currentIndexChanged.connect(self.load_monthly_advances_table)

        filter_bar.addWidget(lbl_f_from)
        filter_bar.addWidget(self.adv_filter_from)
        filter_bar.addWidget(lbl_f_to)
        filter_bar.addWidget(self.adv_filter_to)
        filter_bar.addWidget(btn_this_month)
        filter_bar.addWidget(lbl_dept)
        filter_bar.addWidget(self.adv_filter_dept)
        filter_bar.addWidget(lbl_search)
        filter_bar.addWidget(self.adv_search_input)
        filter_bar.addWidget(btn_refresh)
        filter_bar.addWidget(btn_export)
        main_layout.addLayout(filter_bar)

        # ── 4. جدول السلف الشهرية ─────────────────────────────────────────────
        self.monthly_advances_table = QTableWidget()
        headers = ["ID", "التاريخ", "الكود", "اسم العامل", "القسم", "مبلغ السلفة (ج.م)", "الملاحظات", "تعديل", "حذف"]
        self.monthly_advances_table.setColumnCount(len(headers))
        self.monthly_advances_table.setHorizontalHeaderLabels(headers)
        self.monthly_advances_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.monthly_advances_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.monthly_advances_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.monthly_advances_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.monthly_advances_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.monthly_advances_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        self.monthly_advances_table.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeToContents)
        self.monthly_advances_table.setAlternatingRowColors(True)
        self.monthly_advances_table.verticalHeader().setVisible(False)
        self.monthly_advances_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.monthly_advances_table.setStyleSheet("""
            QTableWidget {
                border: 1.5px solid #cbd5e1; border-radius: 8px; background: #ffffff;
                gridline-color: #e2e8f0; alternate-background-color: #f8fafc;
            }
            QHeaderView::section {
                background-color: #0f172a; color: #ffffff; padding: 9px 8px;
                font-weight: bold; font-size: 12px; border: 1px solid #1e293b;
            }
            QTableWidget::item { padding: 6px; font-size: 12px; color: #0f172a; }
        """)
        main_layout.addWidget(self.monthly_advances_table)

        return tab

    def _adv_set_this_month(self):
        curr_d = QDate.currentDate()
        self.adv_filter_from.setDate(QDate(curr_d.year(), curr_d.month(), 1))
        self.adv_filter_to.setDate(QDate(curr_d.year(), curr_d.month(), curr_d.daysInMonth()))

    def populate_adv_emp_combo(self):
        """تعبئة قائمة اختيار العمال في تبويب السلف الشهرية."""
        if not hasattr(self, "adv_emp_combo") or self.adv_emp_combo is None:
            return
        self.adv_emp_combo.clear()
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            SELECT e.id, e.name, COALESCE(e.code, ''), COALESCE(d.name, 'عام')
            FROM employees e
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE e.is_active = 1
            ORDER BY e.name ASC
        """)
        rows = c.fetchall()
        conn.close()
        for emp_id, name, code, dept in rows:
            label = f"{name} ({dept})" + (f" - كود: {code}" if code else "")
            self.adv_emp_combo.addItem(label, emp_id)

    def populate_adv_dept_filter(self):
        """تعبئة قائمة الأقسام في فلتر السلف."""
        if not hasattr(self, "adv_filter_dept") or self.adv_filter_dept is None:
            return
        curr_text = self.adv_filter_dept.currentText()
        self.adv_filter_dept.blockSignals(True)
        self.adv_filter_dept.clear()
        self.adv_filter_dept.addItem("🏢 جميع الأقسام", 0)
        depts = get_departments()
        for d in depts:
            self.adv_filter_dept.addItem(d[1], d[0])
        idx = self.adv_filter_dept.findText(curr_text)
        if idx >= 0:
            self.adv_filter_dept.setCurrentIndex(idx)
        self.adv_filter_dept.blockSignals(False)

    def add_monthly_advance_handler(self):
        """معالج إضافة سلفة شهرية جديدة."""
        emp_id = self.adv_emp_combo.currentData()
        if not emp_id:
            QMessageBox.warning(self, "تنبيه", "يرجى اختيار عامل أولاً.")
            return
        amt = self.adv_amount_spin.value()
        if amt <= 0:
            QMessageBox.warning(self, "تنبيه", "يرجى إدخال مبلغ صحيح للسلفة.")
            return
        date_str = self.adv_date_picker.date().toString("yyyy-MM-dd")
        notes = self.adv_notes_input.text().strip()

        res = add_monthly_advance(emp_id, amt, date_str, notes)
        if res.get("success"):
            self.show_toast(f"✓ تم تسجيل سلفة بقيمة {amt:,.2f} ج.م بنجاح")
            self.adv_notes_input.clear()
            self.adv_amount_spin.setValue(500.0)
            self.load_monthly_advances_table()
        else:
            QMessageBox.critical(self, "خطأ", f"تعذر تسجيل السلفة: {res.get('error')}")

    def load_monthly_advances_table(self):
        """تحميل وعرض بيانات السلف الشهرية في الجدول."""
        if not hasattr(self, "monthly_advances_table") or self.monthly_advances_table is None:
            return

        start_date = self.adv_filter_from.date().toString("yyyy-MM-dd")
        end_date = self.adv_filter_to.date().toString("yyyy-MM-dd")
        dept_id = self.adv_filter_dept.currentData() if hasattr(self, "adv_filter_dept") else None

        records = get_monthly_advances(start_date, end_date, department_id=dept_id)

        self.monthly_advances_table.setRowCount(len(records))
        total_amount = 0.0
        unique_emps = set()

        for row_idx, r in enumerate(records):
            adv_id, adv_date, emp_id, emp_name, emp_code, dept_name, amount, notes, created_at, d_id = r
            total_amount += amount
            unique_emps.add(emp_id)

            item_id = QTableWidgetItem(str(adv_id))
            item_id.setTextAlignment(Qt.AlignCenter)

            item_date = QTableWidgetItem(adv_date)
            item_date.setTextAlignment(Qt.AlignCenter)

            item_code = QTableWidgetItem(emp_code or "-")
            item_code.setTextAlignment(Qt.AlignCenter)

            item_name = QTableWidgetItem(emp_name)
            item_name.setFont(QFont("Segoe UI", 10, QFont.Bold))

            item_dept = QTableWidgetItem(dept_name or "عام")
            item_dept.setTextAlignment(Qt.AlignCenter)

            item_amt = QTableWidgetItem(f"{amount:,.2f} ج.م")
            item_amt.setTextAlignment(Qt.AlignCenter)
            item_amt.setFont(QFont("Segoe UI", 10, QFont.Bold))
            item_amt.setForeground(QColor("#b91c1c"))

            item_notes = QTableWidgetItem(notes or "-")

            self.monthly_advances_table.setItem(row_idx, 0, item_id)
            self.monthly_advances_table.setItem(row_idx, 1, item_date)
            self.monthly_advances_table.setItem(row_idx, 2, item_code)
            self.monthly_advances_table.setItem(row_idx, 3, item_name)
            self.monthly_advances_table.setItem(row_idx, 4, item_dept)
            self.monthly_advances_table.setItem(row_idx, 5, item_amt)
            self.monthly_advances_table.setItem(row_idx, 6, item_notes)

            # أزرار الإجراءات
            btn_edit = QPushButton("✏️ تعديل")
            btn_edit.setCursor(Qt.PointingHandCursor)
            btn_edit.setStyleSheet("background: #2563eb; color: white; border-radius: 4px; padding: 4px 8px; font-weight: bold; border: none;")
            btn_edit.clicked.connect(lambda chk=False, aid=adv_id, eid=emp_id, a=amount, d=adv_date, n=notes: self.edit_monthly_advance_dialog(aid, eid, a, d, n))

            btn_del = QPushButton("🗑️ حذف")
            btn_del.setCursor(Qt.PointingHandCursor)
            btn_del.setStyleSheet("background: #ef4444; color: white; border-radius: 4px; padding: 4px 8px; font-weight: bold; border: none;")
            btn_del.clicked.connect(lambda chk=False, aid=adv_id, en=emp_name, a=amount: self.delete_monthly_advance_handler(aid, en, a))

            self.monthly_advances_table.setCellWidget(row_idx, 7, btn_edit)
            self.monthly_advances_table.setCellWidget(row_idx, 8, btn_del)

        # تحديث البطاقات
        if hasattr(self, "adv_card_total"):
            self._update_loan_card(self.adv_card_total, f"{total_amount:,.2f} ج.م")
            self._update_loan_card(self.adv_card_count, str(len(records)))
            self._update_loan_card(self.adv_card_emps, str(len(unique_emps)))

        self.filter_monthly_advances_table()

    def filter_monthly_advances_table(self):
        """تصفية جدول السلف الشهرية بحسب نص البحث."""
        if not hasattr(self, "monthly_advances_table") or not hasattr(self, "adv_search_input"):
            return
        query = self.adv_search_input.text().strip().lower()
        for r in range(self.monthly_advances_table.rowCount()):
            match = False
            for c in [1, 2, 3, 4, 6]:
                item = self.monthly_advances_table.item(r, c)
                if item and query in item.text().lower():
                    match = True
                    break
            self.monthly_advances_table.setRowHidden(r, not match)

    def edit_monthly_advance_dialog(self, advance_id: int, current_emp_id: int, current_amount: float, current_date: str, current_notes: str):
        """نافذة تعديل سلفة شهرية."""
        dlg = QDialog(self)
        dlg.setWindowTitle("✏️ تعديل بيانات السلفة الشهرية")
        dlg.setLayoutDirection(Qt.RightToLeft)
        dlg.setMinimumWidth(380)
        layout = QVBoxLayout(dlg)

        form = QFormLayout()
        form.setSpacing(10)

        combo_emp = QComboBox()
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT id, name, COALESCE(code, '') FROM employees WHERE is_active=1 ORDER BY name ASC")
        for eid, name, code in c.fetchall():
            combo_emp.addItem(f"{name}" + (f" ({code})" if code else ""), eid)
        conn.close()

        idx = combo_emp.findData(current_emp_id)
        if idx >= 0:
            combo_emp.setCurrentIndex(idx)

        date_edit = QDateEdit()
        date_edit.setCalendarPopup(True)
        date_edit.setDisplayFormat("yyyy-MM-dd")
        date_edit.setDate(QDate.fromString(current_date, "yyyy-MM-dd"))

        amount_spin = QDoubleSpinBox()
        amount_spin.setRange(1.0, 500000.0)
        amount_spin.setDecimals(2)
        amount_spin.setValue(current_amount)
        amount_spin.setSuffix(" ج.م")

        notes_edit = QLineEdit(current_notes or "")

        form.addRow("العامل:", combo_emp)
        form.addRow("التاريخ:", date_edit)
        form.addRow("المبلغ:", amount_spin)
        form.addRow("الملاحظات:", notes_edit)
        layout.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("💾 حفظ التعديل")
        btns.button(QDialogButtonBox.Cancel).setText("❌ إلغاء")
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        layout.addWidget(btns)

        if dlg.exec() == QDialog.Accepted:
            new_eid = combo_emp.currentData()
            new_amt = amount_spin.value()
            new_date = date_edit.date().toString("yyyy-MM-dd")
            new_notes = notes_edit.text().strip()
            res = update_monthly_advance(advance_id, new_eid, new_amt, new_date, new_notes)
            if res.get("success"):
                self.show_toast("✓ تم تعديل السلفة بنجاح")
                self.load_monthly_advances_table()
            else:
                QMessageBox.critical(self, "خطأ", f"تعذر تعديل السلفة: {res.get('error')}")

    def delete_monthly_advance_handler(self, advance_id: int, emp_name: str, amount: float):
        """تأكيد وحذف سلفة شهرية."""
        reply = QMessageBox.question(
            self, "تأكيد الحذف",
            f"هل تريد بالتأكيد حذف سلفة العامل '{emp_name}' بمبلغ {amount:,.2f} ج.م نهائياً؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            res = delete_monthly_advance(advance_id)
            if res.get("success"):
                self.show_toast("✓ تم حذف السلفة بنجاح")
                self.load_monthly_advances_table()
            else:
                QMessageBox.critical(self, "خطأ", f"تعذر الحذف: {res.get('error')}")

    def export_monthly_advances_excel(self):
        """تصدير تقرير السلف الشهرية إلى Excel."""
        if not hasattr(self, "monthly_advances_table") or self.monthly_advances_table.rowCount() == 0:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات سلف لتصديرها.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "حفظ تقرير السلف الشهرية",
            f"تقرير_السلف_الشهرية_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "السلف الشهرية"
            ws.views.sheetView[0].rightToLeft = True

            # العنوان
            ws.merge_cells("A1:G1")
            title_cell = ws["A1"]
            title_cell.value = f"تقرير السلف الشهرية للعاملين - الفترة من {self.adv_filter_from.date().toString('yyyy-MM-dd')} إلى {self.adv_filter_to.date().toString('yyyy-MM-dd')}"
            title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 35

            headers = ["م", "التاريخ", "كود العامل", "اسم العامل", "القسم", "مبلغ السلفة (ج.م)", "الملاحظات"]
            for col_idx, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col_idx, value=h)
                cell.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[2].height = 25

            thin_border = Border(
                left=Side(style='thin', color='CBD5E1'),
                right=Side(style='thin', color='CBD5E1'),
                top=Side(style='thin', color='CBD5E1'),
                bottom=Side(style='thin', color='CBD5E1')
            )

            row_out = 3
            tot_amt = 0.0
            for r in range(self.monthly_advances_table.rowCount()):
                if self.monthly_advances_table.isRowHidden(r):
                    continue
                ws.cell(row=row_out, column=1, value=row_out-2).alignment = Alignment(horizontal="center")
                ws.cell(row=row_out, column=2, value=self.monthly_advances_table.item(r, 1).text()).alignment = Alignment(horizontal="center")
                ws.cell(row=row_out, column=3, value=self.monthly_advances_table.item(r, 2).text()).alignment = Alignment(horizontal="center")
                ws.cell(row=row_out, column=4, value=self.monthly_advances_table.item(r, 3).text()).alignment = Alignment(horizontal="right")
                ws.cell(row=row_out, column=5, value=self.monthly_advances_table.item(r, 4).text()).alignment = Alignment(horizontal="center")
                
                amt_str = self.monthly_advances_table.item(r, 5).text().replace(" ج.م", "").replace(",", "")
                amt_val = float(amt_str) if amt_str else 0.0
                tot_amt += amt_val
                cell_amt = ws.cell(row=row_out, column=6, value=amt_val)
                cell_amt.number_format = '#,##0.00'
                cell_amt.alignment = Alignment(horizontal="center")
                cell_amt.font = Font(bold=True, color="B91C1C")

                ws.cell(row=row_out, column=7, value=self.monthly_advances_table.item(r, 6).text()).alignment = Alignment(horizontal="right")

                for c in range(1, 8):
                    ws.cell(row=row_out, column=c).border = thin_border
                ws.row_dimensions[row_out].height = 20
                row_out += 1

            # سطر الإجمالي
            ws.merge_cells(start_row=row_out, start_column=1, end_row=row_out, end_column=5)
            tot_lbl = ws.cell(row=row_out, column=1, value="الإجمالي العام للسلف")
            tot_lbl.font = Font(bold=True, size=11, color="1E3A8A")
            tot_lbl.alignment = Alignment(horizontal="center")
            tot_lbl.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

            tot_cell = ws.cell(row=row_out, column=6, value=tot_amt)
            tot_cell.font = Font(bold=True, size=11, color="B91C1C")
            tot_cell.number_format = '#,##0.00'
            tot_cell.alignment = Alignment(horizontal="center")
            tot_cell.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

            ws.cell(row=row_out, column=7, value="").fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
            for c in range(1, 8):
                ws.cell(row=row_out, column=c).border = thin_border

            # ضبط عرض الأعمدة
            col_widths = [8, 15, 14, 28, 18, 20, 30]
            for i, w in enumerate(col_widths, 1):
                col_letter = openpyxl.utils.get_column_letter(i)
                ws.column_dimensions[col_letter].width = w

            wb.save(file_path)
            QMessageBox.information(self, "نجاح التصدير", f"تم تصدير تقرير السلف بنجاح إلى:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "خطأ في التصدير", f"تعذر تصدير التقرير: {str(e)}")


    # ══════════════════════════════════════════════════════════════════════════
    # ── 2. تبويب السلف طويلة الأجل (Long-Term Loans & Installments) ───────────
    # ══════════════════════════════════════════════════════════════════════════

    def setup_long_term_loans_tab(self) -> QWidget:
        tab = QWidget()
        main_layout = QVBoxLayout(tab)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(10)

        # ── 1. البطاقات الإحصائية العلوية (Metrics) ─────────────────────────
        metrics_frame = QFrame()
        metrics_frame.setStyleSheet("background: #f8fafc; border-radius: 8px; padding: 4px; border: 1px solid #e2e8f0;")
        m_layout = QHBoxLayout(metrics_frame)
        m_layout.setContentsMargins(6, 4, 6, 4)
        m_layout.setSpacing(12)

        self.loan_card_total = self._create_loan_metric_card("🏦 إجمالي مبالغ القروض", "0.00 ج.م", "#2563eb", "#eff6ff")
        self.loan_card_paid = self._create_loan_metric_card("✅ إجمالي المسدد", "0.00 ج.م", "#059669", "#f0fdf4")
        self.loan_card_remaining = self._create_loan_metric_card("⏳ إجمالي المتبقي (الديون القائمة)", "0.00 ج.م", "#dc2626", "#fef2f2")
        self.loan_card_active = self._create_loan_metric_card("📌 عدد السلف النشطة", "0", "#7c3aed", "#f5f3ff")

        m_layout.addWidget(self.loan_card_total)
        m_layout.addWidget(self.loan_card_paid)
        m_layout.addWidget(self.loan_card_remaining)
        m_layout.addWidget(self.loan_card_active)
        main_layout.addWidget(metrics_frame)

        # ── 2. نموذج منح سلفة طويلة الأجل جديدة ─────────────────────────────
        form_frame = QFrame()
        form_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1.5px solid #cbd5e1;
                border-radius: 8px;
                padding: 12px;
            }
        """)
        f_layout = QVBoxLayout(form_frame)
        f_layout.setContentsMargins(10, 10, 10, 10)
        f_layout.setSpacing(12)

        form_title = QLabel("➕ منح سلفة طويلة الأجل جديدة (تُخصم أقساطها الشهرية تلقائياً من الرواتب)")
        form_title.setFont(QFont("Segoe UI", 11, QFont.Bold))
        form_title.setStyleSheet("color: #1e3a8a; border: none; margin-bottom: 2px;")
        f_layout.addWidget(form_title)

        # ── السطر الأول: العامل، إجمالي السلفة، القسط الشهري ─────────────────
        row1_layout = QHBoxLayout()
        row1_layout.setSpacing(18)

        # العامل
        emp_box = QHBoxLayout()
        emp_box.setSpacing(8)
        lbl_emp = QLabel("👤 العامل:")
        lbl_emp.setStyleSheet("font-weight: bold; border: none; font-size: 13px; color: #1e293b;")
        self.loan_emp_combo = QComboBox()
        self.loan_emp_combo.setMinimumWidth(260)
        self.loan_emp_combo.setMinimumHeight(34)
        emp_box.addWidget(lbl_emp)
        emp_box.addWidget(self.loan_emp_combo)

        # إجمالي مبلغ السلفة
        tot_box = QHBoxLayout()
        tot_box.setSpacing(8)
        lbl_tot = QLabel("💰 إجمالي مبلغ السلفة:")
        lbl_tot.setStyleSheet("font-weight: bold; border: none; font-size: 13px; color: #1e293b;")
        self.loan_total_spin = QDoubleSpinBox()
        self.loan_total_spin.setRange(100.0, 1000000.0)
        self.loan_total_spin.setDecimals(2)
        self.loan_total_spin.setValue(5000.0)
        self.loan_total_spin.setMinimumWidth(150)
        self.loan_total_spin.setMinimumHeight(34)
        self.loan_total_spin.setSuffix(" ج.م")
        tot_box.addWidget(lbl_tot)
        tot_box.addWidget(self.loan_total_spin)

        # القسط الشهري
        inst_box = QHBoxLayout()
        inst_box.setSpacing(8)
        lbl_inst = QLabel("📅 قيمة القسط الشهري:")
        lbl_inst.setStyleSheet("font-weight: bold; border: none; font-size: 13px; color: #1e293b;")
        self.loan_inst_spin = QDoubleSpinBox()
        self.loan_inst_spin.setRange(50.0, 500000.0)
        self.loan_inst_spin.setDecimals(2)
        self.loan_inst_spin.setValue(500.0)
        self.loan_inst_spin.setMinimumWidth(140)
        self.loan_inst_spin.setMinimumHeight(34)
        self.loan_inst_spin.setSuffix(" ج.م")
        inst_box.addWidget(lbl_inst)
        inst_box.addWidget(self.loan_inst_spin)

        row1_layout.addLayout(emp_box, 3)
        row1_layout.addLayout(tot_box, 2)
        row1_layout.addLayout(inst_box, 2)
        f_layout.addLayout(row1_layout)

        # ── السطر الثاني: تاريخ المنح، تاريخ بدء الخصم، الملاحظات، زر الحفظ ────
        row2_layout = QHBoxLayout()
        row2_layout.setSpacing(18)

        # تاريخ المنح
        date_box = QHBoxLayout()
        date_box.setSpacing(8)
        lbl_date = QLabel("📅 تاريخ المنح:")
        lbl_date.setStyleSheet("font-weight: bold; border: none; font-size: 13px; color: #1e293b;")
        self.loan_date_picker = QDateEdit()
        self.loan_date_picker.setCalendarPopup(True)
        self.loan_date_picker.setDate(QDate.currentDate())
        self.loan_date_picker.setDisplayFormat("yyyy-MM-dd")
        self.loan_date_picker.setMinimumWidth(125)
        self.loan_date_picker.setMinimumHeight(34)
        date_box.addWidget(lbl_date)
        date_box.addWidget(self.loan_date_picker)

        # تاريخ أول قسط
        start_box = QHBoxLayout()
        start_box.setSpacing(8)
        lbl_start = QLabel("🚀 بدء الخصم:")
        lbl_start.setStyleSheet("font-weight: bold; border: none; font-size: 13px; color: #1e293b;")
        self.loan_start_picker = QDateEdit()
        self.loan_start_picker.setCalendarPopup(True)
        self.loan_start_picker.setDate(QDate.currentDate())
        self.loan_start_picker.setDisplayFormat("yyyy-MM-dd")
        self.loan_start_picker.setMinimumWidth(125)
        self.loan_start_picker.setMinimumHeight(34)
        start_box.addWidget(lbl_start)
        start_box.addWidget(self.loan_start_picker)

        # الملاحظات
        notes_box = QHBoxLayout()
        notes_box.setSpacing(8)
        lbl_notes = QLabel("📝 ملاحظات:")
        lbl_notes.setStyleSheet("font-weight: bold; border: none; font-size: 13px; color: #1e293b;")
        self.loan_notes_input = QLineEdit()
        self.loan_notes_input.setPlaceholderText("الغرض من السلفة أو توضيح اختياري...")
        self.loan_notes_input.setMinimumHeight(34)
        notes_box.addWidget(lbl_notes)
        notes_box.addWidget(self.loan_notes_input)

        # زر الحفظ
        self.loan_save_btn = QPushButton("💾 منح السلفة")
        self.loan_save_btn.setCursor(Qt.PointingHandCursor)
        self.loan_save_btn.setMinimumHeight(36)
        self.loan_save_btn.setMinimumWidth(140)
        self.loan_save_btn.setStyleSheet("""
            QPushButton {
                background-color: #059669; color: #ffffff;
                font-weight: bold; padding: 8px 18px;
                border-radius: 6px; font-size: 13px; border: none;
            }
            QPushButton:hover { background-color: #047857; }
            QPushButton:pressed { background-color: #065f46; }
        """)
        self.loan_save_btn.clicked.connect(self.add_long_term_loan_handler)

        row2_layout.addLayout(date_box, 2)
        row2_layout.addLayout(start_box, 2)
        row2_layout.addLayout(notes_box, 3)
        row2_layout.addWidget(self.loan_save_btn)
        f_layout.addLayout(row2_layout)
        main_layout.addWidget(form_frame)

        # ── 3. شريط التصفية والبحث ──────────────────────────────────────────
        filter_bar = QHBoxLayout()
        filter_bar.setSpacing(8)

        lbl_status = QLabel("📌 الحالة:")
        lbl_status.setStyleSheet("font-weight: bold;")
        self.loan_filter_status = QComboBox()
        self.loan_filter_status.addItems(["الكل", "نشطة", "مسددة بالكامل", "موقوفة"])
        self.loan_filter_status.setFixedWidth(120)

        lbl_dept = QLabel("🏢 القسم:")
        lbl_dept.setStyleSheet("font-weight: bold;")
        self.loan_filter_dept = QComboBox()
        self.loan_filter_dept.setFixedWidth(140)

        lbl_search = QLabel("🔍 بحث:")
        lbl_search.setStyleSheet("font-weight: bold;")
        self.loan_search_input = QLineEdit()
        self.loan_search_input.setPlaceholderText("ابحث باسم العامل أو الكود...")
        self.loan_search_input.setMinimumWidth(180)
        self.loan_search_input.textChanged.connect(self.filter_long_term_loans_table)

        btn_refresh = QPushButton("🔄 تحديث")
        btn_refresh.setCursor(Qt.PointingHandCursor)
        btn_refresh.setStyleSheet("background: #f8fafc; color: #334155; font-weight: bold; padding: 6px 12px; border-radius: 6px; border: 1px solid #cbd5e1;")
        btn_refresh.clicked.connect(self.load_long_term_loans_table)

        btn_export = QPushButton("📥 تصدير Excel")
        btn_export.setCursor(Qt.PointingHandCursor)
        btn_export.setStyleSheet("background: #15803d; color: white; font-weight: bold; padding: 6px 14px; border-radius: 6px; border: none;")
        btn_export.clicked.connect(self.export_long_term_loans_excel)

        self.loan_filter_status.currentIndexChanged.connect(self.load_long_term_loans_table)
        self.loan_filter_dept.currentIndexChanged.connect(self.load_long_term_loans_table)

        filter_bar.addWidget(lbl_status)
        filter_bar.addWidget(self.loan_filter_status)
        filter_bar.addWidget(lbl_dept)
        filter_bar.addWidget(self.loan_filter_dept)
        filter_bar.addWidget(lbl_search)
        filter_bar.addWidget(self.loan_search_input)
        filter_bar.addWidget(btn_refresh)
        filter_bar.addWidget(btn_export)
        main_layout.addLayout(filter_bar)

        # ── 4. جدول السلف طويلة الأجل ─────────────────────────────────────────
        self.long_term_loans_table = QTableWidget()
        headers = [
            "ID", "تاريخ المنح", "الكود", "اسم العامل", "القسم",
            "إجمالي السلفة", "القسط الشهري", "تاريخ بدء الخصم",
            "المسدد", "المتبقي", "نسبة السداد", "الحالة",
            "سجل الأقساط / سداد", "تعديل", "حذف"
        ]
        self.long_term_loans_table.setColumnCount(len(headers))
        self.long_term_loans_table.setHorizontalHeaderLabels(headers)
        self.long_term_loans_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        for idx in [0, 1, 2, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]:
            self.long_term_loans_table.horizontalHeader().setSectionResizeMode(idx, QHeaderView.ResizeToContents)
        self.long_term_loans_table.setAlternatingRowColors(True)
        self.long_term_loans_table.verticalHeader().setVisible(False)
        self.long_term_loans_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.long_term_loans_table.setStyleSheet("""
            QTableWidget {
                border: 1.5px solid #cbd5e1; border-radius: 8px; background: #ffffff;
                gridline-color: #e2e8f0; alternate-background-color: #f8fafc;
            }
            QHeaderView::section {
                background-color: #0f172a; color: #ffffff; padding: 9px 6px;
                font-weight: bold; font-size: 12px; border: 1px solid #1e293b;
            }
            QTableWidget::item { padding: 6px; font-size: 12px; color: #0f172a; }
        """)
        main_layout.addWidget(self.long_term_loans_table)

        return tab

    def populate_loan_emp_combo(self):
        """تعبئة قائمة العمال في تبويب السلف طويلة الأجل."""
        if not hasattr(self, "loan_emp_combo") or self.loan_emp_combo is None:
            return
        self.loan_emp_combo.clear()
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            SELECT e.id, e.name, COALESCE(e.code, ''), COALESCE(d.name, 'عام')
            FROM employees e
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE e.is_active = 1
            ORDER BY e.name ASC
        """)
        rows = c.fetchall()
        conn.close()
        for emp_id, name, code, dept in rows:
            label = f"{name} ({dept})" + (f" - كود: {code}" if code else "")
            self.loan_emp_combo.addItem(label, emp_id)

    def populate_loan_dept_filter(self):
        """تعبئة قائمة الأقسام في فلتر السلف طويلة الأجل."""
        if not hasattr(self, "loan_filter_dept") or self.loan_filter_dept is None:
            return
        curr_text = self.loan_filter_dept.currentText()
        self.loan_filter_dept.blockSignals(True)
        self.loan_filter_dept.clear()
        self.loan_filter_dept.addItem("🏢 جميع الأقسام", 0)
        depts = get_departments()
        for d in depts:
            self.loan_filter_dept.addItem(d[1], d[0])
        idx = self.loan_filter_dept.findText(curr_text)
        if idx >= 0:
            self.loan_filter_dept.setCurrentIndex(idx)
        self.loan_filter_dept.blockSignals(False)

    def add_long_term_loan_handler(self):
        """معالج إضافة سلفة طويلة الأجل جديدة."""
        emp_id = self.loan_emp_combo.currentData()
        if not emp_id:
            QMessageBox.warning(self, "تنبيه", "يرجى اختيار عامل أولاً.")
            return
        tot_amt = self.loan_total_spin.value()
        inst_amt = self.loan_inst_spin.value()
        if tot_amt <= 0 or inst_amt <= 0:
            QMessageBox.warning(self, "تنبيه", "يرجى إدخال مبالغ صحيحة للسلفة والقسط.")
            return
        if inst_amt > tot_amt:
            QMessageBox.warning(self, "تنبيه", "قيمة القسط الشهري لا يمكن أن تتجاوز إجمالي مبلغ السلفة.")
            return

        grant_date = self.loan_date_picker.date().toString("yyyy-MM-dd")
        start_date = self.loan_start_picker.date().toString("yyyy-MM-dd")
        notes = self.loan_notes_input.text().strip()

        res = add_long_term_loan(emp_id, tot_amt, inst_amt, start_date, grant_date, notes)
        if res.get("success"):
            self.show_toast(f"✓ تم تسجيل سلفة طويلة الأجل بمبلغ {tot_amt:,.2f} ج.م وقسط {inst_amt:,.2f} ج.م بنجاح")
            self.loan_notes_input.clear()
            self.load_long_term_loans_table()
        else:
            QMessageBox.critical(self, "خطأ", f"تعذر تسجيل السلفة: {res.get('error')}")

    def load_long_term_loans_table(self):
        """تحميل وعرض السلف طويلة الأجل في الجدول."""
        if not hasattr(self, "long_term_loans_table") or self.long_term_loans_table is None:
            return

        status = self.loan_filter_status.currentText() if hasattr(self, "loan_filter_status") else "الكل"
        dept_id = self.loan_filter_dept.currentData() if hasattr(self, "loan_filter_dept") else None

        records = get_long_term_loans(status=status, department_id=dept_id)
        self.long_term_loans_table.setRowCount(len(records))

        total_granted = 0.0
        total_paid = 0.0
        total_remaining = 0.0
        active_count = 0

        for row_idx, r in enumerate(records):
            lid, ldate, emp_id, emp_name, emp_code, dept_name, tot_amt, inst_amt, start_d, rem_amt, st, notes, paid_amt, created_at, d_id = r

            total_granted += tot_amt
            total_paid += paid_amt
            total_remaining += rem_amt
            if st == "نشطة":
                active_count += 1

            item_id = QTableWidgetItem(str(lid))
            item_id.setTextAlignment(Qt.AlignCenter)

            item_date = QTableWidgetItem(ldate)
            item_date.setTextAlignment(Qt.AlignCenter)

            item_code = QTableWidgetItem(emp_code or "-")
            item_code.setTextAlignment(Qt.AlignCenter)

            item_name = QTableWidgetItem(emp_name)
            item_name.setFont(QFont("Segoe UI", 10, QFont.Bold))

            item_dept = QTableWidgetItem(dept_name or "عام")
            item_dept.setTextAlignment(Qt.AlignCenter)

            item_tot = QTableWidgetItem(f"{tot_amt:,.2f}")
            item_tot.setTextAlignment(Qt.AlignCenter)
            item_tot.setFont(QFont("Segoe UI", 10, QFont.Bold))

            item_inst = QTableWidgetItem(f"{inst_amt:,.2f}")
            item_inst.setTextAlignment(Qt.AlignCenter)
            item_inst.setForeground(QColor("#2563eb"))
            item_inst.setFont(QFont("Segoe UI", 10, QFont.Bold))

            item_start = QTableWidgetItem(start_d)
            item_start.setTextAlignment(Qt.AlignCenter)

            item_paid = QTableWidgetItem(f"{paid_amt:,.2f}")
            item_paid.setTextAlignment(Qt.AlignCenter)
            item_paid.setForeground(QColor("#059669"))
            item_paid.setFont(QFont("Segoe UI", 10, QFont.Bold))

            item_rem = QTableWidgetItem(f"{rem_amt:,.2f}")
            item_rem.setTextAlignment(Qt.AlignCenter)
            item_rem.setForeground(QColor("#dc2626") if rem_amt > 0 else QColor("#059669"))
            item_rem.setFont(QFont("Segoe UI", 10, QFont.Bold))

            # نسبة السداد
            pct = (paid_amt / tot_amt * 100) if tot_amt > 0 else 100
            item_pct = QTableWidgetItem(f"{pct:.1f}%")
            item_pct.setTextAlignment(Qt.AlignCenter)
            item_pct.setFont(QFont("Segoe UI", 9, QFont.Bold))

            # الحالة
            item_st = QTableWidgetItem(st)
            item_st.setTextAlignment(Qt.AlignCenter)
            item_st.setFont(QFont("Segoe UI", 9, QFont.Bold))
            if st == "نشطة":
                item_st.setForeground(QColor("#2563eb"))
            elif st == "مسددة بالكامل":
                item_st.setForeground(QColor("#059669"))
            else:
                item_st.setForeground(QColor("#64748b"))

            self.long_term_loans_table.setItem(row_idx, 0, item_id)
            self.long_term_loans_table.setItem(row_idx, 1, item_date)
            self.long_term_loans_table.setItem(row_idx, 2, item_code)
            self.long_term_loans_table.setItem(row_idx, 3, item_name)
            self.long_term_loans_table.setItem(row_idx, 4, item_dept)
            self.long_term_loans_table.setItem(row_idx, 5, item_tot)
            self.long_term_loans_table.setItem(row_idx, 6, item_inst)
            self.long_term_loans_table.setItem(row_idx, 7, item_start)
            self.long_term_loans_table.setItem(row_idx, 8, item_paid)
            self.long_term_loans_table.setItem(row_idx, 9, item_rem)
            self.long_term_loans_table.setItem(row_idx, 10, item_pct)
            self.long_term_loans_table.setItem(row_idx, 11, item_st)

            # أزرار الإجراءات
            btn_details = QPushButton("📋 الأقساط / سداد")
            btn_details.setCursor(Qt.PointingHandCursor)
            btn_details.setStyleSheet("background: #059669; color: white; border-radius: 4px; padding: 4px 8px; font-weight: bold; border: none;")
            btn_details.clicked.connect(lambda chk=False, loan_id=lid: self.show_loan_details_dialog(loan_id))

            btn_edit = QPushButton("✏️ تعديل")
            btn_edit.setCursor(Qt.PointingHandCursor)
            btn_edit.setStyleSheet("background: #2563eb; color: white; border-radius: 4px; padding: 4px 8px; font-weight: bold; border: none;")
            btn_edit.clicked.connect(lambda chk=False, loan_id=lid, inst=inst_amt, s=st, n=notes: self.edit_long_term_loan_dialog(loan_id, inst, s, n))

            btn_del = QPushButton("🗑️ حذف")
            btn_del.setCursor(Qt.PointingHandCursor)
            btn_del.setStyleSheet("background: #ef4444; color: white; border-radius: 4px; padding: 4px 8px; font-weight: bold; border: none;")
            btn_del.clicked.connect(lambda chk=False, loan_id=lid, en=emp_name, amt=tot_amt: self.delete_long_term_loan_handler(loan_id, en, amt))

            self.long_term_loans_table.setCellWidget(row_idx, 12, btn_details)
            self.long_term_loans_table.setCellWidget(row_idx, 13, btn_edit)
            self.long_term_loans_table.setCellWidget(row_idx, 14, btn_del)

        # تحديث البطاقات
        if hasattr(self, "loan_card_total"):
            self._update_loan_card(self.loan_card_total, f"{total_granted:,.2f} ج.م")
            self._update_loan_card(self.loan_card_paid, f"{total_paid:,.2f} ج.م")
            self._update_loan_card(self.loan_card_remaining, f"{total_remaining:,.2f} ج.م")
            self._update_loan_card(self.loan_card_active, str(active_count))

        self.filter_long_term_loans_table()

    def filter_long_term_loans_table(self):
        """تصفية جدول السلف طويلة الأجل بالبحث النصي."""
        if not hasattr(self, "long_term_loans_table") or not hasattr(self, "loan_search_input"):
            return
        query = self.loan_search_input.text().strip().lower()
        for r in range(self.long_term_loans_table.rowCount()):
            match = False
            for c in [1, 2, 3, 4, 11]:
                item = self.long_term_loans_table.item(r, c)
                if item and query in item.text().lower():
                    match = True
                    break
            self.long_term_loans_table.setRowHidden(r, not match)

    def show_loan_details_dialog(self, loan_id: int):
        """نافذة استعراض سجل أقساط السلفة وتسجيل دفعات يدوية أو سداد كامل."""
        loan = get_long_term_loan_by_id(loan_id)
        if not loan:
            QMessageBox.warning(self, "تنبيه", "لم يتم العثور على بيانات السلفة.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f"🏦 كشف حساب وأقساط سلفة: {loan['emp_name']}")
        dlg.setLayoutDirection(Qt.RightToLeft)
        dlg.resize(650, 520)
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # معلومات السلفة
        info_frame = QFrame()
        info_frame.setStyleSheet("background: #f8fafc; border-radius: 8px; padding: 10px; border: 1px solid #cbd5e1;")
        i_layout = QFormLayout(info_frame)
        i_layout.setSpacing(6)

        pct = (loan["paid_amount"] / loan["total_amount"] * 100) if loan["total_amount"] > 0 else 100

        i_layout.addRow("العامل:", QLabel(f"<b>{loan['emp_name']}</b> (القسم: {loan['dept_name']})"))
        i_layout.addRow("إجمالي مبلغ السلفة:", QLabel(f"<b>{loan['total_amount']:,.2f} ج.م</b> (القسط الشهري: {loan['installment_amount']:,.2f} ج.م)"))
        i_layout.addRow("المبلغ المسدد:", QLabel(f"<b style='color:#059669;'>{loan['paid_amount']:,.2f} ج.م</b> | المتبقي: <b style='color:#dc2626;'>{loan['remaining_amount']:,.2f} ج.م</b> (الحالة: <b>{loan['status']}</b>)"))

        p_bar = QProgressBar()
        p_bar.setRange(0, 100)
        p_bar.setValue(int(pct))
        p_bar.setFormat(f"نسبة السداد: {pct:.1f}%")
        p_bar.setStyleSheet("""
            QProgressBar { border: 1px solid #cbd5e1; border-radius: 6px; text-align: center; font-weight: bold; background: #e2e8f0; height: 18px; }
            QProgressBar::chunk { background-color: #059669; border-radius: 5px; }
        """)
        i_layout.addRow("مؤشر السداد:", p_bar)
        layout.addWidget(info_frame)

        # قسم تسجيل دفعة يدوية
        pay_frame = QFrame()
        pay_frame.setStyleSheet("background: #ffffff; border: 1.5px dashed #2563eb; border-radius: 8px; padding: 8px;")
        p_layout = QHBoxLayout(pay_frame)
        p_layout.setContentsMargins(6, 4, 6, 4)
        p_layout.setSpacing(8)

        lbl_pay_title = QLabel("💳 سداد دفعة/قسط:")
        lbl_pay_title.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_pay_title.setStyleSheet("color: #2563eb; border: none;")

        pay_amt_spin = QDoubleSpinBox()
        pay_amt_spin.setRange(1.0, loan["remaining_amount"] if loan["remaining_amount"] > 0 else 100000.0)
        pay_amt_spin.setDecimals(2)
        pay_amt_spin.setValue(min(loan["installment_amount"], loan["remaining_amount"]) if loan["remaining_amount"] > 0 else 0)
        pay_amt_spin.setSuffix(" ج.م")
        pay_amt_spin.setFixedWidth(110)

        pay_date_picker = QDateEdit()
        pay_date_picker.setCalendarPopup(True)
        pay_date_picker.setDate(QDate.currentDate())
        pay_date_picker.setDisplayFormat("yyyy-MM-dd")
        pay_date_picker.setFixedWidth(105)

        pay_notes = QLineEdit()
        pay_notes.setPlaceholderText("ملاحظات السداد...")

        btn_record_pay = QPushButton("💾 تسجيل السداد")
        btn_record_pay.setCursor(Qt.PointingHandCursor)
        btn_record_pay.setStyleSheet("background: #059669; color: white; font-weight: bold; padding: 6px 12px; border-radius: 4px; border: none;")

        p_layout.addWidget(lbl_pay_title)
        p_layout.addWidget(pay_date_picker)
        p_layout.addWidget(pay_amt_spin)
        p_layout.addWidget(pay_notes)
        p_layout.addWidget(btn_record_pay)
        layout.addWidget(pay_frame)

        # جدول سجل الأقساط المسددة
        lbl_tbl = QLabel("📜 سجل الأقساط والدفعات المسددة:")
        lbl_tbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        layout.addWidget(lbl_tbl)

        inst_table = QTableWidget()
        inst_table.setColumnCount(6)
        inst_table.setHorizontalHeaderLabels(["ID", "تاريخ السداد", "المبلغ (ج.م)", "عن شهر / سنة", "ملاحظات", "حذف"])
        inst_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        for ci in [0, 1, 2, 3, 5]:
            inst_table.horizontalHeader().setSectionResizeMode(ci, QHeaderView.ResizeToContents)
        inst_table.setAlternatingRowColors(True)
        inst_table.verticalHeader().setVisible(False)
        layout.addWidget(inst_table)

        def reload_installments():
            inst_rows = get_loan_installments(loan_id=loan_id)
            inst_table.setRowCount(len(inst_rows))
            for rx, ir in enumerate(inst_rows):
                iid, lid, eid, ename, iamt, idate, pmonth, pyear, inotes, icreated = ir
                
                item_iid = QTableWidgetItem(str(iid))
                item_iid.setTextAlignment(Qt.AlignCenter)

                item_idate = QTableWidgetItem(idate)
                item_idate.setTextAlignment(Qt.AlignCenter)

                item_iamt = QTableWidgetItem(f"{iamt:,.2f}")
                item_iamt.setTextAlignment(Qt.AlignCenter)
                item_iamt.setFont(QFont("Segoe UI", 10, QFont.Bold))
                item_iamt.setForeground(QColor("#059669"))

                p_str = f"{pmonth}/{pyear}" if (pmonth > 0 and pyear > 0) else "سداد يدوي"
                item_p = QTableWidgetItem(p_str)
                item_p.setTextAlignment(Qt.AlignCenter)

                item_in = QTableWidgetItem(inotes or "-")

                btn_del_inst = QPushButton("🗑️")
                btn_del_inst.setCursor(Qt.PointingHandCursor)
                btn_del_inst.setStyleSheet("background: #ef4444; color: white; border-radius: 4px; padding: 2px 6px; font-weight: bold; border: none;")
                btn_del_inst.clicked.connect(lambda chk=False, inst_id=iid: delete_inst_handler(inst_id))

                inst_table.setItem(rx, 0, item_iid)
                inst_table.setItem(rx, 1, item_idate)
                inst_table.setItem(rx, 2, item_iamt)
                inst_table.setItem(rx, 3, item_p)
                inst_table.setItem(rx, 4, item_in)
                inst_table.setCellWidget(rx, 5, btn_del_inst)

        def delete_inst_handler(inst_id: int):
            if QMessageBox.question(dlg, "تأكيد", "هل تريد حذف هذا القسط واستعادة قيمته للمتبقي؟") == QMessageBox.Yes:
                res = delete_loan_installment(inst_id)
                if res.get("success"):
                    self.show_toast("✓ تم حذف القسط وتحديث الرصيد المتبقي")
                    reload_installments()
                    self.load_long_term_loans_table()
                    dlg.close()
                    self.show_loan_details_dialog(loan_id)

        def record_payment_action():
            p_amt = pay_amt_spin.value()
            if p_amt <= 0:
                QMessageBox.warning(dlg, "تنبيه", "يرجى إدخال مبلغ صحيح.")
                return
            p_date = pay_date_picker.date().toString("yyyy-MM-dd")
            p_nt = pay_notes.text().strip()
            res = pay_loan_installment(loan_id, p_amt, p_date, notes=p_nt)
            if res.get("success"):
                self.show_toast(f"✓ تم تسجيل سداد مبلغ {p_amt:,.2f} ج.م بنجاح")
                reload_installments()
                self.load_long_term_loans_table()
                dlg.close()
                self.show_loan_details_dialog(loan_id)
            else:
                QMessageBox.critical(dlg, "خطأ", f"تعذر تسجيل السداد: {res.get('error')}")

        btn_record_pay.clicked.connect(record_payment_action)
        reload_installments()

        btn_close = QPushButton("✖ إغلاق")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setStyleSheet("background: #64748b; color: white; font-weight: bold; padding: 6px 16px; border-radius: 6px; border: none;")
        btn_close.clicked.connect(dlg.close)
        layout.addWidget(btn_close, alignment=Qt.AlignLeft)

        dlg.exec()

    def edit_long_term_loan_dialog(self, loan_id: int, current_inst: float, current_status: str, current_notes: str):
        """نافذة تعديل سلفة طويلة الأجل."""
        dlg = QDialog(self)
        dlg.setWindowTitle("✏️ تعديل بيانات السلفة طويلة الأجل")
        dlg.setLayoutDirection(Qt.RightToLeft)
        dlg.setMinimumWidth(380)
        layout = QVBoxLayout(dlg)

        form = QFormLayout()
        form.setSpacing(10)

        inst_spin = QDoubleSpinBox()
        inst_spin.setRange(10.0, 500000.0)
        inst_spin.setDecimals(2)
        inst_spin.setValue(current_inst)
        inst_spin.setSuffix(" ج.م")

        status_combo = QComboBox()
        status_combo.addItems(["نشطة", "مسددة بالكامل", "موقوفة"])
        idx = status_combo.findText(current_status)
        if idx >= 0:
            status_combo.setCurrentIndex(idx)

        notes_edit = QLineEdit(current_notes or "")

        form.addRow("قيمة القسط الشهري:", inst_spin)
        form.addRow("حالة السلفة:", status_combo)
        form.addRow("الملاحظات:", notes_edit)
        layout.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("💾 حفظ التعديل")
        btns.button(QDialogButtonBox.Cancel).setText("❌ إلغاء")
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        layout.addWidget(btns)

        if dlg.exec() == QDialog.Accepted:
            new_inst = inst_spin.value()
            new_st = status_combo.currentText()
            new_nt = notes_edit.text().strip()
            res = update_long_term_loan(loan_id, new_inst, new_st, new_nt)
            if res.get("success"):
                self.show_toast("✓ تم تعديل السلفة بنجاح")
                self.load_long_term_loans_table()
            else:
                QMessageBox.critical(self, "خطأ", f"تعذر التعديل: {res.get('error')}")

    def delete_long_term_loan_handler(self, loan_id: int, emp_name: str, total_amt: float):
        """حذف سلفة طويلة الأجل."""
        reply = QMessageBox.question(
            self, "تأكيد الحذف",
            f"هل تريد بالتأكيد حذف سلفة العامل '{emp_name}' بمبلغ {total_amt:,.2f} ج.م مع جميع سجلات أقساطها نهائياً؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            res = delete_long_term_loan(loan_id)
            if res.get("success"):
                self.show_toast("✓ تم حذف السلفة بنجاح")
                self.load_long_term_loans_table()
            else:
                QMessageBox.critical(self, "خطأ", f"تعذر الحذف: {res.get('error')}")

    def export_long_term_loans_excel(self):
        """تصدير تقرير السلف طويلة الأجل والأقساط إلى Excel."""
        if not hasattr(self, "long_term_loans_table") or self.long_term_loans_table.rowCount() == 0:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات سلف لتصديرها.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "حفظ تقرير السلف طويلة الأجل",
            f"تقرير_السلف_طويلة_الأجل_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "السلف طويلة الأجل"
            ws.views.sheetView[0].rightToLeft = True

            # العنوان
            ws.merge_cells("A1:K1")
            title_cell = ws["A1"]
            title_cell.value = "تقرير السلف طويلة الأجل والأقساط الشهرية للعاملين"
            title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 35

            headers = ["م", "تاريخ المنح", "كود العامل", "اسم العامل", "القسم", "إجمالي السلفة (ج.م)", "القسط الشهري", "تاريخ بدء الخصم", "المسدد", "المتبقي", "الحالة"]
            for col_idx, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col_idx, value=h)
                cell.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="059669", end_color="059669", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[2].height = 25

            thin_border = Border(
                left=Side(style='thin', color='CBD5E1'),
                right=Side(style='thin', color='CBD5E1'),
                top=Side(style='thin', color='CBD5E1'),
                bottom=Side(style='thin', color='CBD5E1')
            )

            row_out = 3
            tot_grant = 0.0
            tot_paid = 0.0
            tot_rem = 0.0

            for r in range(self.long_term_loans_table.rowCount()):
                if self.long_term_loans_table.isRowHidden(r):
                    continue
                ws.cell(row=row_out, column=1, value=row_out-2).alignment = Alignment(horizontal="center")
                ws.cell(row=row_out, column=2, value=self.long_term_loans_table.item(r, 1).text()).alignment = Alignment(horizontal="center")
                ws.cell(row=row_out, column=3, value=self.long_term_loans_table.item(r, 2).text()).alignment = Alignment(horizontal="center")
                ws.cell(row=row_out, column=4, value=self.long_term_loans_table.item(r, 3).text()).alignment = Alignment(horizontal="right")
                ws.cell(row=row_out, column=5, value=self.long_term_loans_table.item(r, 4).text()).alignment = Alignment(horizontal="center")

                # إجمالي السلفة
                v_tot = float(self.long_term_loans_table.item(r, 5).text().replace(",", ""))
                tot_grant += v_tot
                c_tot = ws.cell(row=row_out, column=6, value=v_tot)
                c_tot.number_format = '#,##0.00'
                c_tot.alignment = Alignment(horizontal="center")

                # القسط
                v_inst = float(self.long_term_loans_table.item(r, 6).text().replace(",", ""))
                c_inst = ws.cell(row=row_out, column=7, value=v_inst)
                c_inst.number_format = '#,##0.00'
                c_inst.alignment = Alignment(horizontal="center")

                # تاريخ البدء
                ws.cell(row=row_out, column=8, value=self.long_term_loans_table.item(r, 7).text()).alignment = Alignment(horizontal="center")

                # المسدد
                v_paid = float(self.long_term_loans_table.item(r, 8).text().replace(",", ""))
                tot_paid += v_paid
                c_paid = ws.cell(row=row_out, column=9, value=v_paid)
                c_paid.number_format = '#,##0.00'
                c_paid.alignment = Alignment(horizontal="center")
                c_paid.font = Font(color="059669", bold=True)

                # المتبقي
                v_rem = float(self.long_term_loans_table.item(r, 9).text().replace(",", ""))
                tot_rem += v_rem
                c_rem = ws.cell(row=row_out, column=10, value=v_rem)
                c_rem.number_format = '#,##0.00'
                c_rem.alignment = Alignment(horizontal="center")
                c_rem.font = Font(color="DC2626", bold=True)

                # الحالة
                ws.cell(row=row_out, column=11, value=self.long_term_loans_table.item(r, 11).text()).alignment = Alignment(horizontal="center")

                for c in range(1, 12):
                    ws.cell(row=row_out, column=c).border = thin_border
                ws.row_dimensions[row_out].height = 20
                row_out += 1

            # سطر الإجمالي
            ws.merge_cells(start_row=row_out, start_column=1, end_row=row_out, end_column=5)
            tot_lbl = ws.cell(row=row_out, column=1, value="الإجمالي العام")
            tot_lbl.font = Font(bold=True, size=11, color="1E3A8A")
            tot_lbl.alignment = Alignment(horizontal="center")
            tot_lbl.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

            c_g_tot = ws.cell(row=row_out, column=6, value=tot_grant)
            c_g_tot.number_format = '#,##0.00'
            c_g_tot.font = Font(bold=True)
            c_g_tot.alignment = Alignment(horizontal="center")
            c_g_tot.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

            ws.cell(row=row_out, column=7, value="").fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
            ws.cell(row=row_out, column=8, value="").fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

            c_g_paid = ws.cell(row=row_out, column=9, value=tot_paid)
            c_g_paid.number_format = '#,##0.00'
            c_g_paid.font = Font(bold=True, color="059669")
            c_g_paid.alignment = Alignment(horizontal="center")
            c_g_paid.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

            c_g_rem = ws.cell(row=row_out, column=10, value=tot_rem)
            c_g_rem.number_format = '#,##0.00'
            c_g_rem.font = Font(bold=True, color="DC2626")
            c_g_rem.alignment = Alignment(horizontal="center")
            c_g_rem.fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")

            ws.cell(row=row_out, column=11, value="").fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
            for c in range(1, 12):
                ws.cell(row=row_out, column=c).border = thin_border

            wb.save(file_path)
            QMessageBox.information(self, "نجاح التصدير", f"تم تصدير تقرير السلف بنجاح إلى:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "خطأ في التصدير", f"تعذر التصدير: {str(e)}")


    # ── دالة مساعدة لإنشاء بطاقة إحصائية ──────────────────────────────────
    def _create_loan_metric_card(self, title: str, val_str: str, color_hex: str, bg_hex: str) -> QFrame:
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {bg_hex};
                border: 1.5px solid {color_hex};
                border-radius: 8px;
                padding: 6px 12px;
            }}
        """)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(2)

        lbl_t = QLabel(title)
        lbl_t.setFont(QFont("Segoe UI", 10, QFont.Bold))
        lbl_t.setStyleSheet(f"color: {color_hex}; border: none; background: transparent;")

        lbl_v = QLabel(val_str)
        lbl_v.setFont(QFont("Segoe UI", 13, QFont.Bold))
        lbl_v.setStyleSheet(f"color: {color_hex}; border: none; background: transparent;")

        layout.addWidget(lbl_t)
        layout.addWidget(lbl_v)
        card.val_label = lbl_v
        return card

    def _update_loan_card(self, card: QFrame, new_text: str):
        if hasattr(card, "val_label") and card.val_label is not None:
            card.val_label.setText(new_text)
