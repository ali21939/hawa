
from core.config import *
from core.database import *
from core.utils import *
from gui.widgets import *

from PySide6.QtCore import *
from PySide6.QtWidgets import *
from PySide6.QtGui import *
import sqlite3, os, time, csv, threading
from datetime import datetime, timedelta
import json
from core.config import *
from core.utils import *
from core.database import *
from openpyxl import Workbook
from openpyxl.styles import Font, Fill, Alignment, Border, Side, PatternFill

from contextlib import contextmanager

class TableHelpersMixin:
    DELETE_ENABLED_TABLES = {
        "attendance",
        "deductions",
        "incentives",
        "overtime",
        "employees",
        "departments",
        "credited_leaves",
    }

    @contextmanager
    def preserve_table_scroll(self, table):
        """Preserves vertical and horizontal scroll position and row selection of a QTableWidget during data reload."""
        if table is None:
            yield
            return
        v_bar = table.verticalScrollBar()
        h_bar = table.horizontalScrollBar()
        v_val = v_bar.value() if v_bar else 0
        h_val = h_bar.value() if h_bar else 0
        selected_rows = []
        try:
            if table.selectionModel():
                selected_rows = [idx.row() for idx in table.selectionModel().selectedRows()]
        except Exception:
            pass

        try:
            yield
        finally:
            if v_bar:
                v_bar.setValue(v_val)
            if h_bar:
                h_bar.setValue(h_val)
            if selected_rows:
                try:
                    for r in selected_rows:
                        if r < table.rowCount():
                            table.selectRow(r)
                except Exception:
                    pass
            QTimer.singleShot(0, lambda: (v_bar.setValue(v_val) if v_bar else None, h_bar.setValue(h_val) if h_bar else None))

    def prepare_table_with_delete(self, table, headers):
        table.setColumnCount(len(headers) + 2)
        table.setHorizontalHeaderLabels(headers + ["تعديل", "حذف"])
        table.setStyleSheet("""
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
                padding: 8px 12px;
                border-bottom: 1px solid #e2e8f0;
                color: #0f172a;
                font-size: 13px;
            }
            QTableWidget::item:selected {
                background-color: #dbeafe;
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
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        if headers and "ID" in str(headers[0]).upper():
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        
        for i, h_text in enumerate(headers):
            if any(k in str(h_text) for k in ["حالة", "نسبة", "إجراء", "الرصيد", "صورة", "الحالة"]):
                header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
                
        header.setSectionResizeMode(len(headers), QHeaderView.ResizeToContents)
        header.setSectionResizeMode(len(headers) + 1, QHeaderView.ResizeToContents)

    def prepare_table_simple(self, table, headers):
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setStyleSheet("""
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
                padding: 8px 12px;
                border-bottom: 1px solid #e2e8f0;
                color: #0f172a;
                font-size: 13px;
            }
            QTableWidget::item:selected {
                background-color: #dbeafe;
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
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        if headers and "ID" in str(headers[0]).upper():
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            
        for i, h_text in enumerate(headers):
            if any(k in str(h_text) for k in ["حالة", "نسبة", "إجراء", "الرصيد", "صورة", "الحالة"]):
                header.setSectionResizeMode(i, QHeaderView.ResizeToContents)



    def add_delete_button(self, table, row_index, column_index, table_name, record_id, refresh_callback):
        delete_btn = QPushButton("🗑️ حذف")
        delete_btn.setCursor(Qt.PointingHandCursor)
        delete_btn.setStyleSheet("""
            QPushButton { 
                background-color: #ef4444; 
                color: white; 
                border: none; 
                border-radius: 6px; 
                padding: 6px 14px; 
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover { background-color: #dc2626; }
            QPushButton:pressed { background-color: #b91c1c; }
        """)
        delete_btn.clicked.connect(
            lambda checked=False, tn=table_name, rid=record_id, cb=refresh_callback: self.confirm_delete_record(tn, rid, cb)
        )
        table.setCellWidget(row_index, column_index, delete_btn)

    def add_edit_button(self, table, row_index, column_index, table_name, record_id, refresh_callback):
        edit_btn = QPushButton("✏️ تعديل")
        edit_btn.setCursor(Qt.PointingHandCursor)
        edit_btn.setStyleSheet("""
            QPushButton { 
                background-color: #2563eb; 
                color: white; 
                border: none; 
                border-radius: 6px; 
                padding: 6px 14px; 
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover { background-color: #1d4ed8; }
            QPushButton:pressed { background-color: #1e40af; }
        """)
        edit_btn.clicked.connect(
            lambda checked=False, tn=table_name, rid=record_id, cb=refresh_callback: self.edit_assembly_record(tn, rid, cb)
        )
        table.setCellWidget(row_index, column_index, edit_btn)

    def confirm_delete_record(self, table_name, record_id, refresh_callback):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("تأكيد الحذف")
        msg_box.setText("هل تريد بالتأكيد حذف هذا السجل نهائياً؟")
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setLayoutDirection(Qt.RightToLeft)

        btn_confirm = msg_box.addButton("🗑️ تأكيد الحذف", QMessageBox.YesRole)
        btn_cancel = msg_box.addButton("❌ إلغاء", QMessageBox.NoRole)
        msg_box.setDefaultButton(btn_cancel)

        btn_confirm.setCursor(Qt.PointingHandCursor)
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_confirm.setStyleSheet("background-color: #dc2626; color: #ffffff; font-weight: bold; padding: 7px 18px; border-radius: 6px; font-size: 12px; border: none;")
        btn_cancel.setStyleSheet("background-color: #64748b; color: #ffffff; font-weight: bold; padding: 7px 18px; border-radius: 6px; font-size: 12px; border: none;")

        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: #ffffff;
            }
            QLabel {
                color: #0f172a;
                font-size: 13px;
            }
            QPushButton {
                min-width: 110px;
                min-height: 30px;
            }
        """)

        msg_box.exec()
        if msg_box.clickedButton() != btn_confirm:
            return

        try:
            self.delete_record(table_name, record_id)
        except Exception as e:
            QMessageBox.critical(self, "خطأ", str(e))
            return

        refresh_callback()
        self.refresh_dashboard_data()

    def get_table_key_column(self, cursor, table_name):
        cursor.execute(f"PRAGMA table_info({table_name})")
        table_info = cursor.fetchall()
        if not table_info:
            raise ValueError("لم يتم العثور على بيانات الجدول.")

        pk_columns = [col[1] for col in table_info if col[5]]
        return pk_columns[0] if pk_columns else table_info[0][1]

    def delete_record(self, table_name, record_id):
        if table_name not in self.DELETE_ENABLED_TABLES:
            raise ValueError("هذا الجدول غير مسموح بالحذف منه.")

        conn = sqlite3.connect(DB_NAME)
        conn.execute("PRAGMA foreign_keys = ON")
        c = conn.cursor()
        try:
            key_column = self.get_table_key_column(c, table_name)
            c.execute(f"DELETE FROM {table_name} WHERE {key_column}=?", (record_id,))
            conn.commit()
        finally:
            conn.close()

    def fetch_employees(self):
        """جلب جميع العاملين النشطين مع بيانات الأقسام."""
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("""
            SELECT e.id, e.name, e.code, e.base_salary, e.attendance_incentive,
                   e.daily_wage, e.work_hours, COALESCE(d.name, 'عام'), e.department_id
            FROM employees e
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE e.is_active = 1
            ORDER BY e.name ASC
        """)
        rows = c.fetchall()
        conn.close()
        return rows

    def edit_assembly_record(self, table_name, record_id, refresh_callback):
        """تعديل سجل مالي أو إداري."""
        if table_name == "credited_leaves" and hasattr(self, "edit_credited_leave_dialog"):
            self.edit_credited_leave_dialog(record_id, refresh_callback)
        elif hasattr(self, "edit_financial_record_dialog"):
            self.edit_financial_record_dialog(table_name, record_id, refresh_callback)
        elif hasattr(self, "edit_department_dialog") and table_name == "departments":
            self.edit_department_dialog(record_id)
        else:
            QMessageBox.information(self, "تعديل", "خاصية التعديل السريع متاحة من خلال النوافذ المخصصة.")

    def populate_employee_combo(self, combo, selected_id=None):
        current_id = selected_id if selected_id is not None else combo.currentData()
        combo.blockSignals(True)
        combo.clear()
        for row in self.fetch_employees():
            emp_id = row[0]
            name = row[1]
            dept_name = row[7] if len(row) > 7 else ""
            display_text = f"{name} ({dept_name})" if dept_name and dept_name != "عام" else name
            combo.addItem(display_text, emp_id)
        if current_id is not None:
            idx = combo.findData(current_id)
            if idx >= 0:
                combo.setCurrentIndex(idx)
        combo.blockSignals(False)

    def _write_styled_xlsx(self, file_path, title, headers, data_rows, sheet_name="تقرير", extra_info=None):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.sheet_view.rightToLeft = True

        current_row = 1
        if title:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(headers))
            title_cell = ws.cell(row=current_row, column=1, value=title)
            title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="1E293B")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            current_row += 1

        if extra_info:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(headers))
            info_cell = ws.cell(row=current_row, column=1, value=extra_info)
            info_cell.font = Font(name="Segoe UI", size=10, italic=True, color="64748B")
            info_cell.alignment = Alignment(horizontal="center", vertical="center")
            current_row += 1

        current_row += 1

        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )

        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        current_row += 1

        alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        data_font = Font(name="Segoe UI", size=10, color="0F172A")

        for r_idx, row_data in enumerate(data_rows):
            for col_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = data_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
                if r_idx % 2 == 1:
                    cell.fill = alt_fill
            current_row += 1

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        wb.save(file_path)

