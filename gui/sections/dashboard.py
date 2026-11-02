import sqlite3
from datetime import datetime
from core.config import DB_NAME
from gui.widgets import DashboardCard

from PySide6.QtCore import *
from PySide6.QtWidgets import *
from PySide6.QtGui import *


class DashboardMixin:
    def init_dashboard(self):
        self.dashboard_widget = QWidget()
        self.dashboard_widget.setObjectName("dashboard_widget")
        main_layout = QVBoxLayout(self.dashboard_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ── Header Banner ─────────────────────────────────────────
        header = QFrame()
        header.setFixedHeight(72)
        header.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 #0f172a, stop:0.5 #1e293b, stop:1 #0f172a);
                border-bottom: 2px solid #2563eb;
            }
        """)
        h_layout = QHBoxLayout(header)
        h_layout.setContentsMargins(24, 0, 24, 0)

        factory_lbl = QLabel("👥 مصنع حوا — نظام إدارة شؤون العاملين")
        factory_lbl.setFont(QFont("Segoe UI", 16, QFont.Bold))
        factory_lbl.setStyleSheet("color:#f8fafc;background:transparent;border:none;")

        self.header_date_lbl = QLabel("")
        self.header_date_lbl.setFont(QFont("Segoe UI", 10))
        self.header_date_lbl.setStyleSheet("color:#94a3b8;background:transparent;border:none;")

        self.header_time_lbl = QLabel("")
        self.header_time_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.header_time_lbl.setStyleSheet("color:#60a5fa;background:transparent;border:none;")

        btn_update = QPushButton("🚀 فحص التحديثات")
        btn_update.setFixedSize(130, 34)
        btn_update.setStyleSheet("""
            QPushButton{background:#059669;color:#fff;border-radius:8px;
                font-weight:bold;font-size:12px;border:none;}
            QPushButton:hover{background:#047857;}
        """)
        btn_update.clicked.connect(self.manual_check_for_updates)

        btn_refresh = QPushButton("🔄 تحديث")
        btn_refresh.setFixedSize(90, 34)
        btn_refresh.setStyleSheet("""
            QPushButton{background:#2563eb;color:#fff;border-radius:8px;
                font-weight:bold;font-size:12px;border:none;}
            QPushButton:hover{background:#1d4ed8;}
        """)
        btn_refresh.clicked.connect(self.refresh_dashboard_data)

        btn_lock = QPushButton("🔒 قفل")
        btn_lock.setFixedSize(80, 34)
        btn_lock.setStyleSheet("""
            QPushButton{background:#dc2626;color:#fff;border-radius:8px;
                font-weight:bold;font-size:12px;border:none;}
            QPushButton:hover{background:#b91c1c;}
        """)
        btn_lock.clicked.connect(self.lock_application)

        h_layout.addWidget(factory_lbl)
        h_layout.addStretch()
        h_layout.addWidget(self.header_date_lbl)
        h_layout.addSpacing(20)
        h_layout.addWidget(self.header_time_lbl)
        h_layout.addSpacing(20)
        h_layout.addWidget(btn_update)
        h_layout.addSpacing(8)
        h_layout.addWidget(btn_refresh)
        h_layout.addSpacing(8)
        h_layout.addWidget(btn_lock)
        main_layout.addWidget(header)

        # ── KPI Ribbon ────────────────────────────────────────────
        kpi_bar = QFrame()
        kpi_bar.setFixedHeight(54)
        kpi_bar.setStyleSheet("QFrame{background:#1e293b;border-bottom:1px solid #334155;}")
        kpi_layout = QHBoxLayout(kpi_bar)
        kpi_layout.setContentsMargins(24, 0, 24, 0)
        kpi_layout.setSpacing(32)

        def _kpi_pill(label, color):
            w = QWidget(); wl = QHBoxLayout(w); wl.setSpacing(6); wl.setContentsMargins(0,0,0,0)
            lt = QLabel(label); lt.setStyleSheet(f"color:#94a3b8;background:transparent;border:none;font-size:11px;font-weight:bold;")
            lv = QLabel("—"); lv.setStyleSheet(f"color:{color};background:transparent;border:none;font-size:13px;font-weight:bold;")
            wl.addWidget(lt); wl.addWidget(lv)
            return w, lv

        kpi_layout.addStretch()
        pill1, self.kpi_total_emp = _kpi_pill("👥 إجمالي العاملين:", "#93c5fd")
        pill2, self.kpi_present_val = _kpi_pill("✅ حضور اليوم:", "#34d399")
        pill3, self.kpi_absent_val = _kpi_pill("❌ غياب اليوم:", "#f87171")
        pill4, self.kpi_hours_val = _kpi_pill("⏰ ساعات العمل اليوم:", "#f59e0b")
        pill5, self.kpi_incentives_val = _kpi_pill("💰 مكافآت الشهر:", "#a78bfa")
        pill6, self.kpi_deductions_val = _kpi_pill("📉 خصومات الشهر:", "#fb7185")
        for p in [pill1, pill2, pill3, pill4, pill5, pill6]:
            kpi_layout.addWidget(p)
        kpi_layout.addStretch()
        main_layout.addWidget(kpi_bar)

        # ── Cards Grid ────────────────────────────────────────────
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setStyleSheet("QScrollArea{background:#f0f4f8;}")

        grid_container = QWidget()
        grid_container.setStyleSheet("background:#f0f4f8;")
        grid_layout = QGridLayout(grid_container)
        grid_layout.setContentsMargins(40, 40, 40, 40)
        grid_layout.setSpacing(24)

        self.cards = {}
        # Single focused HR Section Card
        self.card_hr = DashboardCard(
            title="شؤون العاملين",
            bg_color="#2563eb",
            icon_str="👥",
            category="الحضور، الانصراف، الساعات وشؤون الأفراد",
            tint_color="#eff6ff",
            border_color="#bfdbfe",
            click_callback=lambda: self.stacked_widget.setCurrentIndex(1)
        )
        self.card_hr.setMinimumHeight(180)
        grid_layout.addWidget(self.card_hr, 0, 0)
        self.cards["شؤون العاملين"] = self.card_hr

        scroll_area.setWidget(grid_container)
        main_layout.addWidget(scroll_area, 1)
        self.stacked_widget.addWidget(self.dashboard_widget)

        # Clock timer
        self.dashboard_clock_timer = QTimer(self)
        self.dashboard_clock_timer.timeout.connect(self.update_dashboard_clock)
        self.dashboard_clock_timer.start(1000)
        self.update_dashboard_clock()

    def is_friday(self, qdate):
        return qdate.dayOfWeek() == 5

    def update_dashboard_clock(self):
        now = datetime.now()
        days_ar = {1:"الإثنين",2:"الثلاثاء",3:"الأربعاء",4:"الخميس",5:"الجمعة",6:"السبت",7:"الأحد"}
        months_ar = {1:"يناير",2:"فبراير",3:"مارس",4:"أبريل",5:"مايو",6:"يونيو",
                     7:"يوليو",8:"أغسطس",9:"سبتمبر",10:"أكتوبر",11:"نوفمبر",12:"ديسمبر"}
        day_name = days_ar.get(now.isoweekday(), "")
        month_name = months_ar.get(now.month, "")
        date_str = f"📅 {day_name}، {now.day} {month_name} {now.year}"
        hour = now.strftime("%I:%M:%S")
        am_pm = "ص" if now.strftime("%p") == "AM" else "م"
        time_str = f"🕒 {hour} {am_pm}"
        if hasattr(self, "header_date_lbl"):
            self.header_date_lbl.setText(date_str)
        if hasattr(self, "header_time_lbl"):
            self.header_time_lbl.setText(time_str)

    def manual_check_for_updates(self):
        """فحص وجود تحديثات يدوياً عند الضغط على زر فحص التحديثات."""
        if hasattr(self, "show_toast"):
            self.show_toast("🔍 جاري فحص وجود تحديثات جديدة...")
        try:
            from gui.update_dialog import check_updates_manual
            check_updates_manual(self)
        except Exception as e:
            if hasattr(self, "show_toast"):
                self.show_toast(f"⚠️ خطأ: {e}", is_error=True)

    def refresh_dashboard_data(self):
        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            today_date = QDate.currentDate()
            today_str = today_date.toString("yyyy-MM-dd")

            # HR: total employees
            cursor.execute("SELECT COUNT(*) FROM employees")
            total_emp = (cursor.fetchone() or [0])[0] or 0
            if hasattr(self, "kpi_total_emp"):
                self.kpi_total_emp.setText(f"{total_emp} عامل")

            # HR: attendance today
            cursor.execute("SELECT COUNT(*), COALESCE(SUM(work_hours), 0.0) FROM attendance WHERE date=? AND status='حضور'", (today_str,))
            att_row = cursor.fetchone() or [0, 0.0]
            present = att_row[0] or 0
            work_hours = att_row[1] or 0.0

            cursor.execute("SELECT COUNT(*) FROM attendance WHERE date=? AND status='غياب'", (today_str,))
            absent = (cursor.fetchone() or [0])[0] or 0

            if total_emp > 0:
                pct = present / total_emp * 100
                pres_text = f"{pct:.0f}% ({present}/{total_emp})"
            else:
                pres_text = "0 عامل"

            if hasattr(self, "kpi_present_val"):
                self.kpi_present_val.setText(pres_text)
            if hasattr(self, "kpi_absent_val"):
                self.kpi_absent_val.setText(f"{absent} غياب")
            if hasattr(self, "kpi_hours_val"):
                self.kpi_hours_val.setText(f"{work_hours:.1f} ساعة")

            # Monthly financial totals
            cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM incentives WHERE strftime('%Y-%m', date)=strftime('%Y-%m','now')")
            inc_val = (cursor.fetchone() or [0])[0] or 0
            if hasattr(self, "kpi_incentives_val"):
                self.kpi_incentives_val.setText(f"{inc_val:,.0f} ج.م")

            cursor.execute("SELECT COALESCE(SUM(amount), 0) FROM deductions WHERE strftime('%Y-%m', date)=strftime('%Y-%m','now')")
            ded_val = (cursor.fetchone() or [0])[0] or 0
            if hasattr(self, "kpi_deductions_val"):
                self.kpi_deductions_val.setText(f"{ded_val:,.0f} ج.م")

            cursor.execute("SELECT COUNT(*) FROM departments")
            dept_count = (cursor.fetchone() or [0])[0] or 0

            if "شؤون العاملين" in self.cards:
                self.cards["شؤون العاملين"].info_label.setText(
                    f"إجمالي العاملين: {total_emp} عامل | الأقسام: {dept_count} قسم | الحضور اليوم: {present} | الغياب: {absent}"
                )

            conn.close()
        except Exception:
            pass

