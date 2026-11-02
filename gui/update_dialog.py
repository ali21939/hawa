"""
واجهة التحديث التلقائي — نافذة أنيقة تظهر للمستخدم لما يكون في تحديث جديد.
"""
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QProgressBar, QMessageBox,
)
from PySide6.QtGui import QFont


class _UpdateCheckThread(QThread):
    """Thread يفحص التحديثات في الخلفية بدون تعليق الواجهة."""
    result_ready = Signal(dict)

    def run(self):
        from core.updater import check_for_updates
        info = check_for_updates()
        self.result_ready.emit(info)


class _DownloadThread(QThread):
    """Thread يحمل التحديث ويبلّغ بالنسبة المئوية."""
    progress = Signal(int)
    done = Signal(bool, str)  # success, path_or_error

    def __init__(self, url, parent=None):
        super().__init__(parent)
        self._url = url

    def run(self):
        from core.updater import download_update
        try:
            path = download_update(self._url, progress_callback=self.progress.emit)
            self.done.emit(True, path)
        except Exception as e:
            self.done.emit(False, str(e))


class UpdateDialog(QDialog):
    """نافذة تحديث التطبيق."""

    def __init__(self, update_info: dict, parent=None):
        super().__init__(parent)
        self._info = update_info
        self._downloaded_path = None
        self.setWindowTitle("تحديث متاح")
        self.setFixedSize(500, 320)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self._build_ui()

    # ── بناء الواجهة ──────────────────────────────────────────
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(24, 20, 24, 20)

        # العنوان
        title = QLabel(f"🎉 إصدار جديد متاح:  v{self._info['version']}")
        title.setFont(QFont("Segoe UI", 15, QFont.Bold))
        title.setStyleSheet("color: #2563eb;")
        layout.addWidget(title)

        # الإصدار الحالي
        from core.config import APP_VERSION
        current_lbl = QLabel(f"الإصدار الحالي:  v{APP_VERSION}")
        current_lbl.setStyleSheet("color: #64748b; font-size: 12px;")
        layout.addWidget(current_lbl)

        # ملاحظات التحديث
        notes = self._info.get("notes", "").strip()
        if notes:
            notes_lbl = QLabel(notes[:500])
            notes_lbl.setWordWrap(True)
            notes_lbl.setStyleSheet("""
                background: #f8fafc; border: 1px solid #e2e8f0;
                border-radius: 8px; padding: 12px;
                color: #334155; font-size: 12px;
            """)
            layout.addWidget(notes_lbl)

        # حجم التحميل
        size_mb = self._info.get("size", 0) / (1024 * 1024)
        if size_mb > 0:
            size_lbl = QLabel(f"📦 حجم التحميل: {size_mb:.1f} MB")
            size_lbl.setStyleSheet("color: #64748b; font-size: 11px;")
            layout.addWidget(size_lbl)

        layout.addStretch()

        # شريط التقدم (مخفي في البداية)
        self._progress = QProgressBar()
        self._progress.setVisible(False)
        self._progress.setTextVisible(True)
        self._progress.setStyleSheet("""
            QProgressBar {
                border: 1px solid #cbd5e1; border-radius: 6px;
                background: #f1f5f9; height: 22px; text-align: center;
                font-size: 11px; color: #334155;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #3b82f6, stop:1 #2563eb);
                border-radius: 5px;
            }
        """)
        layout.addWidget(self._progress)

        # رسالة الحالة
        self._status = QLabel("")
        self._status.setAlignment(Qt.AlignCenter)
        self._status.setStyleSheet("font-size: 12px; color: #475569;")
        layout.addWidget(self._status)

        # أزرار
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)

        self._skip_btn = QPushButton("لاحقاً")
        self._skip_btn.setStyleSheet("""
            QPushButton {
                padding: 10px 28px; border-radius: 8px;
                border: 1.5px solid #cbd5e1; background: #ffffff;
                color: #475569; font-size: 13px; font-weight: bold;
            }
            QPushButton:hover { background: #f1f5f9; }
        """)
        self._skip_btn.clicked.connect(self.reject)

        self._update_btn = QPushButton("⬆️  تحديث الآن")
        self._update_btn.setStyleSheet("""
            QPushButton {
                padding: 10px 28px; border-radius: 8px;
                background: #2563eb; color: #ffffff;
                font-size: 13px; font-weight: bold; border: none;
            }
            QPushButton:hover { background: #1d4ed8; }
            QPushButton:disabled { background: #94a3b8; }
        """)
        self._update_btn.clicked.connect(self._start_download)

        btn_row.addWidget(self._skip_btn)
        btn_row.addStretch()
        btn_row.addWidget(self._update_btn)
        layout.addLayout(btn_row)

    # ── تحميل التحديث ─────────────────────────────────────────
    def _start_download(self):
        url = self._info.get("download_url")
        if not url:
            QMessageBox.warning(self, "خطأ", "لم يتم العثور على ملف التحديث!")
            return

        self._update_btn.setEnabled(False)
        self._skip_btn.setEnabled(False)
        self._progress.setVisible(True)
        self._progress.setValue(0)
        self._status.setText("⏳ جاري التحميل...")

        self._dl_thread = _DownloadThread(url, self)
        self._dl_thread.progress.connect(self._on_progress)
        self._dl_thread.done.connect(self._on_download_done)
        self._dl_thread.start()

    def _on_progress(self, percent):
        self._progress.setValue(percent)

    def _on_download_done(self, success, path_or_error):
        if success:
            self._downloaded_path = path_or_error
            self._status.setText("✅ اكتمل التحميل! جاري التحديث...")
            self._progress.setValue(100)
            # تطبيق التحديث (سيغلق التطبيق ويعيد تشغيله)
            from core.updater import apply_update
            try:
                apply_update(self._downloaded_path)
            except Exception as e:
                self._status.setText(f"❌ فشل التطبيق: {e}")
                self._update_btn.setEnabled(True)
                self._skip_btn.setEnabled(True)
        else:
            self._status.setText(f"❌ فشل التحميل: {path_or_error}")
            self._update_btn.setEnabled(True)
            self._skip_btn.setEnabled(True)


# ╔══════════════════════════════════════════════════════════════╗
# ║   دوال مساعدة — استخدمها في MainWindow                      ║
# ╚══════════════════════════════════════════════════════════════╝

def start_update_check(parent_window):
    """
    تشغيل فحص التحديثات في الخلفية.
    استدعيها في MainWindow.__init__() أو بعد ما الواجهة تظهر.

    مثال:
        from gui.update_dialog import start_update_check
        start_update_check(self)
    """
    checker = _UpdateCheckThread(parent_window)
    checker.result_ready.connect(lambda info: _handle_check_result(info, parent_window))
    checker.start()
    # نحتفظ بمرجع عشان الـ garbage collector ما يمسحوش
    parent_window._update_checker_thread = checker


def _handle_check_result(info, parent_window):
    if info.get("available"):
        dlg = UpdateDialog(info, parent_window)
        dlg.exec()


def check_updates_manual(parent_window):
    """
    فحص التحديثات يدويًا بناءً على طلب المستخدم (زر فحص التحديثات).
    يعرض إشعار بالنتيجة سواء في تحديث أو لا.
    """
    checker = _UpdateCheckThread(parent_window)
    def _on_manual_result(info):
        from core.config import APP_VERSION
        if info.get("available"):
            dlg = UpdateDialog(info, parent_window)
            dlg.exec()
        elif info.get("error"):
            if hasattr(parent_window, "show_toast"):
                parent_window.show_toast("⚠️ تعذر الاتصال بالسيرفر لفحص التحديثات", is_error=True)
            else:
                QMessageBox.warning(parent_window, "فحص التحديثات", "تعذر الاتصال بالسيرفر لفحص التحديثات.")
        else:
            if hasattr(parent_window, "show_toast"):
                parent_window.show_toast(f"✓ أنت تستخدم أحدث إصدار بالفعل (v{APP_VERSION})")
            else:
                QMessageBox.information(parent_window, "فحص التحديثات", f"أنت تستخدم أحدث إصدار بالفعل (v{APP_VERSION})")

    checker.result_ready.connect(_on_manual_result)
    checker.start()
    parent_window._manual_update_checker_thread = checker

