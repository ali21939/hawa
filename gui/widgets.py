from PySide6.QtCore import *
from PySide6.QtWidgets import *
from PySide6.QtGui import *
import os, sys
from datetime import datetime
from core.config import DB_NAME, _BASE_DIR, BASE_DIR


class DashboardCard(QFrame):
    def __init__(self, title, bg_color, click_callback, icon_str="📊", category="قسم تشغيلي", tint_color=None, border_color=None, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.StyledPanel)
        self.setCursor(Qt.PointingHandCursor)
        self.click_callback = click_callback
        self.primary_color = bg_color
        self.tint_color = tint_color or "#f8fafc"
        self.border_color = border_color or "#e2e8f0"
        self.category = category
        self.icon_str = icon_str
        
        self.shadow = QGraphicsDropShadowEffect(self)
        self.shadow.setBlurRadius(14)
        self.shadow.setXOffset(0)
        self.shadow.setYOffset(3)
        self.shadow.setColor(QColor(15, 23, 42, 18))
        self.setGraphicsEffect(self.shadow)
        
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.setMinimumHeight(148)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        
        top_row = QHBoxLayout()
        top_row.setSpacing(12)
        top_row.setContentsMargins(0, 0, 0, 0)
        
        self.icon_box = QFrame()
        self.icon_box.setFixedSize(50, 50)
        self.icon_box.setStyleSheet(f"""
            QFrame {{
                background-color: {self.tint_color};
                border: 1.5px solid {self.border_color};
                border-radius: 12px;
            }}
        """)
        icon_box_layout = QVBoxLayout(self.icon_box)
        icon_box_layout.setContentsMargins(0, 0, 0, 0)
        icon_box_layout.setAlignment(Qt.AlignCenter)
        
        icon_label = QLabel(icon_str)
        icon_label.setFont(QFont("Segoe UI", 20))
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet("background: transparent; border: none;")
        icon_box_layout.addWidget(icon_label)
        
        text_vbox = QVBoxLayout()
        text_vbox.setSpacing(2)
        text_vbox.setContentsMargins(0, 0, 0, 0)
        
        cat_label = QLabel(self.category)
        cat_label.setFont(QFont("Segoe UI", 9, QFont.Bold))
        cat_label.setStyleSheet(f"color: {self.primary_color}; background: transparent; border: none;")
        
        title_label = QLabel(title)
        title_label.setFont(QFont("Segoe UI", 15, QFont.Bold))
        title_label.setStyleSheet("color: #0f172a; background: transparent; border: none;")
        
        text_vbox.addWidget(cat_label)
        text_vbox.addWidget(title_label)
        
        arrow_label = QLabel("←")
        arrow_label.setFont(QFont("Segoe UI", 16, QFont.Bold))
        arrow_label.setStyleSheet("color: #94a3b8; background: transparent; border: none;")
        
        top_row.addWidget(self.icon_box)
        top_row.addLayout(text_vbox)
        top_row.addStretch()
        top_row.addWidget(arrow_label)
        
        layout.addLayout(top_row)
        
        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setStyleSheet("background-color: #f1f5f9; max-height: 1px; border: none;")
        layout.addWidget(divider)
        
        self.info_label = QLabel("انقر لفتح القسم وإدارته...")
        self.info_label.setFont(QFont("Segoe UI", 10))
        self.info_label.setStyleSheet("color: #64748b; background: transparent; border: none;")
        layout.addWidget(self.info_label)
        
        self.setStyleSheet(f"""
            QFrame {{
                background-color: #ffffff;
                border: 1.5px solid {self.border_color};
                border-radius: 14px;
            }}
            QFrame:hover {{
                border: 2px solid {self.primary_color};
                background-color: #ffffff;
            }}
        """)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.click_callback()
        super().mousePressEvent(event)


class BaseSectionWidget(QWidget):
    """
    عنصر الحاوية الموحد لشاشات الأقسام لتوفير هيدر فاخر وموحد وأزرار عودة منسقة
    """
    def __init__(self, title, back_callback, icon_str="📦", primary_color="#2563eb", tint_color="#eff6ff", border_color="#bfdbfe", category="قسم رئيسي", parent=None):
        super().__init__(parent)
        self.title = title
        self.primary_color = primary_color
        self.tint_color = tint_color
        self.border_color = border_color
        
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(18, 14, 18, 14)
        self.layout.setSpacing(12)
        
        self.header_card = QFrame()
        self.header_card.setStyleSheet("""
            QFrame {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
        """)
        header_shadow = QGraphicsDropShadowEffect(self.header_card)
        header_shadow.setBlurRadius(10)
        header_shadow.setYOffset(2)
        header_shadow.setColor(QColor(15, 23, 42, 10))
        self.header_card.setGraphicsEffect(header_shadow)
        
        header_layout = QHBoxLayout(self.header_card)
        header_layout.setContentsMargins(16, 10, 16, 10)
        header_layout.setSpacing(14)
        
        self.icon_box = QFrame()
        self.icon_box.setFixedSize(46, 46)
        self.icon_box.setStyleSheet(f"""
            QFrame {{
                background-color: {self.tint_color};
                border: 1.5px solid {self.border_color};
                border-radius: 12px;
            }}
        """)
        icon_box_layout = QVBoxLayout(self.icon_box)
        icon_box_layout.setContentsMargins(0, 0, 0, 0)
        icon_box_layout.setAlignment(Qt.AlignCenter)
        
        self.icon_lbl = QLabel(icon_str)
        self.icon_lbl.setFont(QFont("Segoe UI", 18))
        self.icon_lbl.setAlignment(Qt.AlignCenter)
        self.icon_lbl.setStyleSheet("background: transparent; border: none;")
        icon_box_layout.addWidget(self.icon_lbl)
        
        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(2)
        
        self.title_label = QLabel(title)
        self.title_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        self.title_label.setStyleSheet("color: #0f172a; background: transparent; border: none;")
        
        self.cat_label = QLabel(category)
        self.cat_label.setFont(QFont("Segoe UI", 9))
        self.cat_label.setStyleSheet("color: #64748b; background: transparent; border: none;")
        
        title_vbox.addWidget(self.title_label)
        title_vbox.addWidget(self.cat_label)
        
        header_layout.addWidget(self.icon_box)
        header_layout.addLayout(title_vbox)
        header_layout.addStretch()
        
        back_btn = QPushButton("🏠 العودة للرئيسية")
        back_btn.setFont(QFont("Segoe UI", 9, QFont.Bold))
        back_btn.setCursor(Qt.PointingHandCursor)
        back_btn.setStyleSheet(f"""
            QPushButton {{ 
                background-color: #0f172a; 
                color: #ffffff; 
                border-radius: 8px; 
                padding: 8px 16px; 
                border: none;
            }}
            QPushButton:hover {{ 
                background-color: {self.primary_color}; 
            }}
            QPushButton:pressed {{ 
                background-color: #0f172a; 
            }}
        """)
        back_btn.clicked.connect(back_callback)
        header_layout.addWidget(back_btn)
        
        self.layout.addWidget(self.header_card)


class AppLockDialog(QDialog):
    """
    نافذة قفل وتسجيل الدخول للتطبيق مع دعم تلميح كلمة المرور وكلمة مرور الاسترداد الاحتياطية (password123).
    """
    def __init__(self, parent=None, is_startup=True):
        super().__init__(parent)
        self.setWindowTitle("🔒 تسجيل الدخول — مصنع حوا")
        self.setFixedWidth(460)
        self.setLayoutDirection(Qt.RightToLeft)
        self.is_startup = is_startup
        self.is_master_recovery = False
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        self.setStyleSheet("background-color: #f8fafc; font-family: 'Segoe UI', Tahoma, sans-serif;")

        self.setup_ui()

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(14)

        header_card = QFrame()
        header_card.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #1e293b, stop:1 #020617);
                border-radius: 16px;
                border: 1px solid #334155;
                padding: 24px;
            }
        """)
        h_layout = QVBoxLayout(header_card)
        h_layout.setAlignment(Qt.AlignCenter)
        h_layout.setSpacing(6)

        icon_lbl = QLabel("🔒")
        icon_lbl.setFont(QFont("Segoe UI", 32))
        icon_lbl.setAlignment(Qt.AlignCenter)
        icon_lbl.setStyleSheet("border:none; background:transparent;")

        title_lbl = QLabel("مصنع حوا — شؤون العاملين")
        title_lbl.setFont(QFont("Segoe UI", 16, QFont.Bold))
        title_lbl.setStyleSheet("color: white; border:none; background:transparent;")

        subtitle_lbl = QLabel("يرجى إدخال كلمة المرور للمتابعة إلى النظام")
        subtitle_lbl.setFont(QFont("Segoe UI", 10))
        subtitle_lbl.setStyleSheet("color: #94a3b8; border:none; background:transparent;")

        h_layout.addWidget(icon_lbl)
        h_layout.addWidget(title_lbl)
        h_layout.addWidget(subtitle_lbl)
        main_layout.addWidget(header_card)

        form_card = QFrame()
        form_card.setStyleSheet("""
            QFrame {
                background: white;
                border-radius: 16px;
                border: 1px solid #e2e8f0;
                padding: 20px;
            }
        """)
        f_layout = QVBoxLayout(form_card)
        f_layout.setSpacing(14)

        pass_lbl = QLabel("كلمة المرور:")
        pass_lbl.setFont(QFont("Segoe UI", 11, QFont.Bold))
        pass_lbl.setStyleSheet("color: #1e293b; border:none;")
        f_layout.addWidget(pass_lbl)

        input_container = QFrame()
        input_container.setStyleSheet("""
            QFrame {
                border: 2px solid #cbd5e1;
                border-radius: 10px;
                background: #f8fafc;
            }
            QFrame:focus-within {
                border-color: #2563eb;
                background: white;
            }
        """)
        ic_layout = QHBoxLayout(input_container)
        ic_layout.setContentsMargins(10, 2, 10, 2)
        ic_layout.setSpacing(6)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("أدخل كلمة المرور...")
        self.password_input.setFont(QFont("Segoe UI", 12))
        self.password_input.setStyleSheet("border: none; background: transparent; color: #0f172a; padding: 6px 0;")
        self.password_input.returnPressed.connect(self.check_password)

        self.toggle_eye_btn = QPushButton("👁️")
        self.toggle_eye_btn.setCursor(Qt.PointingHandCursor)
        self.toggle_eye_btn.setFixedSize(32, 32)
        self.toggle_eye_btn.setStyleSheet("""
            QPushButton {
                border: none;
                background: transparent;
                font-size: 14px;
            }
            QPushButton:hover { background: #e2e8f0; border-radius: 6px; }
        """)
        self.toggle_eye_btn.clicked.connect(self.toggle_password_visibility)

        ic_layout.addWidget(self.password_input)
        ic_layout.addWidget(self.toggle_eye_btn)
        f_layout.addWidget(input_container)

        self.status_lbl = QLabel("")
        self.status_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.status_lbl.setStyleSheet("color: #dc2626; border:none; background:transparent;")
        self.status_lbl.setWordWrap(True)
        self.status_lbl.hide()
        f_layout.addWidget(self.status_lbl)

        self.unlock_btn = QPushButton("🔓 تسجيل الدخول")
        self.unlock_btn.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.unlock_btn.setCursor(Qt.PointingHandCursor)
        self.unlock_btn.setMinimumHeight(44)
        self.unlock_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #1d4ed8);
                color: white;
                border-radius: 10px;
                border: none;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1d4ed8, stop:1 #1e40af);
            }
            QPushButton:pressed {
                background: #1e3a8a;
            }
        """)
        self.unlock_btn.clicked.connect(self.check_password)
        f_layout.addWidget(self.unlock_btn)

        self.forgot_btn = QPushButton("نسيت كلمة المرور؟")
        self.forgot_btn.setCursor(Qt.PointingHandCursor)
        self.forgot_btn.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.forgot_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #2563eb;
                border: none;
                text-decoration: underline;
                padding: 6px;
            }
            QPushButton:hover { color: #1d4ed8; }
        """)
        self.forgot_btn.clicked.connect(self.toggle_hint_box)
        f_layout.addWidget(self.forgot_btn, alignment=Qt.AlignCenter)

        self.hint_box = QFrame()
        self.hint_box.setStyleSheet("""
            QFrame {
                background: #f0fdfa;
                border: 1px solid #ccfbf1;
                border-radius: 12px;
                padding: 14px;
            }
        """)
        h_vbox = QVBoxLayout(self.hint_box)
        h_vbox.setSpacing(8)

        self.hint_text_lbl = QLabel("")
        self.hint_text_lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.hint_text_lbl.setStyleSheet("color: #0f766e; border:none; background:transparent;")
        self.hint_text_lbl.setWordWrap(True)

        recovery_lbl = QLabel("🔑 للاسترداد الاحتياطي استخدم: <code style='color:#0f766e; background:#ccfbf1; padding:2px 6px; border-radius:4px;'>password123</code>")
        recovery_lbl.setTextFormat(Qt.RichText)
        recovery_lbl.setFont(QFont("Segoe UI", 9))
        recovery_lbl.setStyleSheet("color: #0d9488; border:none; background:transparent;")

        h_vbox.addWidget(self.hint_text_lbl)
        h_vbox.addWidget(recovery_lbl)
        self.hint_box.hide()
        f_layout.addWidget(self.hint_box)

        main_layout.addWidget(form_card)
        self.password_input.setFocus()

    def toggle_password_visibility(self):
        if self.password_input.echoMode() == QLineEdit.Password:
            self.password_input.setEchoMode(QLineEdit.Normal)
            self.toggle_eye_btn.setText("🔒")
        else:
            self.password_input.setEchoMode(QLineEdit.Password)
            self.toggle_eye_btn.setText("👁️")

    def toggle_hint_box(self):
        if self.hint_box.isVisible():
            self.hint_box.hide()
        else:
            from core.database import get_app_password_hint
            hint = get_app_password_hint()
            if hint and hint.strip():
                self.hint_text_lbl.setText(f"💡 تلميح كلمة المرور: <b>{hint.strip()}</b>")
            else:
                self.hint_text_lbl.setText("💡 تلميح: لم يتم تسجيل تلميح مسبق. يمكنك استخدام كلمة مرور الاسترداد.")
            self.hint_box.show()
            self.adjustSize()

    def check_password(self):
        entered = self.password_input.text()
        from core.database import get_app_password_hash
        from core.utils import verify_app_password

        stored_hash = get_app_password_hash()
        is_valid, is_master = verify_app_password(entered, stored_hash)

        if is_valid:
            self.is_master_recovery = is_master
            self.status_lbl.hide()
            self.accept()
        else:
            self.status_lbl.setText("❌ كلمة المرور غير صحيحة! يرجى المحاولة مجدداً.")
            self.status_lbl.show()
            self.password_input.selectAll()
            self.password_input.setFocus()
