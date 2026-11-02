import sys
import os
import ctypes
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from gui.main_window import MainWindow
from core.database import init_database, is_app_password_enabled
from gui.widgets import AppLockDialog
from core.config import BASE_DIR, get_asset_path

def main():
    try:
        myappid = 'hawa.factorymanager.v2'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    icon_file = get_asset_path("app_icon.ico")
    if not os.path.exists(icon_file):
        icon_file = get_asset_path("app_icon.png")
    if os.path.exists(icon_file):
        app.setWindowIcon(QIcon(icon_file))

    init_database()

    # Check App Password Protection
    is_recovered = False
    if is_app_password_enabled():
        login_dlg = AppLockDialog(is_startup=True)
        if login_dlg.exec() != AppLockDialog.Accepted:
            sys.exit(0)
        is_recovered = getattr(login_dlg, "is_master_recovery", False)

    app.setStyleSheet("""
        QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus, QDateEdit:focus, QTimeEdit:focus {
            border: 2px solid #2563eb;
        }
        
        QTableWidget QComboBox, 
        QTableWidget QTimeEdit, 
        QTableWidget QSpinBox, 
        QTableWidget QDoubleSpinBox {
            min-width: 120px;
            min-height: 30px;
        }
    """)

    window = MainWindow()
    if is_recovered:
        window.notify_master_recovery_used()

    window.showMaximized()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

