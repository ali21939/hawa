import os as _os
import sys

# ── معلومات التطبيق ──────────────────────────────────────────
APP_NAME = "FactoryManager"
APP_VERSION = "1.0.0"  # ← زوّد الرقم ده مع كل تحديث

# ── مسار التطبيق (الكود / exe) ────────────────────────────────
if getattr(sys, 'frozen', False):
    _BASE_DIR = _os.path.dirname(sys.executable)
else:
    _BASE_DIR = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))

BASE_DIR = _BASE_DIR

# ── مسار البيانات الدائم (لا يتأثر بالتحديث أبداً) ──────────
# الداتابيز والنسخ الاحتياطية هتكون في AppData
# حتى لو مسحت التطبيق ونزلت نسخة جديدة، البيانات آمنة
if sys.platform == "win32":
    DATA_DIR = _os.path.join(
        _os.environ.get("LOCALAPPDATA", _os.path.expanduser("~")),
        APP_NAME
    )
else:
    DATA_DIR = _os.path.join(_os.path.expanduser("~"), f".{APP_NAME}")

_os.makedirs(DATA_DIR, exist_ok=True)

BACKUP_DIR = _os.path.join(DATA_DIR, "backups")
_os.makedirs(BACKUP_DIR, exist_ok=True)

# ── قاعدة البيانات ───────────────────────────────────────────
# أولاً: لو الداتابيز في المكان الجديد (AppData) → استخدمها
# ثانياً: لو لسه في المكان القديم (assets/) → انقلها تلقائي للمكان الجديد
# ثالثاً: لو أول مرة → أنشئها في المكان الجديد

_new_db_path = _os.path.join(DATA_DIR, "factory_production_v2.db")
_old_db_assets = _os.path.join(_BASE_DIR, "assets", "factory_production_v2.db")
_old_db_root = _os.path.join(_BASE_DIR, "factory_production_v2.db")

if _os.path.exists(_new_db_path):
    # الداتابيز موجودة في المكان الجديد ✅
    DB_NAME = _new_db_path
elif _os.path.exists(_old_db_assets):
    # نقل تلقائي من المكان القديم للجديد
    import shutil as _shutil
    _shutil.copy2(_old_db_assets, _new_db_path)
    DB_NAME = _new_db_path
    print(f"[OK] Database migrated to: {_new_db_path}")
elif _os.path.exists(_old_db_root):
    import shutil as _shutil
    _shutil.copy2(_old_db_root, _new_db_path)
    DB_NAME = _new_db_path
    print(f"[OK] Database migrated to: {_new_db_path}")
else:
    # أول مرة — هتتنشأ في المكان الجديد
    DB_NAME = _new_db_path


FRIDAY_DAY = 5


def get_asset_path(filename):
    """Find asset file either in PyInstaller MEIPASS, local assets/, or root."""
    if hasattr(sys, '_MEIPASS'):
        p_mei = _os.path.join(sys._MEIPASS, 'assets', filename)
        if _os.path.exists(p_mei):
            return p_mei
        p_mei_root = _os.path.join(sys._MEIPASS, filename)
        if _os.path.exists(p_mei_root):
            return p_mei_root
    p_assets = _os.path.join(_BASE_DIR, 'assets', filename)
    if _os.path.exists(p_assets):
        return p_assets
    p_root = _os.path.join(_BASE_DIR, filename)
    if _os.path.exists(p_root):
        return p_root
    return p_assets
