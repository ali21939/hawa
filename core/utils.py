import hashlib
from typing import Tuple
from PySide6.QtCore import QTime, QLocale

MASTER_RECOVERY_PASSWORD = "password123"

# Arabic UAE locale provides Western digits (0-9) with Arabic AM/PM indicators (ص / م)
ARABIC_LOCALE = QLocale(QLocale.Language.Arabic, QLocale.Country.UnitedArabEmirates)


def parse_time_str(time_str):
    if not time_str:
        return QTime()
    ts = str(time_str).strip()
    # Normalize Arabic AM/PM to English for robust parsing
    ts_norm = ts.replace("ص", "AM").replace("م", "PM").strip()
    for fmt in ["hh:mm AP", "h:mm AP", "hh:mm", "h:mm", "HH:mm", "H:mm", "hh:mm:ss AP", "HH:mm:ss"]:
        t = QTime.fromString(ts_norm, fmt)
        if t.isValid():
            return t
        t_orig = QTime.fromString(ts, fmt)
        if t_orig.isValid():
            return t_orig
    return QTime()


def format_time_ar(t) -> str:
    """Format QTime or time string into Arabic 12-hour format 'hh:mm ص' / 'hh:mm م'."""
    if not t:
        return ""
    if isinstance(t, str):
        t = parse_time_str(t)
    if not isinstance(t, QTime) or not t.isValid():
        return ""
    hour_12 = t.hour() % 12
    if hour_12 == 0:
        hour_12 = 12
    suffix = "ص" if t.hour() < 12 else "م"
    return f"{hour_12:02d}:{t.minute():02d} {suffix}"




def hash_password(password: str) -> str:
    """تشفير كلمة المرور باستخدام SHA-256."""
    if not password:
        return ""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verify_app_password(entered_password: str, stored_hash: str) -> Tuple[bool, bool]:
    """
    التحقق من كلمة المرور المدخلة.
    يرجع (is_valid, is_master_recovery).
    """
    if entered_password == MASTER_RECOVERY_PASSWORD:
        return True, True
    if not stored_hash:
        return False, False
    return (hash_password(entered_password) == stored_hash), False

