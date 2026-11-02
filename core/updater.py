"""
نظام التحديث التلقائي عبر GitHub Releases.

الطريقة:
1. التطبيق يفحص GitHub عند الفتح — هل في إصدار جديد؟
2. لو في تحديث → يسأل المستخدم
3. يحمل النسخة الجديدة ويستبدل القديمة تلقائي
4. يعمل إعادة تشغيل

أنت بس بترفع Release جديد على GitHub وأنت في البيت، والتطبيق في الشركة يلاقيه لوحده.
"""
import json
import os
import sys
import shutil
import subprocess
import tempfile
import urllib.request
from datetime import datetime

from core.config import APP_VERSION, DATA_DIR, BACKUP_DIR, DB_NAME

# ─── مستودع التطبيق على GitHub ───────────────────────────────────────────
GITHUB_REPO = "ali21939/hawa"
# ──────────────────────────────────────────────────────────────────────────

RELEASES_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


# ╔══════════════════════════════════════════════════════════════╗
# ║                   فحص التحديثات                              ║
# ╚══════════════════════════════════════════════════════════════╝

def check_for_updates():
    """
    فحص وجود إصدار جديد على GitHub.
    يرجع dict فيه معلومات التحديث أو إنه مفيش.
    """
    try:
        req = urllib.request.Request(RELEASES_URL)
        req.add_header("User-Agent", "FactoryManager-Updater")
        with urllib.request.urlopen(req, timeout=15) as response:
            data = json.loads(response.read().decode())

        latest_version = data.get("tag_name", "").lstrip("v")
        if not latest_version:
            return {"available": False}

        if _is_newer(latest_version, APP_VERSION):
            download_url, file_size = _get_asset_info(data)
            return {
                "available": True,
                "version": latest_version,
                "notes": data.get("body", ""),
                "download_url": download_url,
                "size": file_size,
            }
        return {"available": False}
    except Exception as e:
        return {"available": False, "error": str(e)}


def _is_newer(latest, current):
    """مقارنة إصدارين: 1.2.0 > 1.1.0"""
    def parse(v):
        return tuple(int(x) for x in v.strip().split("."))
    try:
        return parse(latest) > parse(current)
    except Exception:
        return False


def _get_asset_info(release_data):
    """الحصول على رابط وحجم ملف التحديث (.exe أو .zip)."""
    for asset in release_data.get("assets", []):
        name = asset.get("name", "")
        if name.endswith(".exe") or name.endswith(".zip"):
            return asset["browser_download_url"], asset.get("size", 0)
    return None, 0


# ╔══════════════════════════════════════════════════════════════╗
# ║                  النسخ الاحتياطي التلقائي                     ║
# ╚══════════════════════════════════════════════════════════════╝

def auto_backup(label="auto"):
    """
    نسخة احتياطية من قاعدة البيانات قبل أي تحديث.
    يحتفظ بآخر 10 نسخ فقط.
    """
    if not os.path.exists(DB_NAME):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"backup_{label}_{timestamp}.db"
    backup_path = os.path.join(BACKUP_DIR, backup_name)
    shutil.copy2(DB_NAME, backup_path)
    _cleanup_old_backups(max_keep=10)
    return backup_path


def _cleanup_old_backups(max_keep=10):
    """حذف النسخ الاحتياطية القديمة (يحتفظ بآخر 10)."""
    if not os.path.isdir(BACKUP_DIR):
        return
    backups = sorted(
        [f for f in os.listdir(BACKUP_DIR) if f.startswith("backup_")],
        reverse=True,
    )
    for old in backups[max_keep:]:
        try:
            os.remove(os.path.join(BACKUP_DIR, old))
        except OSError:
            pass


# ╔══════════════════════════════════════════════════════════════╗
# ║                  تحميل وتطبيق التحديث                       ║
# ╚══════════════════════════════════════════════════════════════╝

def download_update(download_url, progress_callback=None):
    """
    تحميل ملف التحديث من GitHub.
    progress_callback(percent: int) → يتنده كل شوية بالنسبة المئوية.
    يرجع مسار الملف المحمل.
    """
    if not download_url:
        raise ValueError("لا يوجد رابط تحميل")

    update_dir = os.path.join(DATA_DIR, "_updates")
    os.makedirs(update_dir, exist_ok=True)

    filename = download_url.split("/")[-1]
    download_path = os.path.join(update_dir, filename)

    # حذف نسخة قديمة لو موجودة
    if os.path.exists(download_path):
        os.remove(download_path)

    def _hook(block_num, block_size, total_size):
        if progress_callback and total_size > 0:
            percent = min(100, int(block_num * block_size * 100 / total_size))
            progress_callback(percent)

    urllib.request.urlretrieve(download_url, download_path, _hook)
    return download_path


def apply_update(downloaded_path):
    """
    تطبيق التحديث:
    - لو .exe → batch script يستبدل الملف بعد إغلاق التطبيق
    - لو .zip → يفك الضغط ويستبدل الملفات

    التطبيق بيقفل وبيفتح تاني لوحده.
    """
    # نسخة احتياطية أولاً
    auto_backup("pre_update")

    if downloaded_path.endswith(".exe"):
        _apply_exe_update(downloaded_path)
    elif downloaded_path.endswith(".zip"):
        _apply_zip_update(downloaded_path)
    else:
        raise ValueError(f"نوع ملف غير مدعوم: {downloaded_path}")


def _apply_exe_update(new_exe_path):
    """
    استبدال ملف exe الحالي بالجديد.
    بيستخدم batch script عشان مش ممكن تستبدل ملف وهو شغال.
    """
    current_exe = sys.executable
    bat_path = os.path.join(DATA_DIR, "_updates", "_do_update.bat")

    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(f'''@echo off
chcp 65001 >nul
echo.
echo ══════════════════════════════════════
echo    جاري تحديث التطبيق... انتظر
echo ══════════════════════════════════════
echo.
timeout /t 3 /nobreak >nul
copy /y "{new_exe_path}" "{current_exe}"
if errorlevel 1 (
    echo ❌ فشل التحديث!
    pause
    exit /b 1
)
del "{new_exe_path}" >nul 2>&1
echo ✅ تم التحديث بنجاح!
timeout /t 1 /nobreak >nul
start "" "{current_exe}"
del "%~f0" >nul 2>&1
''')

    # تشغيل السكريبت وإغلاق التطبيق
    subprocess.Popen(
        ["cmd", "/c", bat_path],
        creationflags=0x00000010,  # CREATE_NEW_CONSOLE
    )
    sys.exit(0)


def _apply_zip_update(zip_path):
    """
    فك ضغط التحديث واستبدال ملفات التطبيق.
    مناسب لو بتوزع التطبيق كمجلد مش exe واحد.
    """
    import zipfile
    current_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    bat_path = os.path.join(DATA_DIR, "_updates", "_do_update.bat")
    extract_dir = os.path.join(DATA_DIR, "_updates", "_extracted")

    # فك الضغط
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(extract_dir)

    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(f'''@echo off
chcp 65001 >nul
echo.
echo ══════════════════════════════════════
echo    جاري تحديث التطبيق... انتظر
echo ══════════════════════════════════════
echo.
timeout /t 3 /nobreak >nul
xcopy /s /y /q "{extract_dir}\\*" "{current_dir}\\"
if errorlevel 1 (
    echo ❌ فشل التحديث!
    pause
    exit /b 1
)
rmdir /s /q "{extract_dir}" >nul 2>&1
del "{zip_path}" >nul 2>&1
echo ✅ تم التحديث بنجاح!
timeout /t 1 /nobreak >nul
start "" "{os.path.join(current_dir, os.path.basename(sys.executable))}"
del "%~f0" >nul 2>&1
''')

    subprocess.Popen(
        ["cmd", "/c", bat_path],
        creationflags=0x00000010,
    )
    sys.exit(0)
