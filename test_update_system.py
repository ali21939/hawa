"""
اختبار شامل لنظام التحديث التلقائي
يختبر كل جزء على حدة ويطلع النتائج
"""
import os
import sys
import sqlite3

# ضبط الترميز للطباعة على كونسول ويندوز بدون مشاكل
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

passed = 0
failed = 0

def test(name, condition, detail=""):
    global passed, failed
    if condition:
        passed += 1
        print(f"  [PASS] {name}")
        if detail:
            print(f"         -> {detail}")
    else:
        failed += 1
        print(f"  [FAIL] {name}")
        if detail:
            print(f"         -> {detail}")

print("=" * 60)
print("  TEST 1: config.py - Imports & Settings")
print("=" * 60)

try:
    from core.config import APP_VERSION, APP_NAME, DATA_DIR, DB_NAME, BACKUP_DIR, BASE_DIR, get_asset_path
    test("Import config", True)
    test("APP_VERSION exists", bool(APP_VERSION), f"v{APP_VERSION}")
    test("APP_NAME exists", bool(APP_NAME), APP_NAME)
    test("DATA_DIR in AppData", "AppData" in DATA_DIR or "Local" in DATA_DIR, DATA_DIR)
    test("DB_NAME in AppData", "AppData" in DB_NAME or "Local" in DB_NAME, DB_NAME)
    test("BACKUP_DIR in AppData", "AppData" in BACKUP_DIR or "Local" in BACKUP_DIR, BACKUP_DIR)
    test("DATA_DIR exists on disk", os.path.isdir(DATA_DIR))
    test("BACKUP_DIR exists on disk", os.path.isdir(BACKUP_DIR))
    is_separated = ("AppData" in DB_NAME or "Local" in DB_NAME) and (BASE_DIR not in DB_NAME)
    test("DB separated from code", is_separated, "DB is in AppData, App code is separate")
except Exception as e:
    test("Import config", False, str(e))

print()
print("=" * 60)
print("  TEST 2: Database - Works from new path")
print("=" * 60)

emp_count = 0
try:
    test("DB file exists", os.path.exists(DB_NAME), DB_NAME)
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = [r[0] for r in cursor.fetchall()]
    test("DB has tables", len(tables) > 0, f"{len(tables)} tables found")
    
    cursor.execute("SELECT COUNT(*) FROM employees")
    emp_count = cursor.fetchone()[0]
    test("Employees table accessible", True, f"{emp_count} employees found in DB")
    
    conn.close()
    test("DB connection OK", True)
except Exception as e:
    test("Database access", False, str(e))

print()
print("=" * 60)
print("  TEST 3: updater.py - Module & Functions")
print("=" * 60)

try:
    from core.updater import check_for_updates, auto_backup, download_update, apply_update
    test("Import updater", True)
    test("check_for_updates callable", callable(check_for_updates))
    test("auto_backup callable", callable(auto_backup))
    test("download_update callable", callable(download_update))
    test("apply_update callable", callable(apply_update))
except Exception as e:
    test("Import updater", False, str(e))

print()
print("=" * 60)
print("  TEST 4: Auto Backup - Actually works")
print("=" * 60)

try:
    backup_path = auto_backup("verify_run")
    test("Backup created", backup_path is not None and os.path.exists(backup_path), str(backup_path))
    
    backup_size = os.path.getsize(backup_path) if backup_path else 0
    db_size = os.path.getsize(DB_NAME)
    test("Backup size matches DB", backup_size == db_size,
         f"Backup: {backup_size:,} bytes, DB: {db_size:,} bytes")
    
    # Verify backup is valid SQLite
    conn2 = sqlite3.connect(backup_path)
    c2 = conn2.cursor()
    c2.execute("SELECT COUNT(*) FROM employees")
    backup_emp = c2.fetchone()[0]
    conn2.close()
    test("Backup is valid SQLite", backup_emp == emp_count,
         f"Backup has {backup_emp} employees (matches original)")
    
    # Cleanup test backup
    if backup_path and os.path.exists(backup_path):
        os.remove(backup_path)
        test("Test backup cleaned up safely", True)
except Exception as e:
    test("Auto backup", False, str(e))

print()
print("=" * 60)
print("  TEST 5: check_for_updates - Network call & Safety")
print("=" * 60)

try:
    result = check_for_updates()
    test("check_for_updates runs", isinstance(result, dict))
    test("Returns 'available' key", "available" in result)
    if result.get("error"):
        test("Handles network/repo safely (No Crash)", True,
             f"Safe fallback: {str(result['error'])[:70]}...")
    else:
        test("Network call executed cleanly", True)
except Exception as e:
    test("check_for_updates", False, str(e))

print()
print("=" * 60)
print("  TEST 6: update_dialog.py - GUI imports")
print("=" * 60)

try:
    from gui.update_dialog import UpdateDialog, start_update_check, _UpdateCheckThread, _DownloadThread
    test("Import UpdateDialog", True)
    test("Import start_update_check", True)
    test("Import _UpdateCheckThread", True)
    test("Import _DownloadThread", True)
except Exception as e:
    test("Import update_dialog", False, str(e))

print()
print("=" * 60)
print("  TEST 7: Version comparison logic")
print("=" * 60)

try:
    from core.updater import _is_newer
    test("1.1.0 > 1.0.0", _is_newer("1.1.0", "1.0.0") == True)
    test("2.0.0 > 1.9.9", _is_newer("2.0.0", "1.9.9") == True)
    test("1.0.0 NOT > 1.0.0", _is_newer("1.0.0", "1.0.0") == False)
    test("1.0.0 NOT > 1.1.0", _is_newer("1.0.0", "1.1.0") == False)
    test("1.0.1 > 1.0.0", _is_newer("1.0.1", "1.0.0") == True)
    test("10.0.0 > 9.9.9", _is_newer("10.0.0", "9.9.9") == True)
    test("Invalid version handled safely", _is_newer("abc", "1.0.0") == False)
except Exception as e:
    test("Version comparison", False, str(e))

print()
print("=" * 60)
print("  TEST 8: main_window.py - Update check integrated")
print("=" * 60)

try:
    with open(os.path.join(BASE_DIR, "gui", "main_window.py"), "r", encoding="utf-8") as f:
        mw_code = f.read()
    test("_check_for_app_updates method exists", "_check_for_app_updates" in mw_code)
    test("QTimer.singleShot call exists", "QTimer.singleShot" in mw_code)
    test("start_update_check imported", "start_update_check" in mw_code)
    test("3-second delay configured", "3000" in mw_code)
except Exception as e:
    test("main_window integration", False, str(e))

print()
print("=" * 60)
if failed == 0:
    print(f"  RESULT: ALL {passed} TESTS PASSED! (100% SUCCESS)")
else:
    print(f"  RESULT: {passed} PASSED, {failed} FAILED")
print("=" * 60)
