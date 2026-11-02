"""
=============================================================================
  أداة نشر وتحديث التطبيق أوتوماتيكياً (Hawa Factory Manager Publisher)
=============================================================================
كل اللي عليك: شغل هذا الملف وأنت في البيت:
    python publish.py

الملف بيعمل الآتي تلقائياً:
 1. يسألك عن رقم الإصدار الجديد (ويقترح عليك التالي تلقائياً).
 2. يسألك عن ملاحظات التحديث (إيه الجديد في النسخة).
 3. يعدل رقم الإصدار في core/config.py.
 4. يبني ملف FactoryManager.exe جاهز للتشغيل باستخدام PyInstaller.
 5. يرفع التحديث ويسجل Release على GitHub أوتوماتيكياً عبر API.
 6. يرفع التغييرات إلى Git.

في المصنع:
 أي جهاز يفتح البرنامج أو يضغط زر "🚀 فحص التحديثات" هيتحمل عنده فوراً!
=============================================================================
"""

import os
import sys
import re
import json
import shutil
import subprocess
import urllib.request
import urllib.error
from datetime import datetime

# ضبط ترميز الكونسول للغة العربية والرموز في ويندوز
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "core", "config.py")
TOKEN_FILE = os.path.join(BASE_DIR, ".github_token")
GITHUB_REPO = "ali21939/hawa"


def print_banner():
    print("=" * 65)
    print("      🚀 ناشر تحديثات نظام مصنع حوا (Hawa Factory Manager)")
    print("=" * 65)
    print()


def get_current_version():
    """قراءة رقم الإصدار الحالي من core/config.py."""
    if not os.path.exists(CONFIG_FILE):
        return "1.0.0"
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    match = re.search(r'APP_VERSION\s*=\s*["\']([^"\']+)["\']', content)
    return match.group(1) if match else "1.0.0"


def calculate_next_version(current):
    """اقتراح الإصدار التالي تلقائياً (1.0.0 -> 1.0.1)."""
    try:
        parts = [int(p) for p in current.strip().split(".")]
        parts[-1] += 1
        return ".".join(str(p) for p in parts)
    except Exception:
        return "1.0.1"


def set_version_in_config(new_ver):
    """كتابة رقم الإصدار الجديد في core/config.py."""
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    new_content = re.sub(
        r'APP_VERSION\s*=\s*["\'][^"\']+["\']',
        f'APP_VERSION = "{new_ver}"',
        content
    )
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(new_content)


def get_or_prompt_github_token():
    """الحصول على GitHub Token (يقرأه من الملف أو يسأل عنه ويحفظه)."""
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, "r", encoding="utf-8") as f:
                token = f.read().strip()
                if token:
                    return token
        except Exception:
            pass

    print("\n🔑 مطلوب رمز GitHub Token (لأول مرة فقط لرفع التحديث لـ GitHub):")
    print("   1. افتح الرابط: https://github.com/settings/tokens/new?scopes=repo")
    print("   2. اضغط 'Generate token' وانسخه وضعه هنا.")
    print("   (سيتم حفظه محلياً في .github_token ولن تحتاج كتابته مرة أخرى)")
    token = input("   الصق الـ Token هنا: ").strip()

    if token:
        try:
            with open(TOKEN_FILE, "w", encoding="utf-8") as f:
                f.write(token)
            print("   ✅ تم حفظ الرمز محلياً بنجاح.")
        except Exception as e:
            print(f"   ⚠️ لم يتمكن من حفظ الملف: {e}")
    return token


def build_executable():
    """بناء ملف الـ exe باستخدام PyInstaller."""
    print("\n🔨 [1/4] جاري بناء ملف التطبيق FactoryManager.exe (يرجى الانتظار دقيقة)...")
    
    icon_path = os.path.join(BASE_DIR, "assets", "app_icon.ico")
    assets_dir = os.path.join(BASE_DIR, "assets")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name", "FactoryManager",
        f"--add-data={assets_dir};assets",
        os.path.join(BASE_DIR, "main.py")
    ]
    if os.path.exists(icon_path):
        cmd.insert(cmd.index(os.path.join(BASE_DIR, "main.py")), f"--icon={icon_path}")

    # تشغيل أمر البناء
    process = subprocess.run(cmd, cwd=BASE_DIR, capture_output=True, text=True)
    if process.returncode != 0:
        print("\n❌ حدث خطأ أثناء بناء ملف الـ exe:")
        print(process.stderr[-1000:] if process.stderr else process.stdout[-1000:])
        return None

    exe_path = os.path.join(BASE_DIR, "dist", "FactoryManager.exe")
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print(f"   ✅ تم إنشاء الملف بنجاح! الحجم: {size_mb:.1f} MB")
        return exe_path
    else:
        print("❌ لم يتم العثور على ملف FactoryManager.exe في مجلد dist/")
        return None


def create_github_release(token, version, notes, exe_path):
    """إنشاء Release على GitHub ورفع ملف FactoryManager.exe عليه."""
    print(f"\n☁️ [2/4] جاري إنشاء إصدار Release v{version} على GitHub...")

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "Hawa-FactoryManager-Publisher"
    }

    # 1. إنشاء الـ Release
    release_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases"
    release_payload = {
        "tag_name": f"v{version}",
        "name": f"إصدار v{version} — {datetime.now().strftime('%Y-%m-%d')}",
        "body": notes,
        "draft": False,
        "prerelease": False
    }

    req = urllib.request.Request(
        release_url,
        data=json.dumps(release_payload).encode("utf-8"),
        headers=headers,
        method="POST"
    )

    try:
        with urllib.request.urlopen(req) as resp:
            rel_data = json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode()
        print(f"   ❌ فشل إنشاء Release على GitHub: HTTP {e.code}")
        print(f"   {err_msg}")
        return False
    except Exception as e:
        print(f"   ❌ خطأ في الاتصال: {e}")
        return False

    upload_url_template = rel_data.get("upload_url", "")
    # قالب الرابط ينتهي بـ {?name,label}
    upload_base = upload_url_template.split("{")[0]

    # 2. رفع ملف FactoryManager.exe
    print(f"\n📦 [3/4] جاري رفع ملف FactoryManager.exe إلى GitHub (يرجى الانتظار)...")
    file_size = os.path.getsize(exe_path)
    file_name = "FactoryManager.exe"
    upload_url = f"{upload_base}?name={file_name}"

    upload_headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream",
        "Content-Length": str(file_size),
        "User-Agent": "Hawa-FactoryManager-Publisher"
    }

    try:
        with open(exe_path, "rb") as f:
            exe_bytes = f.read()

        req_upload = urllib.request.Request(
            upload_url,
            data=exe_bytes,
            headers=upload_headers,
            method="POST"
        )
        with urllib.request.urlopen(req_upload) as up_resp:
            print("   ✅ تم رفع الملف إلى GitHub Release بنجاح!")
            return True
    except urllib.error.HTTPError as e:
        print(f"   ❌ فشل رفع الملف: HTTP {e.code}")
        print(f"   {e.read().decode()}")
        return False
    except Exception as e:
        print(f"   ❌ خطأ أثناء الرفع: {e}")
        return False


def git_commit_and_push(version, notes):
    """حفظ التغييرات في Git وإرسالها إلى GitHub."""
    print(f"\n🌿 [4/4] جاري حفظ التحديث في Git...")
    try:
        subprocess.run(["git", "add", "."], cwd=BASE_DIR, check=False)
        subprocess.run(["git", "commit", "-m", f"الإصدار v{version}: {notes}"], cwd=BASE_DIR, check=False)
        subprocess.run(["git", "tag", f"v{version}"], cwd=BASE_DIR, check=False)
        res = subprocess.run(["git", "push", "origin", "main", "--tags"], cwd=BASE_DIR, capture_output=True, text=True)
        if res.returncode == 0:
            print("   ✅ تم رفع التغييرات لفرع main على GitHub.")
        else:
            print("   ℹ️ تم حفظ التغييرات محلياً في Git.")
    except Exception:
        pass


def main():
    print_banner()

    curr_ver = get_current_version()
    next_ver = calculate_next_version(curr_ver)

    print(f"📌 الإصدار الحالي:  v{curr_ver}")
    new_ver_input = input(f"🎯 رقم الإصدار الجديد [اضغط Enter لاختيار v{next_ver}]: ").strip()
    new_ver = new_ver_input.lstrip("v") if new_ver_input else next_ver

    print()
    notes_input = input("📝 ما الجديد في هذا التحديث؟ [اضغط Enter للوضع الافتراضي]: ").strip()
    notes = notes_input if notes_input else "تحديث عام وتحسينات في الأداء والميزات."

    token = get_or_prompt_github_token()
    if not token:
        print("\n❌ لم يتم إدخال الـ Token. تم إلغاء العملية.")
        return

    # تحديث رقم الإصدار في الكود
    set_version_in_config(new_ver)
    print(f"\n✓ تم ضبط رقم الإصدار v{new_ver} في core/config.py")

    # 1. بناء التطبيق
    exe_path = build_executable()
    if not exe_path:
        print("\n❌ توقفت العملية بسبب فشل البناء.")
        return

    # 2. رفع الريليز
    success = create_github_release(token, new_ver, notes, exe_path)
    if not success:
        print("\n❌ توقفت العملية أثناء رفع التحديث على GitHub.")
        return

    # 3. مزامنة Git
    git_commit_and_push(new_ver, notes)

    print("\n" + "=" * 65)
    print(f"  🎉 تم نشر التحديث v{new_ver} بنجاح تام!")
    print("=" * 65)
    print("\nماذا يحدث الآن في أجهزة المصنع؟")
    print("  1. أي جهاز يفتح البرنامج سيظهر له إشعار فوري: 'يوجد تحديث جديد'.")
    print("  2. أو يمكنهم في أي وقت الضغط على زر '🚀 فحص التحديثات' في أعلى الشاشة.")
    print("  3. بضغطة زر واحدة: يتم تحميل الإصدار الجديد، أخذ نسخة احتياطية من الداتا بيز،")
    print("     واستبدال البرنامج وإعادة تشغيله تلقائياً بدون أن تذهب للمصنع إطلاقاً!")
    print("\n" + "=" * 65)


if __name__ == "__main__":
    main()
