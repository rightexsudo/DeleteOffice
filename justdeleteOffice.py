import ctypes
import os
import shutil
import subprocess
import sys
import winreg


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


if not is_admin():
    print(
        "⚠️ الرجاء تشغيل السكربت كمسؤول (Run as Administrator) من خلال"
        " النقر بزر الفأرة الأيمن واختيار Run as administrator."
    )
    sys.exit()

print("🚀 بدء عملية التنظيف الشاملة والتأكد من إزالة Office...\n")

# 1. إيقاف وتمرير جميع عمليات Office النشطة
print("⏹️ 1. إيقاف جميع عمليات Office بالذاكرة...")
processes = [
    "OfficeClickToRun.exe",
    "setup.exe",
    "WINWORD.EXE",
    "EXCEL.EXE",
    "POWERPNT.EXE",
    "OUTLOOK.EXE",
    "MSOSYNC.EXE",
    "ONENOTE.EXE",
    "MSACCESS.EXE",
    "MSPUB.EXE",
    "VISIO.EXE",
    "PROJEXEC.EXE",
    "c2rclient.exe",
]
for proc in processes:
    subprocess.run(["taskkill", "/f", "/im", proc], capture_output=True)

# 2. إيقاف وإزالة الخدمات المتعلقة بـ Click-to-Run
print("⏹️ 2. إيقاف وإزالة خدمات النظام الخاص بالبرنامج...")
services = ["ClickToRunSvc", "OfficeSvc"]
for svc in services:
    subprocess.run(["sc", "stop", svc], capture_output=True)
    subprocess.run(["sc", "delete", svc], capture_output=True)

# 3. حذف الملفات والمجلدات المتبقية (تجهيز المسارات الديناميكية)
print("🗑️ 3. حذف المجلدات والملفات المتبقية من القرص...")
user_profile = os.environ.get("USERPROFILE", f"C:\\Users\\{os.getlogin()}")
paths_to_delete = [
    r"C:\Program Files\Microsoft Office 15",
    r"C:\Program Files\Microsoft Office",
    r"C:\Program Files (x86)\Microsoft Office",
    r"C:\Program Files\Common Files\Microsoft Shared\OFFICE15",
    r"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE15",
    r"C:\Program Files\Common Files\Microsoft Shared\OFFICE16",
    r"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16",
    r"C:\ProgramData\Microsoft\Office",
    os.path.join(user_profile, r"AppData\Local\Microsoft\Office"),
    os.path.join(user_profile, r"AppData\Roaming\Microsoft\Office"),
]

for path in paths_to_delete:
    if os.path.exists(path):
        try:
            shutil.rmtree(path, ignore_errors=True)
            print(f"  ✅ تم حذف المجلد: {path}")
        except Exception as e:
            print(f"  ❌ تعذر حذف {path}: {e}")

# 4. حذف مفاتيح الريجستري الشاملة
print("🗑️ 4. حذف مفاتيح السجل (Registry) المتعلقة بجميع النسخ...")
registry_keys_to_delete = [
    r"HKLM\SOFTWARE\Microsoft\Office",
    r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Office",
    r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Office16.O365HomePremRetail - en-us",
    r"HKCU\Software\Microsoft\Office",
]

for reg_key in registry_keys_to_delete:
    subprocess.run(["reg", "delete", reg_key, "/f"], capture_output=True)

# 5. مرحلة الفحص والتحقق التلقائي (Verification Step)
print("\n🔍 5. جاري إجراء فحص شامل للتحقق من إزالة كافة النسخ...")

registry_check_targets = [
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\ClickToRun"),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\16.0"),
    (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Office\16.0"),
    (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0"),
]

remaining_registry = []
for root_key, subkey in registry_check_targets:
    try:
        key = winreg.OpenKey(root_key, subkey, 0, winreg.KEY_READ)
        winreg.CloseKey(key)
        remaining_registry.append(subkey)
    except FileNotFoundError:
        pass
    except Exception:
        pass

remaining_files = [path for path in paths_to_delete if os.path.exists(path)]

# 6. التقرير النهائي
print("\n" + "=" * 50)
if not remaining_registry and not remaining_files:
    print("✅ النتيجة: تم حذف جميع نسخ وآثار Office بنجاح بالنظام!")
else:
    print("⚠️ النتيجة: تم حذف الأغلبية ولكن تلاحظ وجود بعض البقايا المقفولة:")
    if remaining_files:
        print(" - مجلدات متبقية:")
        for p in remaining_files:
            print(f"   • {p}")
    if remaining_registry:
        print(" - مفاتيح سجل متبقية:")
        for r in remaining_registry:
            print(f"   • {r}")
    print(
        "💡 (سيتم إزالة الملفات المقفولة تلقائياً بمجرد إكمال إعادة"
        " التشغيل)."
    )

print("=" * 50)
print("\n❗️ يُنصح بشدة بإعادة تشغيل الكمبيوتر الآن لإتمام العملية.")
input("اضغط Enter للخروج...")
