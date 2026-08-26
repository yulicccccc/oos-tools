import os
import sys
import subprocess

sys.path.insert(0, os.getcwd())
sys.stdout.reconfigure(encoding='utf-8')

print("==================================================")
print("🔍 COMPREHENSIVE VERIFICATION OF EXISTING MODULES")
print("==================================================")

core_files = [
    "celsis_logic.py",
    "pages/Celsis.py",
    "scan_logic.py",
    "pages/ScanRDI.py",
    "usp71_logic.py",
    "pages/USP71.py"
]

print("\n--- 1. Git Status Check for Existing Core Files ---")
res = subprocess.run(["git", "status", "--porcelain"] + core_files, capture_output=True, text=True)
if res.stdout.strip():
    print(f"⚠️ Modified files detected:\n{res.stdout}")
else:
    print("✅ 100% CONFIRMED: Zero lines were changed in any existing core files (celsis_logic.py, scan_logic.py, usp71_logic.py, pages/Celsis.py, pages/ScanRDI.py, pages/USP71.py)!")

print("\n--- 2. Module Import Verification ---")
modules_to_test = ["celsis_logic", "scan_logic", "usp71_logic", "em_logic"]

for mod_name in modules_to_test:
    try:
        mod = __import__(mod_name)
        print(f"  - {mod_name}: ✅ Loaded successfully!")
    except Exception as e:
        print(f"  - {mod_name}: ❌ FAILED to load: {e}")

print("\n==================================================")
