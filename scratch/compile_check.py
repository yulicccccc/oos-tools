import glob
import py_compile
import sys

sys.stdout.reconfigure(encoding='utf-8')

py_files = glob.glob("*.py") + glob.glob("pages/*.py")

print("Checking Python syntax for all files...")
for fname in py_files:
    try:
        py_compile.compile(fname, doraise=True)
        print(f"  {fname}: [OK] Valid Syntax")
    except Exception as e:
        print(f"  {fname}: [ERROR] SYNTAX ERROR: {e}")
