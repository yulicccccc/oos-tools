import time
import shutil
import sys

src = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\CORP-FORM-22- 13 Oct 2026 OOS-262329 GoGoMeds Select (E10747) - Phase II.pdf"
dst = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-22- 13 Oct 2026 OOS-262329 GoGoMeds Select (E10747) - Phase II.pdf"

print("Waiting for file lock on Desktop to release...", flush=True)
for i in range(15):
    try:
        shutil.copy2(src, dst)
        print("SUCCESS: Desktop CORP-FORM-22 successfully updated with Day 4!", flush=True)
        sys.exit(0)
    except Exception:
        time.sleep(1)

print("Still locked, leaving (Updated Day 4).pdf on Desktop as fallback.", flush=True)
