import os
import glob
import fitz # PyMuPDF
from datetime import datetime

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
print("Scanning Desktop files:")
for f in os.listdir(desktop):
    if any(k in f.lower() for k in ["corp-form-21", "oos-262017", "celsis"]):
        full = os.path.join(desktop, f)
        mtime = datetime.fromtimestamp(os.path.getmtime(full)).strftime('%Y-%m-%d %H:%M:%S')
        print(f"  {f} ({os.path.getsize(full)} bytes, modified: {mtime})")

target_form = os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf")
if os.path.exists(target_form):
    doc = fitz.open(target_form)
    print(f"\n--- Target Form: {target_form} ---")
    print(f"Page count: {len(doc)}")
    for i, page in enumerate(doc):
        widgets = list(page.widgets())
        print(f"Page {i+1}: {len(widgets)} widgets")
        for w in widgets:
            print(f"  Widget: name='{w.field_name}', type={w.field_type_string}, val='{w.field_value}'")
else:
    print(f"\nTarget form NOT found: {target_form}")

user_oos_pdf = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
if os.path.exists(user_oos_pdf):
    doc_oos = fitz.open(user_oos_pdf)
    print(f"\n--- User OOS PDF: {user_oos_pdf} ---")
    print(f"Page count: {len(doc_oos)}")
    for i, page in enumerate(doc_oos):
        widgets = list(page.widgets())
        for w in widgets:
            if "qyc" in str(w.field_value).lower() or "n/a" in str(w.field_value).lower() or "07oct" in str(w.field_value).lower() or "05oct" in str(w.field_value).lower():
                print(f"Page {i+1} field '{w.field_name}': '{w.field_value}'")
