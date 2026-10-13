import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
target_form_path = os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf")
doc = fitz.open(target_form_path)

print(f"Target Form: {target_form_path}")
print(f"Pages: {len(doc)}")
for i, p in enumerate(doc):
    text = p.get_text()
    first_lines = [line.strip() for line in text.split('\n') if line.strip()][:5]
    print(f"Page {i+1} first lines: {first_lines}")
    widgets = list(p.widgets())
    print(f"Page {i+1} widget count: {len(widgets)}")

user_oos_path = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
doc_user = fitz.open(user_oos_path)
print(f"\nUser OOS: {user_oos_path}")
print(f"Pages: {len(doc_user)}")
for i, p in enumerate(doc_user):
    widgets = list(p.widgets())
    print(f"Page {i+1} widget count: {len(widgets)}")
