import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
target_form = fitz.open(os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf"))
user_oos = fitz.open(os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf"))

print("=== Target Form Field Values ===")
non_empty_target = 0
for i, page in enumerate(target_form):
    for w in page.widgets():
        if w.field_value:
            non_empty_target += 1
            print(f"P{i+1} {w.field_name}: '{w.field_value}'")
print(f"Total non-empty widgets in target: {non_empty_target}")

print("\n=== User OOS Field Values ===")
for i, page in enumerate(user_oos):
    for w in page.widgets():
        if w.field_value:
            print(f"P{i+1} {w.field_name} ({w.field_type_string}): '{w.field_value}'")
