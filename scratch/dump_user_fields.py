import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
user_oos_pdf = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
doc = fitz.open(user_oos_pdf)

with open(r"scratch\user_oos_fields_dump.txt", "w", encoding="utf-8") as out:
    for i, page in enumerate(doc):
        out.write(f"\n--- PAGE {i+1} ---\n")
        for w in page.widgets():
            out.write(f"name: '{w.field_name}', type: {w.field_type_string}, val: {repr(w.field_value)}\n")

print("Dumped all fields to scratch/user_oos_fields_dump.txt")
