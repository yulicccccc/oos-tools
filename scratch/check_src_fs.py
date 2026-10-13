import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
src_pdf = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
doc = fitz.open(src_pdf)

for pno, p in enumerate(doc):
    print(f"\n--- Page {pno+1} font sizes in source ---")
    for w in p.widgets():
        if w.field_value and w.field_type_string == 'Text':
            if w.text_fontsize != 8.5:
                print(f"  {w.field_name}: fs={w.text_fontsize}, val={repr(w.field_value[:30])}")
