import os, sys
import fitz

sys.stdout.reconfigure(encoding='utf-8')
desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
pdf_path = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
doc = fitz.open(pdf_path)

for pno in [2, 3, 4]:
    print(f"\n==================== PAGE {pno+1} TEXT ====================")
    w_narrative = [w for w in doc[pno].widgets() if w.field_name in ['Text Field49', 'Text Field50', 'Text Field51']][0]
    print(w_narrative.field_value)
