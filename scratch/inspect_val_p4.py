import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
source_pdf_path = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
doc_src = fitz.open(source_pdf_path)

val_p4 = [w for w in doc_src[3].widgets() if w.field_name == 'Text Field50'][0].field_value
print(f"val_p4 length: {len(val_p4)}")
print("val_p4 last 200 chars:")
print(repr(val_p4[-200:]))
