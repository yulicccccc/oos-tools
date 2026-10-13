import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
target_form_path = os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf")
source_pdf_path = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")

doc_src = fitz.open(source_pdf_path)
doc_tgt = fitz.open(target_form_path)

p3_src = [w for w in doc_src[2].widgets() if w.field_name == 'Text Field49'][0]
p3_tgt = [w for w in doc_tgt[2].widgets() if w.field_name == 'Text Field49'][0]

print(f"Source Page 3 Text Field49 rect: {p3_src.rect}, height={p3_src.rect.height}")
print(f"Target Page 3 Text Field49 rect: {p3_tgt.rect}, height={p3_tgt.rect.height}")
