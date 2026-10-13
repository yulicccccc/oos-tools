import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
source_pdf_path = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
doc_src = fitz.open(source_pdf_path)

val_p3 = [w for w in doc_src[2].widgets() if w.field_name == 'Text Field49'][0].field_value
val_p4 = [w for w in doc_src[3].widgets() if w.field_name == 'Text Field50'][0].field_value
val_p5 = [w for w in doc_src[4].widgets() if w.field_name == 'Text Field51'][0].field_value

print(f"val_p3 length: {len(val_p3)}")
print("val_p3 text:\n", val_p3[:100], "...", val_p3[-100:])

print(f"\nval_p4 length: {len(val_p4)}")
print("val_p4 text:\n", val_p4[:100], "...", val_p4[-100:])

print(f"\nval_p5 length: {len(val_p5)}")
print("val_p5 text:\n", val_p5[:100], "...", val_p5[-100:])
