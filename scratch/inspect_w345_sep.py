import os
import fitz
from check_user_layout import p_p3, p_p4, p_p5

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
src_pdf = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
doc = fitz.open(src_pdf)

w3 = [w for w in doc[2].widgets() if w.field_name == 'Text Field49'][0].field_value
w4 = [w for w in doc[3].widgets() if w.field_name == 'Text Field50'][0].field_value
w5 = [w for w in doc[4].widgets() if w.field_name == 'Text Field51'][0].field_value

print("W3 repr first 200:", repr(w3[:200]))
print("W4 repr first 200:", repr(w4[:200]))
print("W5 repr first 200:", repr(w5[:200]))
