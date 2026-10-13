import os
import fitz
from check_user_layout import p_p3, p_p4, p_p5

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
src_pdf = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
doc = fitz.open(src_pdf)

w3 = [w for w in doc[2].widgets() if w.field_name == 'Text Field49'][0].field_value
w4 = [w for w in doc[3].widgets() if w.field_name == 'Text Field50'][0].field_value
w5 = [w for w in doc[4].widgets() if w.field_name == 'Text Field51'][0].field_value

full_desktop_text = (w3 + "\n\n" + w4 + "\n\n" + w5).replace('\r', '').strip()
full_user_layout_text = ("\n\n".join(p_p3) + "\n\n" + "\n\n".join(p_p4) + "\n\n" + "\n\n".join(p_p5)).replace('\r', '').strip()

print(f"Full desktop text len: {len(full_desktop_text)}")
print(f"Full user layout text len: {len(full_user_layout_text)}")
print(f"Are they identical? {full_desktop_text == full_user_layout_text}")

if full_desktop_text != full_user_layout_text:
    import difflib
    diff = list(difflib.unified_diff(full_desktop_text.splitlines(), full_user_layout_text.splitlines(), lineterm=''))
    print("Diff lines:")
    for d in diff[:20]:
        print(d)
