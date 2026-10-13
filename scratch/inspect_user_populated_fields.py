import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
user_pdf = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
doc = fitz.open(user_pdf)

print(f"User PDF: {user_pdf}")
print(f"Pages: {len(doc)}")

for pno, page in enumerate(doc):
    print(f"\n==================== PAGE {pno+1} ====================")
    for w in page.widgets():
        val = w.field_value
        if val != '' and val != 'Off' and val is not None:
            # print field name, type, font size, and value (first 100 chars)
            val_str = str(val).replace('\n', '\\n').replace('\r', '\\r')
            if len(val_str) > 100:
                val_str = val_str[:100] + "..."
            print(f"  {w.field_name} ({w.field_type_string}, fs={w.text_fontsize}): {val_str}")
