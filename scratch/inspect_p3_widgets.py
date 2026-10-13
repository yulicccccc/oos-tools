import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
target_form = os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf")
doc = fitz.open(target_form)
p3 = doc[2]
for w in p3.widgets():
    rect = w.rect
    # get text around rect
    search_rect = fitz.Rect(rect.x0 - 50, rect.y0 - 20, rect.x1 + 50, rect.y1 + 20)
    surrounding = p3.get_text("text", clip=search_rect).strip().replace('\n', ' ')
    print(f"Widget {w.field_name} (type: {w.field_type_string}): rect={w.rect}, surrounding='{surrounding}'")
