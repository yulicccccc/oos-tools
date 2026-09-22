import fitz

pdf_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf"
doc = fitz.open(pdf_path)

for p_num in range(len(doc)):
    p = doc[p_num]
    print(f"=== Page {p_num+1} ===")
    for w in p.widgets():
        if w.rect.y0 < 225 and w.rect.x0 > 400:
            print(f"  Widget: {w.field_name} (type={w.field_type}) rect={w.rect}")
    for b in p.get_text("words"):
        if b[1] < 225 and b[0] > 400:
            print(f"  Word: {b[4]} rect={b[:4]}")
