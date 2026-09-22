import fitz

pdf_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf"
doc = fitz.open(pdf_path)

for p_num in [0, 1]:
    p = doc[p_num]
    print(f"=== Page {p_num+1} ===")
    for w in p.widgets():
        print(f"  Widget: {w.field_name} (type={w.field_type}) rect={w.rect}")
