import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 - P1 - 16 Sep 2026.pdf")
for pno in range(6):
    p = doc[pno]
    print(f"=== PAGE {pno+1} BLANK FIELDS ===")
    for w in p.widgets():
        if not w.field_value:
            print(f"  {w.field_name} (type={w.field_type}): rect={w.rect}")
