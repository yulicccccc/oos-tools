import glob, fitz

pdfs = [
    r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\oos--262158.pdf",
    r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 - P1 - 16 Sep 2026.pdf",
]
for p in glob.glob(r"scratch/*QYC*.pdf") + glob.glob(r"scratch/*2620*.pdf"):
    pdfs.append(p)

for pdf in pdfs[:6]:
    try:
        doc = fitz.open(pdf)
        if len(doc) >= 6:
            p6 = doc[5]
            print(f"\n=== {pdf} (Page 6) ===")
            for w in p6.widgets():
                if w.field_value:
                    val = repr(w.field_value[:60].replace('\n', ' '))
                    print(f"  {w.field_name}: {val}")
    except Exception as e:
        pass
