import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\oos--262158.pdf")
for pno in range(6):
    print(f"\n=================== PAGE {pno+1} ===================")
    p = doc[pno]
    for b in p.get_text("dict")["blocks"]:
        if "lines" in b:
            for l in b["lines"]:
                for s in l["spans"]:
                    txt = s["text"].strip()
                    if txt:
                        # Print only if likely filled content
                        if any(k in txt for k in ["Yes", "Not Applicable", "See", "091626", "MICRO", "ENG", "Jorge", "Believe", "Tirzepatide", "E00", "S.aureus", "04232026", "Nov 2026", "Aug 2027", "July 2027", "N/A"]):
                            print(f"  size={s['size']:.2f}, font={s['font']}, text={txt}")
