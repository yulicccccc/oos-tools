import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\oos--262158.pdf")
for pno in range(6):
    p = doc[pno]
    for b in p.get_text("dict")["blocks"]:
        if "lines" in b:
            for l in b["lines"]:
                for s in l["spans"]:
                    if "262158" in s["text"]:
                        print(f"Page {pno+1}: font={s['font']}, size={s['size']:.2f}, text={s['text']}")
