import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\oos--262158.pdf")
p = doc[0]
for b in p.get_text("dict")["blocks"]:
    if "lines" in b:
        for l in b["lines"]:
            for s in l["spans"]:
                if "16-Sep-2026" in s["text"]:
                    print(f"Date: font={s['font']}, size={s['size']:.2f}, text={s['text']}")
                elif "Kathan" in s["text"] or "Injectable" in s["text"] or "ETX-" in s["text"]:
                    print(f"Header: font={s['font']}, size={s['size']:.2f}, text={s['text']}")
