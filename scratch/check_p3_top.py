import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\oos--262158.pdf")
p = doc[2]
for b in p.get_text("dict")["blocks"]:
    if "lines" in b:
        for l in b["lines"]:
            for s in l["spans"]:
                if s["bbox"][1] < 280:
                    print(f"{s['bbox']}: {s['text']}")
