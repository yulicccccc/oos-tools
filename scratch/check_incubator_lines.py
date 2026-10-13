import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\oos--262158.pdf")
p = doc[1]
for b in p.get_text("dict")["blocks"]:
    if "lines" in b:
        for l in b["lines"]:
            for s in l["spans"]:
                if any(k in s["text"] for k in ["Incubator", "Sensor", "2027"]):
                    print(f"{s['bbox']}: sz={s['size']:.2f}, text={s['text']}")
