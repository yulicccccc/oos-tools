import fitz

pdf_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 - P1 - 16 Sep 2026.pdf"
doc = fitz.open(pdf_path)
print("Page count:", len(doc))
for i, p in enumerate(doc):
    print(f"--- PAGE {i+1} ---")
    widgets = list(p.widgets())
    print(f"Widgets count: {len(widgets)}")
    for w in widgets:
        if w.field_value:
            print(f"  {w.field_name} ({w.field_type}): {repr(w.field_value[:50])}")
    # Also print text snippets
    txt = p.get_text("text").strip()
    lines = [line.strip() for line in txt.split("\n") if line.strip()]
    print("  Text lines count:", len(lines))
    if len(lines) > 0:
        print("  First 5 text lines:", lines[:5])
