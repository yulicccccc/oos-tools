import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 - P1 - 16 Sep 2026.pdf")
print("=== PRE-FILLED FIELDS IN USER PDF ===")
for i, page in enumerate(doc):
    for w in page.widgets():
        if w.field_value:
            print(f"Page {i+1}: {w.field_name} ({w.field_type}) = {repr(w.field_value)}")
