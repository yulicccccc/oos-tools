import fitz

pdf_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261932_Boudreauxs New Drug Store (E01607) - RS Reviewed - Signed by DK-RS.pdf"

doc = fitz.open(pdf_path)
print(f"Total pages: {len(doc)}")

for page_num in range(len(doc)):
    page = doc[page_num]
    text = page.get_text()
    print(f"\n{'='*40} PAGE {page_num+1} {'='*40}")
    print(text[:1500])
    
    # Also inspect widget fields on this page
    widgets = list(page.widgets())
    if widgets:
        print(f"\n--- Widgets on Page {page_num+1} ({len(widgets)} widgets) ---")
        for w in widgets:
            if w.field_value and str(w.field_value).strip():
                print(f"  {w.field_name}: {repr(w.field_value[:80]) if len(str(w.field_value)) > 80 else repr(w.field_value)}")
