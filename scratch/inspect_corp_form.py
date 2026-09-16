import pypdf, json

pdf_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\CORP-FORM-21 Laboratory OOS Investigation Form (v11.1) (1).pdf"
reader = pypdf.PdfReader(pdf_path)
print("Page count:", len(reader.pages))

fields = reader.get_fields()
print("Total fields:", len(fields) if fields else 0)

template_path = r"ScanRDI OOS template.pdf"
reader_tpl = pypdf.PdfReader(template_path)
tpl_fields = reader_tpl.get_fields()
print("Template fields count:", len(tpl_fields) if tpl_fields else 0)

# Check if fields match
new_keys = set(fields.keys()) if fields else set()
tpl_keys = set(tpl_fields.keys()) if tpl_fields else set()
print("Intersection count:", len(new_keys & tpl_keys))
print("Only in new form:", len(new_keys - tpl_keys))
print("Only in template:", len(tpl_keys - new_keys))

if new_keys - tpl_keys:
    print("Sample keys only in new:", list(new_keys - tpl_keys)[:20])
if tpl_keys - new_keys:
    print("Sample keys only in tpl:", list(tpl_keys - new_keys)[:20])


