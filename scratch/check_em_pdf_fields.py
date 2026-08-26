import pypdf

pdf_path = "EM OOS P1 template.pdf"
reader = pypdf.PdfReader(pdf_path)
fields = reader.get_fields()

print(f"Total AcroForm Fields in {pdf_path}: {len(fields)}")
for fname in fields.keys():
    print(f"  - '{fname}'")
