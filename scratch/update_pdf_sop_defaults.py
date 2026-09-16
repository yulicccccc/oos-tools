from pypdf import PdfWriter, PdfReader
import os

pdf_targets = ["ScanRDI OOS template.pdf", "ScanRDI OOS P1 template.pdf"]

field_updates = {
    'Text Field8': "MICRO-SOP-12 (16)\rENG-SOP-4 (05)",
    'Text Field9': "24Jul26\r24Jul26",
    'Text Field10': "Rev: 16\rRev: 05",
    'Text Field15': "Yes, as per MICRO-SOP-12, ENG-SOP-4",
    'Text Field16': "Yes, as per MICRO-SOP-12, ENG-SOP-4",
}

for pdf_file in pdf_targets:
    if not os.path.exists(pdf_file):
        continue
    reader = PdfReader(pdf_file)
    writer = PdfWriter()
    writer.append(reader)
    for p in writer.pages:
        writer.update_page_form_field_values(p, field_updates)
    with open(pdf_file, "wb") as f:
        writer.write(f)
    print(f"Updated default fields in {pdf_file} successfully!")
