import os
import shutil
from datetime import datetime
from pypdf import PdfReader, PdfWriter
import pypdfium2 as pdfium

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docs_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
history_dir = os.path.join(docs_dir, ".history")
os.makedirs(history_dir, exist_ok=True)

corp_form_path = os.path.join(desktop_dir, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")
source_oos_pdf = os.path.join(desktop_dir, "OOS-261187 ScanC_O CGS E001309 S1 11MAY2026 - EM.pdf")
user_table_pdf = os.path.join(desktop_dir, "EM table OOS-261187 11MAY2026.pdf")

# 1. Backup original CORP-FORM-21 to .history
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
backup_name = f"CORP-FORM-21 Laboratory OOS Investigation Form (v11.1)_backup_{timestamp}.pdf"
backup_path = os.path.join(history_dir, backup_name)
shutil.copy2(corp_form_path, backup_path)
print(f"Backed up original CORP-FORM-21 to: {backup_path}")

# 2. Extract fields from source OOS PDF
r_src = PdfReader(source_oos_pdf)
src_fields = r_src.get_fields()

pdf_map = {}
for k, v in src_fields.items():
    val = v.get('/V')
    if val is not None:
        pdf_map[k] = val

print(f"Extracted {len(pdf_map)} fields from source OOS PDF.")

# 3. Read CORP-FORM-21 base template
writer = PdfWriter(clone_from=corp_form_path)

# Ensure it only has pages 1-6 before updating
while len(writer.pages) > 6:
    del writer.pages[-1]

# Populate all fields across pages
for p in writer.pages:
    writer.update_page_form_field_values(p, pdf_map)

# 4. Append Page 7 (User Table PDF)
if os.path.exists(user_table_pdf):
    table_reader = PdfReader(user_table_pdf)
    for tp in table_reader.pages:
        writer.add_page(tp)
    print(f"Appended Table PDF as Page 7. Total pages: {len(writer.pages)}")
else:
    print("WARNING: User table PDF not found!")

# 5. Write to CORP-FORM-21 on Desktop directly
with open(corp_form_path, "wb") as f:
    writer.write(f)
print(f"Successfully transcribed into: {corp_form_path}")

# Also update the OOS-named PDF on Desktop and Documents with this official v11.1 base!
oos_named_desktop = os.path.join(desktop_dir, "OOS-261187 ScanC_O CGS E001309 S1 11MAY2026 - EM.pdf")
oos_named_docs = os.path.join(docs_dir, "OOS-261187 ScanC_O CGS E001309 S1 11MAY2026 - EM.pdf")

shutil.copy2(corp_form_path, oos_named_desktop)
shutil.copy2(corp_form_path, oos_named_docs)
print(f"Updated: {oos_named_desktop}")
print(f"Updated: {oos_named_docs}")

# 6. Render images of all 7 pages to verify visual quality
scratch_dir = os.path.join(docs_dir, "scratch")
pdf_doc = pdfium.PdfDocument(corp_form_path)
for idx, page in enumerate(pdf_doc):
    img = page.render(scale=2).to_pil()
    out_img = os.path.join(scratch_dir, f"corp_form_p{idx+1}.png")
    img.save(out_img)
    print(f"Rendered Page {idx+1} to: {out_img}")
