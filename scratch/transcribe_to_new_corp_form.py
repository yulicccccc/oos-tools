import os, sys, shutil
from pypdf import PdfReader, PdfWriter
import fitz

OOS_ROOT = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS'
DOCUMENTS_DIR = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents'
DESKTOP_DIR = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop'

FRESH_FORM = os.path.join(DESKTOP_DIR, 'CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf')
SOURCE_FINAL_PDF = os.path.join(DESKTOP_DIR, 'OOS-262074 Hillstone Pharmacy (E17047) - USP71 - QYC.pdf')

OUTPUT_PDF_NAME = 'OOS-262074 Hillstone Pharmacy (E17047) - USP71 - QYC.pdf'
OUTPUT_SCRATCH = os.path.join(OOS_ROOT, 'scratch', OUTPUT_PDF_NAME)
OUTPUT_DOCUMENTS = os.path.join(DOCUMENTS_DIR, OUTPUT_PDF_NAME)
OUTPUT_DESKTOP = os.path.join(DESKTOP_DIR, OUTPUT_PDF_NAME)

print('=== STARTING TRANSCRIPTION TO FRESH CORP-FORM-21 (v11.1) ===')
print('Fresh Template:', FRESH_FORM)
print('Source Final:', SOURCE_FINAL_PDF)

# 1. Read Source PDF to extract all field values
r_src = PdfReader(SOURCE_FINAL_PDF)
src_fields = r_src.get_fields()

field_map = {}
for k, obj in src_fields.items():
    v = obj.get('/V', None)
    if v is not None and v != '':
        field_map[k] = v

print(f'Extracted {len(field_map)} populated field values from source PDF.')

# 2. Populate Fresh Form (Pages 1 to 6)
writer = PdfWriter(clone_from=FRESH_FORM)

# Ensure fresh form has exactly 6 pages initially
while len(writer.pages) > 6:
    del writer.pages[-1]

for page in writer.pages:
    writer.update_page_form_field_values(page, field_map)

# 3. Extract Table Page (Page 7) from source PDF
if len(r_src.pages) >= 7:
    table_page = r_src.pages[6]
    writer.add_page(table_page)
    print('Appended Table 1 & Table 2 as Page 7. Total pages:', len(writer.pages))
else:
    print('Warning: Source PDF did not have 7 pages!')

# 4. Write intermediate PDF
temp_pdf = os.path.join(OOS_ROOT, 'scratch', 'transcribed_temp.pdf')
with open(temp_pdf, 'wb') as f:
    writer.write(f)

# 5. Fine-tune font sizes with PyMuPDF to ensure identical aesthetic layout
doc_fitz = fitz.open(temp_pdf)
font_sizes = {
    'Text Field49': 8.5,
    'Text Field50': 8.5,
    'Text Field51': 7.8,
}
for page in doc_fitz:
    for w in page.widgets():
        if w.field_name in font_sizes:
            w.text_fontsize = font_sizes[w.field_name]
            w.update()

doc_fitz.save(OUTPUT_SCRATCH)
doc_fitz.close()
if os.path.exists(temp_pdf):
    os.remove(temp_pdf)

print(f'Successfully transcribed and saved to: {OUTPUT_SCRATCH}')

# 6. Verify Page Count and Footer in Transcribed PDF
r_verify = PdfReader(OUTPUT_SCRATCH)
print(f'Verified Page Count: {len(r_verify.pages)} (Must be 7)')

p1_txt = r_verify.pages[0].extract_text()
for line in p1_txt.splitlines():
    if 'Printed by' in line or '18-Sep-2026' in line:
        print('Verified New ZenQMS Footer:', line.strip())

# 7. Sync to Documents and Desktop
for dest in [OUTPUT_DOCUMENTS, OUTPUT_DESKTOP]:
    try:
        shutil.copy2(OUTPUT_SCRATCH, dest)
        print('Synced to:', dest)
    except PermissionError:
        print(f'Notice: {dest} is open in another program, skipping overwrite.')
    except Exception as e:
        print(f'Sync error for {dest}: {e}')

print('=== TRANSCRIPTION COMPLETE! ===')
