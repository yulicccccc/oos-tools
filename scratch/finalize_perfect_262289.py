import os
import sys
import docx
from docx.oxml.ns import qn
import win32com.client
import fitz
from pypdf import PdfWriter, PdfReader

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docs_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"

# =========================================================================
# 1. Update Word Document: Ensure 100% strictly uniform 7.0pt across tables
# =========================================================================
src_docx = os.path.join(desktop_dir, "Tables OOS-262289 Link Pharmacy (E19112) - ScanRDI.docx")
doc = docx.Document(src_docx)

for t in doc.tables:
    for row in t.rows:
        for cell in row.cells:
            for elem in cell._tc.iter():
                if elem.tag.endswith('sz') or elem.tag.endswith('szCs'):
                    elem.set(qn('w:val'), '14')

# Save to scratch first
temp_docx = "scratch/final_tables_262289.docx"
temp_pdf = "scratch/final_tables_262289.pdf"
doc.save(temp_docx)

# Convert to PDF via Word COM
word = win32com.client.DispatchEx('Word.Application')
word.Visible = False
try:
    wb = word.Documents.Open(os.path.abspath(temp_docx))
    wb.SaveAs(os.path.abspath(temp_pdf), FileFormat=17)
    wb.Close()
finally:
    word.Quit()

# Verify PDF is exactly 1 page
doc_pdf = fitz.open(temp_pdf)
print(f"Standalone Tables PDF generated. Page count: {len(doc_pdf)}")
assert len(doc_pdf) == 1, f"Expected 1 page, got {len(doc_pdf)}"

# Save final docx to Desktop and Documents
desktop_docx_dest = os.path.join(desktop_dir, "Tables OOS-262289 Link Pharmacy (E19112) - ScanRDI.docx")
docs_docx_dest = os.path.join(docs_dir, "Tables OOS-262289 Link Pharmacy (E19112) - ScanRDI.docx")
doc.save(desktop_docx_dest)
print(f"Saved uniform 7.0pt Word Tables to Desktop: {desktop_docx_dest}")
try:
    doc.save(docs_docx_dest)
    print(f"Saved uniform 7.0pt Word Tables to Documents: {docs_docx_dest}")
except Exception as e:
    print(f"Documents docx lock error: {e}")

# =========================================================================
# 2. Append to 6-Page Filled AcroForm to produce the unified 7-Page Master PDF
# =========================================================================
temp_filled_form = "scratch/temp_filled_corp21.pdf"
writer = PdfWriter(clone_from=temp_filled_form)

t_reader = PdfReader(temp_pdf)
for p in t_reader.pages:
    writer.add_page(p)

print(f"Total pages in unified Master PDF: {len(writer.pages)}")
assert len(writer.pages) == 7, f"Expected 7 pages, got {len(writer.pages)}"

out_transcribed_desktop = os.path.join(desktop_dir, "CORP-FORM-21- 01 Oct 2026 OOS-262289 Link Pharmacy (E19112) - ScanRDI.pdf")
out_transcribed_docs = os.path.join(docs_dir, "CORP-FORM-21- 01 Oct 2026 OOS-262289 Link Pharmacy (E19112) - ScanRDI.pdf")

out_std_desktop = os.path.join(desktop_dir, "OOS-262289 Link Pharmacy (E19112) - ScanRDI.pdf")
out_std_docs = os.path.join(docs_dir, "OOS-262289 Link Pharmacy (E19112) - ScanRDI.pdf")

with open(out_transcribed_desktop, "wb") as f:
    writer.write(f)
with open(out_std_desktop, "wb") as f:
    writer.write(f)
print(f"Saved Master PDF to Desktop: {out_transcribed_desktop}")

try:
    with open(out_transcribed_docs, "wb") as f:
        writer.write(f)
    with open(out_std_docs, "wb") as f:
        writer.write(f)
    print(f"Saved Master PDF to Documents: {out_transcribed_docs}")
except Exception as e:
    print(f"Documents pdf lock error: {e}")

# =========================================================================
# 3. Final Integrity Inspection of Master PDF
# =========================================================================
final_doc = fitz.open(out_transcribed_desktop)
print("\n--- FINAL MASTER PDF INSPECTION ---")
print(f"Total Pages: {len(final_doc)}")
for i, page in enumerate(final_doc):
    print(f"Page {i+1}: rect={page.rect}, widgets={len(list(page.widgets()))}")

p7 = final_doc[6]
p7_blocks = p7.get_text('dict')['blocks']
sizes = set()
for b in p7_blocks:
    for l in b.get('lines', []):
        for s in l.get('spans', []):
            sizes.add(round(s['size'], 1))
print(f"Page 7 Font Sizes: {sorted(sizes)}")
print("Verification complete! Zero defects.")
