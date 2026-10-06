import os
import sys
import docx
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
import win32com.client
import fitz
import shutil

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

tables_docx_src = os.path.join(SCRATCH_DIR, "Celsis table OOS-262080.docx")
tables_docx_out = os.path.join(DESKTOP_DIR, "Celsis table OOS-262080 (Updated).docx")
tables_pdf_out = os.path.join(DESKTOP_DIR, "Celsis table OOS-262080 (Updated).pdf")
tables_pdf_orig = os.path.join(DESKTOP_DIR, "Celsis table OOS-262080.pdf")

URL_0487 = "https://etrax.eagleanalytical.com/SubmissionTest/Details/fd3G2StZClcy1TP2ES6BLw__"
SAMPLE_0487 = "ETX-260914-0487"

print("=== STEP 1: UPDATE Table 2 from SCRATCH DOCX ===")
# If scratch docx exists, load from scratch docx (or copy desktop to scratch)
doc_tbl = docx.Document(tables_docx_src)
t2 = doc_tbl.tables[1]

# Find row for 0487
target_row = None
for r in t2.rows:
    text = ' '.join(c.text for c in r.cells)
    if SAMPLE_0487 in text:
        target_row = r
        break

assert target_row is not None, "0487 row not found in Table 2!"

# Update ETX ID cell with clickable link
for i, c in enumerate(target_row.cells):
    if SAMPLE_0487 in c.text:
        part = c.part
        r_id = part.relate_to(URL_0487, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        tc_xml = (
            f'<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<w:tcPr>'
            f'<w:tcW w:w="1444" w:type="dxa"/>'
            f'<w:tcBorders>'
            f'<w:top w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
            f'<w:left w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
            f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
            f'<w:right w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
            f'</w:tcBorders>'
            f'<w:vAlign w:val="center"/>'
            f'</w:tcPr>'
            f'<w:p>'
            f'<w:pPr>'
            f'<w:spacing w:line="360" w:lineRule="auto"/>'
            f'<w:jc w:val="center"/>'
            f'<w:rPr>'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'<w:sz w:val="14"/>'
            f'<w:szCs w:val="14"/>'
            f'</w:rPr>'
            f'</w:pPr>'
            f'<w:hyperlink r:id="{r_id}" w:history="1">'
            f'<w:r>'
            f'<w:rPr>'
            f'<w:rStyle w:val="Hyperlink"/>'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'<w:color w:val="0000FF"/>'
            f'<w:sz w:val="14"/>'
            f'<w:szCs w:val="14"/>'
            f'</w:rPr>'
            f'<w:t>{SAMPLE_0487}</w:t>'
            f'</w:r>'
            f'</w:hyperlink>'
            f'</w:p>'
            f'</w:tc>'
        )
        new_tc = parse_xml(tc_xml)
        c._tc.getparent().replace(c._tc, new_tc)
        print(f"  Inserted hyperlink for {SAMPLE_0487} in Table 2")

# Update Microbial ID cell
for i, c in enumerate(target_row.cells):
    if 'coccobacilli' in c.text or 'rods' in c.text or 'Corynebacterium' in c.text:
        for p in c.paragraphs:
            p.text = ""
        run = c.paragraphs[0].add_run("Corynebacterium ureicelerivorans &\nMycobacterium grossiae")
        run.font.name = "Times New Roman"
        run.font.size = docx.shared.Pt(6.5)
        c.paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        print(f"  Updated Microbial ID to 'Corynebacterium ureicelerivorans & Mycobacterium grossiae' in Table 2")

scratch_tables_docx = os.path.join(SCRATCH_DIR, "Celsis table OOS-262080.docx")
doc_tbl.save(scratch_tables_docx)
print(f"Saved scratch tables DOCX: {scratch_tables_docx}")

try:
    doc_tbl.save(tables_docx_out)
    print(f"Saved updated desktop tables DOCX: {tables_docx_out}")
except Exception as e:
    print(f"Could not save {tables_docx_out}: {e}")

try:
    doc_tbl.save(os.path.join(DESKTOP_DIR, "Celsis table OOS-262080.docx"))
    print(f"Saved to Desktop Celsis table OOS-262080.docx")
except Exception as e:
    print(f"File locked on Desktop: {e}")

print("\n=== STEP 2: CONVERT TO PDF VIA WORD COM ===")
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0

doc_com = word.Documents.Open(scratch_tables_docx)
scratch_tables_pdf = os.path.join(SCRATCH_DIR, "Celsis table OOS-262080.pdf")
doc_com.SaveAs2(scratch_tables_pdf, FileFormat=17)
doc_com.Close()
word.Quit()

try:
    shutil.copy2(scratch_tables_pdf, tables_pdf_orig)
    print(f"Copied to {tables_pdf_orig}")
except Exception as e:
    print(f"Locked {tables_pdf_orig}: {e}")

shutil.copy2(scratch_tables_pdf, tables_pdf_out)
print(f"Copied to {tables_pdf_out}")

print("\n=== STEP 3: VERIFY NEW PDF LINKS & TEXT ===")
doc_pdf = fitz.open(scratch_tables_pdf)
print(f"Pages: {len(doc_pdf)}")
for i, p in enumerate(doc_pdf):
    print(f"Page {i+1} links: {p.get_links()}")
    t = p.get_text()
    if 'Corynebacterium' in t:
        print(f"  Found 'Corynebacterium' on Page {i+1}!")
    if 'Micrococcus' in t:
        print(f"  Found 'Micrococcus' on Page {i+1}!")
doc_pdf.close()
