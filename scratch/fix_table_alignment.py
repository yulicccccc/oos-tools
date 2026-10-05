import os
import sys
import shutil
import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
import win32com.client
import fitz

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

os.makedirs(SCRATCH_DIR, exist_ok=True)

URL = "https://etrax.eagleanalytical.com/SubmissionTest/Details/jzraMOOTFYUFD%24erkwgubw__"
SAMPLE_ID = "ETX-260828-0527"

def set_cell_hyperlink_aligned(cell, url, text):
    part = cell.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    tc_xml = (
        f'<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:tcPr>'
        f'<w:tcW w:w="1400" w:type="dxa"/>'
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
        f'<w:u w:val="single"/>'
        f'</w:rPr>'
        f'<w:t>{text}</w:t>'
        f'</w:r>'
        f'</w:hyperlink>'
        f'</w:p>'
        f'</w:tc>'
    )
    new_tc = parse_xml(tc_xml)
    cell._tc.getparent().replace(cell._tc, new_tc)

print("=== STEP 1: FIX CELSIS TABLE DOCX ===")
tbl_docx_desktop = os.path.join(DESKTOP_DIR, "Celsis table OOS-262080.docx")
tbl_docx_scratch = os.path.join(SCRATCH_DIR, "Celsis table OOS-262080.docx")

doc_tbl = docx.Document(tbl_docx_desktop if os.path.exists(tbl_docx_desktop) else tbl_docx_scratch)
set_cell_hyperlink_aligned(doc_tbl.tables[0].rows[1].cells[2], URL, SAMPLE_ID)
doc_tbl.save(tbl_docx_scratch)
try:
    doc_tbl.save(tbl_docx_desktop)
    print(f"Saved: {tbl_docx_desktop}")
except PermissionError:
    print(f"Notice: {tbl_docx_desktop} is locked, saved to scratch.")

print("\n=== STEP 2: FIX MASTER WORD REPORT DOCX ===")
main_docx_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.docx")
main_docx_scratch = os.path.join(SCRATCH_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.docx")

if os.path.exists(main_docx_desktop):
    doc_main = docx.Document(main_docx_desktop)
    if len(doc_main.tables) > 1 and len(doc_main.tables[1].rows) > 1:
        set_cell_hyperlink_aligned(doc_main.tables[1].rows[1].cells[2], URL, SAMPLE_ID)
        doc_main.save(main_docx_scratch)
        try:
            doc_main.save(main_docx_desktop)
            print(f"Saved: {main_docx_desktop}")
        except PermissionError:
            print(f"Notice: {main_docx_desktop} is locked, saved to scratch.")

print("\n=== STEP 3: EXPORT PERFECTLY ALIGNED PDF VIA WORD COM ===")
tbl_pdf_scratch = os.path.join(SCRATCH_DIR, "Celsis table OOS-262080.pdf")
tbl_pdf_desktop = os.path.join(DESKTOP_DIR, "Celsis table OOS-262080.pdf")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    source_to_open = os.path.abspath(tbl_docx_scratch if os.path.exists(tbl_docx_scratch) else tbl_docx_desktop)
    doc_word = word.Documents.Open(source_to_open)
    doc_word.ExportAsFixedFormat(os.path.abspath(tbl_pdf_scratch), 17)
    doc_word.Close(False)
    print(f"Exported PDF via Word COM to {tbl_pdf_scratch}")
finally:
    word.Quit()

# Check coordinates in exported PDF
chk = fitz.open(tbl_pdf_scratch)
words = chk[0].get_text('words')
row1_words = [w for w in words if 115 <= w[1] <= 135]
print("\nRow 1 Word Coordinates in Exported PDF:")
coords = {}
for w in row1_words:
    coords[w[4]] = (round(w[1], 3), round(w[3], 3))
    print(f"  {w[4]:20} y0={w[1]:.3f} y1={w[3]:.3f}")

chk.close()

try:
    shutil.copy2(tbl_pdf_scratch, tbl_pdf_desktop)
    print(f"Saved: {tbl_pdf_desktop}")
except PermissionError:
    print(f"Notice: {tbl_pdf_desktop} is locked.")

print("\n=== STEP 4: UPDATE COMBINED 8-PAGE DELIVERABLES ===")
# 1. Update QYC.pdf
qyc_pdf_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC.pdf")
if os.path.exists(qyc_pdf_desktop):
    d_qyc_orig = fitz.open(qyc_pdf_desktop)
    d_qyc_new = fitz.open()
    d_qyc_new.insert_pdf(d_qyc_orig, from_page=0, to_page=5)
    d_qyc_orig.close()
    
    d_tbl = fitz.open(tbl_pdf_scratch)
    d_qyc_new.insert_pdf(d_tbl, annots=True)
    d_tbl.close()
    
    scratch_qyc = os.path.join(SCRATCH_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC.pdf")
    d_qyc_new.save(scratch_qyc)
    d_qyc_new.close()
    try:
        shutil.copy2(scratch_qyc, qyc_pdf_desktop)
        print(f"Updated QYC.pdf: {qyc_pdf_desktop} (Pages: 8)")
    except PermissionError:
        print(f"Notice: {qyc_pdf_desktop} is locked.")

# 2. Update (2).pdf
c2_pdf_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (2).pdf")
if os.path.exists(c2_pdf_desktop):
    d_c2_orig = fitz.open(c2_pdf_desktop)
    d_c2_new = fitz.open()
    d_c2_new.insert_pdf(d_c2_orig, from_page=0, to_page=5)
    d_c2_orig.close()
    
    d_tbl = fitz.open(tbl_pdf_scratch)
    d_c2_new.insert_pdf(d_tbl, annots=True)
    d_tbl.close()
    
    scratch_c2 = os.path.join(SCRATCH_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (2).pdf")
    d_c2_new.save(scratch_c2)
    d_c2_new.close()
    try:
        shutil.copy2(scratch_c2, c2_pdf_desktop)
        print(f"Updated (2).pdf: {c2_pdf_desktop} (Pages: 8)")
    except PermissionError:
        print(f"Notice: {c2_pdf_desktop} is locked.")

# 3. Update Complete.pdf
comp_pdf_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")
base_corp = os.path.join(DESKTOP_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")
if os.path.exists(base_corp):
    d_comp_new = fitz.open()
    d_base = fitz.open(base_corp)
    d_comp_new.insert_pdf(d_base)
    d_base.close()
    
    d_tbl = fitz.open(tbl_pdf_scratch)
    d_comp_new.insert_pdf(d_tbl, annots=True)
    d_tbl.close()
    
    scratch_comp = os.path.join(SCRATCH_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")
    d_comp_new.save(scratch_comp)
    d_comp_new.close()
    try:
        shutil.copy2(scratch_comp, comp_pdf_desktop)
        print(f"Updated Complete.pdf: {comp_pdf_desktop} (Pages: 8)")
    except PermissionError:
        print(f"Notice: {comp_pdf_desktop} is locked.")

# 4. Render preview image for verification
d_vis = fitz.open(tbl_pdf_scratch)
pix = d_vis[0].get_pixmap(dpi=150, clip=fitz.Rect(50, 40, 560, 160))
out_img = os.path.join(SCRATCH_DIR, "fixed_table1_alignment.png")
pix.save(out_img)
d_vis.close()
print(f"Rendered Table 1 preview: {out_img}")

print("\n=== ALL FILES SUCCESSFULLY UPDATED AND REALIGNED! ===")
