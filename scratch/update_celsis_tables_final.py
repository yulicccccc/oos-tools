"""
Script: update_celsis_tables_final.py
Purpose: Update Celsis tables for OOS-262017 incorporating user edits and RS QA review standards.
"""

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

DOCS_DIR = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
DESKTOP_DIR = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

SRC_DOCX = os.path.join(SCRATCH_DIR, "Celsis table OOS-262017 (User Updated Backup).docx")
OUT_DOCX_SCRATCH = os.path.join(SCRATCH_DIR, "Celsis table OOS-262017.docx")
OUT_PDF_SCRATCH = os.path.join(SCRATCH_DIR, "Celsis table OOS-262017.pdf")
OUT_DOCX_DESKTOP = os.path.join(DESKTOP_DIR, "Celsis table OOS-262017.docx")
OUT_PDF_DESKTOP = os.path.join(DESKTOP_DIR, "Celsis table OOS-262017.pdf")

URL_SAMPLE = "https://etrax.eagleanalytical.com/Submission/Details/OIX27OGNKb64xLa70p0RRQ__"
URL_MICRO = "https://etrax.eagleanalytical.com/SubmissionTest/Details/OCvPHc7TodycYzvim2KtjQ__"
URL_0487 = "https://etrax.eagleanalytical.com/SubmissionTest/Details/fd3G2StZClcy1TP2ES6BLw__"

print("=== STEP 1: Load and Process Document ===")
doc = docx.Document(SRC_DOCX)

def set_cell_clean_text(cell, text, font_size_pt=7):
    """Sets clean centered text in cell with Times New Roman and 0 spacing."""
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    p_xml = (
        f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:pPr>'
        f'<w:jc w:val="center"/>'
        f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'<w:t>{text}</w:t>'
        f'</w:r>'
        f'</w:p>'
    )
    tc.append(parse_xml(p_xml))
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER

def add_hyperlink_to_cell(cell, url, text, font_size_pt=7):
    """Safely adds a hyperlink to a cell without altering original spacing."""
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    part = cell.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    p_xml = (
        f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:pPr>'
        f'<w:jc w:val="center"/>'
        f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:hyperlink r:id="{r_id}" w:history="1">'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:color w:val="0000FF"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'<w:u w:val="single"/>'
        f'</w:rPr>'
        f'<w:t>{text}</w:t>'
        f'</w:r>'
        f'</w:hyperlink>'
        f'</w:p>'
    )
    tc.append(parse_xml(p_xml))
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER

def set_dual_microbial_id_cell(cell, org1, org2, font_size_pt=6.5):
    """Sets two italicized bacterial organisms separated by an empty line, with 0 ampersands."""
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    p_xml = (
        f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:pPr>'
        f'<w:jc w:val="center"/>'
        f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:i/>'
        f'<w:iCs/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'<w:t>{org1}</w:t>'
        f'</w:r>'
        f'</w:p>'
    )
    p_empty_xml = (
        f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:pPr>'
        f'<w:jc w:val="center"/>'
        f'<w:spacing w:before="0" w:after="0" w:line="180" w:lineRule="auto"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'</w:p>'
    )
    p2_xml = (
        f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:pPr>'
        f'<w:jc w:val="center"/>'
        f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:i/>'
        f'<w:iCs/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'<w:t>{org2}</w:t>'
        f'</w:r>'
        f'</w:p>'
    )
    tc.append(parse_xml(p_xml))
    tc.append(parse_xml(p_empty_xml))
    tc.append(parse_xml(p2_xml))
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER

def set_table1_microbial_id_cell(cell, species, gram_stain, font_size_pt=7):
    """Sets single species in italics followed by gram stain non-italic."""
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    p_xml = (
        f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:pPr>'
        f'<w:jc w:val="center"/>'
        f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:i/>'
        f'<w:iCs/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'<w:t>{species}</w:t>'
        f'<w:br/>'
        f'</w:r>'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'<w:t>{gram_stain}</w:t>'
        f'</w:r>'
        f'</w:p>'
    )
    tc.append(parse_xml(p_xml))
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER

# ----------------- TABLE 1 -----------------
t1 = doc.tables[0]
set_cell_clean_text(t1.rows[1].cells[0], "ES")
set_cell_clean_text(t1.rows[1].cells[1], "America Alanis")
add_hyperlink_to_cell(t1.rows[1].cells[2], URL_SAMPLE, "ETX-260821-0259")
add_hyperlink_to_cell(t1.rows[1].cells[3], URL_MICRO, "ETX-260831-0608")
set_cell_clean_text(t1.rows[1].cells[4], "2 x 300mL TSB")
set_table1_microbial_id_cell(t1.rows[1].cells[5], "Terribacillus goriensis", "(Gram (+) rods)")

# ----------------- TABLE 2 -----------------
t2 = doc.tables[1]
# Daily rows: 2, 3, 4, 6, 7, 8, 9, 10, 11
for r_idx in [2, 3, 4, 6, 7, 8, 9, 10, 11]:
    row = t2.rows[r_idx]
    if r_idx in [2, 6, 9]:
        set_cell_clean_text(row.cells[3], "21Aug26")
        set_cell_clean_text(row.cells[4], "ES")
        set_cell_clean_text(row.cells[6], "Date Before Testing")
    elif r_idx in [3, 7, 10]:
        set_cell_clean_text(row.cells[3], "24Aug26")
        set_cell_clean_text(row.cells[4], "ES")
        set_cell_clean_text(row.cells[6], "Date of Testing")
    elif r_idx in [4, 8, 11]:
        set_cell_clean_text(row.cells[3], "25Aug26")
        set_cell_clean_text(row.cells[4], "ES")
        set_cell_clean_text(row.cells[6], "Date After Testing")
        
    set_cell_clean_text(row.cells[7], "No growth")
    set_cell_clean_text(row.cells[9], "N/A")
    set_cell_clean_text(row.cells[10], "N/A")
    set_cell_clean_text(row.cells[11], "None")

# Row 13: Weekly Active Air on 25Aug26 by SMO
set_cell_clean_text(t2.rows[13].cells[3], "25Aug26")
set_cell_clean_text(t2.rows[13].cells[4], "SMO")
set_cell_clean_text(t2.rows[13].cells[5], "Week of Testing")
set_cell_clean_text(t2.rows[13].cells[8], "4 CFU (ISO 8 114)")
set_cell_clean_text(t2.rows[13].cells[9], "ETX-260901-0112")
set_cell_clean_text(t2.rows[13].cells[10], "Pending ID")
set_cell_clean_text(t2.rows[13].cells[12], "None")

# Row 15: Weekly Surface on 25Aug26 by SMO
set_cell_clean_text(t2.rows[15].cells[3], "25Aug26")
set_cell_clean_text(t2.rows[15].cells[4], "SMO")
set_cell_clean_text(t2.rows[15].cells[5], "Week of Testing")
set_cell_clean_text(t2.rows[15].cells[8], "No growth")
set_cell_clean_text(t2.rows[15].cells[9], "N/A")
set_cell_clean_text(t2.rows[15].cells[10], "N/A")
set_cell_clean_text(t2.rows[15].cells[12], "None")

# ----------------- TABLE 3 -----------------
t3 = doc.tables[2]
# Daily rows: 2, 3, 4, 6, 7, 8, 9, 10, 11
for r_idx in [2, 3, 4, 6, 7, 8, 9, 10, 11]:
    row = t3.rows[r_idx]
    if r_idx in [2, 6, 9]:
        set_cell_clean_text(row.cells[3], "28Aug26")
        set_cell_clean_text(row.cells[4], "ALA")
        set_cell_clean_text(row.cells[6], "Date Before Testing")
    elif r_idx in [3, 7, 10]:
        set_cell_clean_text(row.cells[3], "31Aug26")
        set_cell_clean_text(row.cells[4], "ALA")
        set_cell_clean_text(row.cells[6], "Date of Testing")
    elif r_idx in [4, 8, 11]:
        set_cell_clean_text(row.cells[3], "01Sep26")
        set_cell_clean_text(row.cells[4], "ALA")
        set_cell_clean_text(row.cells[6], "Date After Testing")
        
    set_cell_clean_text(row.cells[7], "No growth")
    set_cell_clean_text(row.cells[9], "N/A")
    set_cell_clean_text(row.cells[10], "N/A")
    set_cell_clean_text(row.cells[11], "None")

# Row 13: Weekly Active Air on 04Sep26 by ISS
set_cell_clean_text(t3.rows[13].cells[3], "04Sep26")
set_cell_clean_text(t3.rows[13].cells[4], "ISS")
set_cell_clean_text(t3.rows[13].cells[5], "Week of Testing")
set_cell_clean_text(t3.rows[13].cells[8], "2 CFU (ISO 8 114)")
add_hyperlink_to_cell(t3.rows[13].cells[9], URL_0487, "ETX-260914-0487")
set_dual_microbial_id_cell(t3.rows[13].cells[10], "Corynebacterium ureicelerivorans", "Mycobacterium grossiae")
set_cell_clean_text(t3.rows[13].cells[12], "None")

# Row 15: Weekly Surface on 04Sep26 by ISS
set_cell_clean_text(t3.rows[15].cells[3], "04Sep26")
set_cell_clean_text(t3.rows[15].cells[4], "ISS")
set_cell_clean_text(t3.rows[15].cells[5], "Week of Testing")
set_cell_clean_text(t3.rows[15].cells[8], "No growth")
set_cell_clean_text(t3.rows[15].cells[9], "N/A")
set_cell_clean_text(t3.rows[15].cells[10], "N/A")
set_cell_clean_text(t3.rows[15].cells[12], "None")

# Ensure page break before Table 3
for p in doc.paragraphs:
    if "Table 3:" in p.text:
        p.paragraph_format.page_break_before = True

# Verification: Zero ampersands
amp_count = 0
for t_idx, t in enumerate(doc.tables):
    for r_idx, r in enumerate(t.rows):
        for c_idx, c in enumerate(r.cells):
            if '&' in c.text:
                print(f"ERROR: Table {t_idx+1} Row {r_idx} Cell {c_idx} has '&': {c.text}")
                amp_count += 1
assert amp_count == 0, f"Found {amp_count} ampersands in document!"
print(">>> ZERO AMPERSANDS CONFIRMED! <<<")

# Save DOCX
doc.save(OUT_DOCX_SCRATCH)
print(f"Saved scratch docx: {OUT_DOCX_SCRATCH}")
doc.save(OUT_DOCX_DESKTOP)
print(f"Saved desktop docx: {OUT_DOCX_DESKTOP}")

# Also overwrite 'Celsis table OOS-262017 (Updated).docx' on Desktop to keep it synced
doc.save(os.path.join(DESKTOP_DIR, "Celsis table OOS-262017 (Updated).docx"))
print(f"Synced updated docx on desktop: Celsis table OOS-262017 (Updated).docx")

print("=== STEP 2: Convert to PDF via Word COM ===")
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = False
try:
    wdoc = word.Documents.Open(OUT_DOCX_SCRATCH)
    wdoc.SaveAs2(OUT_PDF_SCRATCH, FileFormat=17)
    wdoc.Close()
    print(f"Exported scratch PDF: {OUT_PDF_SCRATCH}")
    shutil.copy2(OUT_PDF_SCRATCH, OUT_PDF_DESKTOP)
    print(f"Copied to desktop PDF: {OUT_PDF_DESKTOP}")
except Exception as e:
    print(f"Word COM error: {e}")
finally:
    word.Quit()

print("=== STEP 3: Verify and Render PDF Pages ===")
pdf = fitz.open(OUT_PDF_SCRATCH)
print(f"PDF Page count: {len(pdf)}")
for i, page in enumerate(pdf):
    pix = page.get_pixmap(dpi=200)
    img_path = os.path.join(SCRATCH_DIR, f"celsis_table_262017_final_p{i+1}.png")
    pix.save(img_path)
    print(f"Saved page {i+1} render to: {img_path}")
    links = page.get_links()
    print(f"Page {i+1} links ({len(links)}):")
    for lk in links:
        print(f"  URI: {lk.get('uri')}")
pdf.close()

print("=== ALL STEPS COMPLETED! ===")
