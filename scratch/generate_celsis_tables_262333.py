"""
Script: generate_celsis_tables_262333.py
Purpose: Generate draft standalone Celsis tables document (DOCX and PDF) for OOS-262333
         using currently available information, adhering strictly to Robin Seymour QA standards,
         Rule 50, Section 36, and Table 1 interactive hyperlink requirements.
"""

import os
import sys
import shutil
import docx
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
from docxtpl import DocxTemplate
import win32com.client
import fitz

sys.stdout.reconfigure(encoding='utf-8')

DOCS_DIR = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
DESKTOP_DIR = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")
TEMPLATE_PATH = os.path.join(DOCS_DIR, "tables for celsis.docx")

OUT_DOCX_DESKTOP = os.path.join(DESKTOP_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis (Draft).docx")
OUT_DOCX_SCRATCH = os.path.join(SCRATCH_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis (Draft).docx")

# Direct live EagleTrax links (unwrapped from Outlook SafeLinks)
URL_SAMPLE = "https://etrax.eagleanalytical.com/Submission/Details/fZlFl0vWDpqR-VWBsez7IA__"
URL_MICRO = "https://etrax.eagleanalytical.com/Submission/Details/-2W3%24lpbkR6si5jTUhP7pg__"

table_context = {
    # Table 1: Information for ETX-261002-0288 under investigation
    "analyst_name": "ES",
    "aliquoting_name": "Cuong Du",
    "sample_id": "ETX-261002-0288",
    "positive_id": "ETX-261012-0654",
    "positive_media": "1 x 300mL FTM",
    "positive_org": "TBD\n(Pending Identification)",
    
    # Table 2: Processing Phase (05Oct26)
    "process_date": "05Oct26",
    "pro_before_test": "TBD",
    "pro_test_date": "05Oct26",
    "pro_after_test": "TBD",
    "pro_analyst_initial": "ES",
    "pro_date_of_weekly": "TBD",
    
    "pro_be_obs_pers_dur_pro": "TBD", "pro_be_etx_pers_dur_pro": "TBD", "pro_be_id_pers_dur_pro": "TBD",
    "pro_obs_pers_dur_pro": "TBD",    "pro_etx_pers_dur_pro": "TBD",    "pro_id_pers_dur_pro": "TBD",
    "pro_af_obs_pers_dur_pro": "TBD", "pro_af_etx_pers_dur_pro": "TBD", "pro_af_id_pers_dur_pro": "TBD",
    
    "pro_be_obs_surf_dur_pro": "TBD", "pro_be_etx_surf_dur_pro": "TBD", "pro_be_id_surf_dur_pro": "TBD",
    "pro_obs_surf_dur_pro": "TBD",    "pro_etx_surf_dur_pro": "TBD",    "pro_id_surf_dur_pro": "TBD",
    "pro_af_obs_surf_dur_pro": "TBD", "pro_af_etx_surf_dur_pro": "TBD", "pro_af_id_surf_dur_pro": "TBD",
    
    "pro_be_obs_sett_dur_pro": "TBD", "pro_be_etx_sett_dur_pro": "TBD", "pro_be_id_sett_dur_pro": "TBD",
    "pro_obs_sett_dur_pro": "TBD",    "pro_etx_sett_dur_pro": "TBD",    "pro_id_sett_dur_pro": "TBD",
    "pro_af_obs_sett_dur_pro": "TBD", "pro_af_etx_sett_dur_pro": "TBD", "pro_af_id_sett_dur_pro": "TBD",
    
    "pro_obs_air_wk_of": "TBD",       "pro_etx_air_wk_of": "TBD",       "pro_id_air_wk_of": "TBD",
    "pro_obs_room_wk_of": "TBD",      "pro_etx_room_wk_of": "TBD",      "pro_id_room_wk_of": "TBD",
    
    # Table 3: Aliquoting Phase (12Oct26)
    "alq_before_test": "TBD",
    "alq_test_date": "12Oct26",
    "alq_after_test": "TBD",
    "alq_analyst_initial": "CCD",
    "alq_date_of_weekly": "TBD",
    
    "alq_be_obs_pers_dur_pro": "TBD", "alq_be_etx_pers_dur_pro": "TBD", "alq_be_id_pers_dur_pro": "TBD",
    "alq_obs_pers_dur_pro": "TBD",    "alq_etx_pers_dur_pro": "TBD",    "alq_id_pers_dur_pro": "TBD",
    "alq_af_obs_pers_dur_pro": "TBD", "alq_af_etx_pers_dur_pro": "TBD", "alq_af_id_pers_dur_pro": "TBD",
    
    "alq_be_obs_surf_dur_pro": "TBD", "alq_be_etx_surf_dur_pro": "TBD", "alq_be_id_surf_dur_pro": "TBD",
    "alq_obs_surf_dur_pro": "TBD",    "alq_etx_surf_dur_pro": "TBD",    "alq_id_surf_dur_pro": "TBD",
    "alq_af_obs_surf_dur_pro": "TBD", "alq_af_etx_surf_dur_pro": "TBD", "alq_af_id_surf_dur_pro": "TBD",
    
    "alq_be_obs_sett_dur_pro": "TBD", "alq_be_etx_sett_dur_pro": "TBD", "alq_be_id_sett_dur_pro": "TBD",
    "alq_obs_sett_dur_pro": "TBD",    "alq_etx_sett_dur_pro": "TBD",    "alq_id_sett_dur_pro": "TBD",
    "alq_af_obs_sett_dur_pro": "TBD", "alq_af_etx_sett_dur_pro": "TBD", "alq_af_id_sett_dur_pro": "TBD",
    
    "alq_obs_air_wk_of": "TBD",       "alq_etx_air_wk_of": "TBD",       "alq_id_air_wk_of": "TBD",
    "alq_obs_room_wk_of": "TBD",      "alq_etx_room_wk_of": "TBD",      "alq_id_room_wk_of": "TBD",
}

print("=== STEP 1: Render Jinja Template ===")
tpl = DocxTemplate(TEMPLATE_PATH)
tpl.render(table_context)
tpl.save(OUT_DOCX_SCRATCH)

print("=== STEP 2: Post-Processing Word Document ===")
doc = docx.Document(OUT_DOCX_SCRATCH)

# 1. Page break before Table 3
for p in doc.paragraphs:
    if "Table 3:" in p.text:
        p.paragraph_format.page_break_before = True

def add_hyperlink_to_cell(cell, url, text):
    """
    Safely adds a hyperlink to a cell without altering the original paragraph spacing,
    table cell borders, or vertical alignment. Avoids w:rStyle to eliminate line-spacing traps.
    """
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
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:sz w:val="14"/>'
        f'<w:szCs w:val="14"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:hyperlink r:id="{r_id}" w:history="1">'
        f'<w:r>'
        f'<w:rPr>'
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
    )
    tc.append(parse_xml(p_xml))
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER

def set_cell_clean_text(cell, text):
    p = cell.paragraphs[0]
    p.text = text
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for run in p.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(7)

# ----------------- TABLE 1 -----------------
t1 = doc.tables[0]
set_cell_clean_text(t1.rows[1].cells[0], "ES")
set_cell_clean_text(t1.rows[1].cells[1], "Cuong Du")
add_hyperlink_to_cell(t1.rows[1].cells[2], URL_SAMPLE, "ETX-261002-0288")
add_hyperlink_to_cell(t1.rows[1].cells[3], URL_MICRO, "ETX-261012-0654")
set_cell_clean_text(t1.rows[1].cells[4], "1 x 300mL FTM")
set_cell_clean_text(t1.rows[1].cells[5], "TBD (Pending ID)")

# Verify zero ampersands in entire document
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
try:
    doc.save(OUT_DOCX_DESKTOP)
    print(f"Saved desktop docx: {OUT_DOCX_DESKTOP}")
except Exception as e:
    print(f"Error saving to desktop: {e}")

print("=== STEP 3: Convert to PDF via Word COM ===")
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = False
try:
    wdoc = word.Documents.Open(OUT_DOCX_SCRATCH)
    OUT_PDF_SCRATCH = os.path.join(SCRATCH_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis (Draft).pdf")
    wdoc.SaveAs2(OUT_PDF_SCRATCH, FileFormat=17)
    wdoc.Close()
    print(f"Exported scratch PDF for preview: {OUT_PDF_SCRATCH}")
except Exception as e:
    print(f"Word COM error: {e}")
finally:
    word.Quit()

print("=== STEP 4: Render PDF Pages to PNG for Visual Verification ===")
pdf_doc = fitz.open(OUT_PDF_SCRATCH)
print(f"PDF Page count: {len(pdf_doc)}")
for i, page in enumerate(pdf_doc):
    pix = page.get_pixmap(dpi=200)
    png_path = os.path.join(SCRATCH_DIR, f"celsis_table_262333_p{i+1}.png")
    pix.save(png_path)
    print(f"Saved verification PNG: {png_path}")

# Verify links in PDF
for i, p in enumerate(pdf_doc):
    links = p.get_links()
    print(f"Page {i+1} link count: {len(links)}")
    for lk in links:
        print(f"  Page {i+1} link: {lk.get('uri')}")
pdf_doc.close()

print("ALL DRAFT TABLE STEPS COMPLETED SUCCESSFULLY!")
