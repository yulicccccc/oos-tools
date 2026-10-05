"""
Script: generate_celsis_tables_262017.py
Purpose: Generate standalone Celsis tables document (DOCX and PDF) for OOS-262017
         adhering strictly to cGMP, user rules, and zero-error standards.
"""

import os
import sys
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.opc.constants import RELATIONSHIP_TYPE
from docxtpl import DocxTemplate
import win32com.client
import fitz

sys.stdout.reconfigure(encoding='utf-8')

DOCS_DIR = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
DESKTOP_DIR = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")
TEMPLATE_PATH = os.path.join(DOCS_DIR, "tables for celsis.docx")

OUT_DOCX_DESKTOP = os.path.join(DESKTOP_DIR, "Celsis table OOS-262017.docx")
OUT_PDF_DESKTOP = os.path.join(DESKTOP_DIR, "Celsis table OOS-262017.pdf")
OUT_DOCX_SCRATCH = os.path.join(SCRATCH_DIR, "Celsis table OOS-262017.docx")
OUT_PDF_SCRATCH = os.path.join(SCRATCH_DIR, "Celsis table OOS-262017.pdf")

# Data
URL_SAMPLE = "https://etrax.eagleanalytical.com/Submission/Details/OIX27OGNKb64xLa70p0RRQ__"
URL_MICRO = "https://etrax.eagleanalytical.com/Submission/Details/G5-Kx58aPO5xsBX-06R1Dw__"

table_context = {
    # Table 1
    "analyst_name": "ES",
    "aliquoting_name": "America Alanis",
    "sample_id": "ETX-260821-0259",
    "positive_id": "ETX-260831-0608",
    "positive_media": "2 x 300mL TSB",
    "positive_org": "Pending (Gram stain ongoing)",
    
    # Table 2: Processing Phase (24Aug26)
    "process_date": "24Aug26",
    "pro_before_test": "21Aug26",
    "pro_test_date": "24Aug26",
    "pro_after_test": "25Aug26",
    "pro_analyst_initial": "ES",
    "pro_date_of_weekly": "25Aug26",
    
    "pro_be_obs_pers_dur_pro": "TBD", "pro_be_etx_pers_dur_pro": "TBD", "pro_be_id_pers_dur_pro": "TBD",
    "pro_obs_pers_dur_pro": "No growth", "pro_etx_pers_dur_pro": "N/A", "pro_id_pers_dur_pro": "N/A",
    "pro_af_obs_pers_dur_pro": "TBD", "pro_af_etx_pers_dur_pro": "TBD", "pro_af_id_pers_dur_pro": "TBD",
    
    "pro_be_obs_surf_dur_pro": "TBD", "pro_be_etx_surf_dur_pro": "TBD", "pro_be_id_surf_dur_pro": "TBD",
    "pro_obs_surf_dur_pro": "No growth", "pro_etx_surf_dur_pro": "N/A", "pro_id_surf_dur_pro": "N/A",
    "pro_af_obs_surf_dur_pro": "TBD", "pro_af_etx_surf_dur_pro": "TBD", "pro_af_id_surf_dur_pro": "TBD",
    
    "pro_be_obs_sett_dur_pro": "TBD", "pro_be_etx_sett_dur_pro": "TBD", "pro_be_id_sett_dur_pro": "TBD",
    "pro_obs_sett_dur_pro": "No growth", "pro_etx_sett_dur_pro": "N/A", "pro_id_sett_dur_pro": "N/A",
    "pro_af_obs_sett_dur_pro": "TBD", "pro_af_etx_sett_dur_pro": "TBD", "pro_af_id_sett_dur_pro": "TBD",
    
    "pro_obs_air_wk_of": "4 CFU (ISO 8 114)", "pro_etx_air_wk_of": "ETX-260901-0112", "pro_id_air_wk_of": "Pending ID",
    "pro_obs_room_wk_of": "No growth", "pro_etx_room_wk_of": "N/A", "pro_id_room_wk_of": "N/A",
    
    # Table 3: Aliquoting Phase (31Aug26)
    "alq_before_test": "28Aug26",
    "alq_test_date": "31Aug26",
    "alq_after_test": "01Sep26",
    "alq_analyst_initial": "ALA",
    "alq_date_of_weekly": "04Sep26",
    
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
    "alq_obs_room_wk_of": "No growth", "alq_etx_room_wk_of": "N/A",     "alq_id_room_wk_of": "N/A",
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
    table cell borders, or vertical alignment.
    """
    # Clear text in existing runs
    p = cell.paragraphs[0]
    p.text = ''
    # Ensure standard paragraph formatting matching original template
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    
    part = p.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    
    hyperlink_xml = (
        f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'r:id="{r_id}" w:history="1">'
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
    )
    p._p.append(parse_xml(hyperlink_xml))
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER

def set_cell_clean_text(cell, text):
    p = cell.paragraphs[0]
    p.text = text
    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for run in p.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(7)

# Apply hyperlinks to Table 1 (doc.tables[0])
t1 = doc.tables[0]
add_hyperlink_to_cell(t1.rows[1].cells[2], URL_SAMPLE, "ETX-260821-0259")
add_hyperlink_to_cell(t1.rows[1].cells[3], URL_MICRO, "ETX-260831-0608")

# Ensure Col 4 is strictly 2 x 300mL TSB
set_cell_clean_text(t1.rows[1].cells[4], "2 x 300mL TSB")

# Update Table 2 (doc.tables[1])
t2 = doc.tables[1]
# Row 13: Weekly Active Air on 25Aug26 by SMO
set_cell_clean_text(t2.rows[13].cells[3], "25Aug26")
set_cell_clean_text(t2.rows[13].cells[4], "SMO")
set_cell_clean_text(t2.rows[13].cells[8], "4 CFU (ISO 8 114)")
set_cell_clean_text(t2.rows[13].cells[9], "ETX-260901-0112")
set_cell_clean_text(t2.rows[13].cells[10], "Pending ID")
set_cell_clean_text(t2.rows[13].cells[11], "Pending ID")

# Row 15: Weekly Surface on 25Aug26 by SMO
set_cell_clean_text(t2.rows[15].cells[3], "25Aug26")
set_cell_clean_text(t2.rows[15].cells[4], "SMO")
set_cell_clean_text(t2.rows[15].cells[8], "No growth")
set_cell_clean_text(t2.rows[15].cells[9], "N/A")
set_cell_clean_text(t2.rows[15].cells[10], "N/A")
set_cell_clean_text(t2.rows[15].cells[11], "N/A")

# Update Table 3 (doc.tables[2])
t3 = doc.tables[2]
# Weekly Surface on 04Sep26 by ISS
set_cell_clean_text(t3.rows[15].cells[3], "04Sep26")
set_cell_clean_text(t3.rows[15].cells[4], "ISS")
set_cell_clean_text(t3.rows[15].cells[8], "No growth")
set_cell_clean_text(t3.rows[15].cells[9], "N/A")
set_cell_clean_text(t3.rows[15].cells[10], "N/A")
set_cell_clean_text(t3.rows[15].cells[11], "N/A")

# Weekly Active Air for week of 04Sep26 is Pending/TBD
set_cell_clean_text(t3.rows[13].cells[3], "04Sep26")
set_cell_clean_text(t3.rows[13].cells[4], "TBD")
set_cell_clean_text(t3.rows[13].cells[8], "TBD")
set_cell_clean_text(t3.rows[13].cells[9], "TBD")
set_cell_clean_text(t3.rows[13].cells[10], "TBD")
set_cell_clean_text(t3.rows[13].cells[11], "TBD")

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
    wdoc.SaveAs2(OUT_PDF_SCRATCH, FileFormat=17)
    wdoc.Close()
    print(f"Exported scratch PDF: {OUT_PDF_SCRATCH}")
    # Copy to Desktop
    import shutil
    shutil.copy2(OUT_PDF_SCRATCH, OUT_PDF_DESKTOP)
    print(f"Copied to desktop PDF: {OUT_PDF_DESKTOP}")
except Exception as e:
    print(f"Word COM error: {e}")
finally:
    word.Quit()

print("=== STEP 4: Render PDF Pages to PNG for Visual Verification ===")
pdf_doc = fitz.open(OUT_PDF_SCRATCH)
print(f"PDF Page count: {len(pdf_doc)}")
for i, page in enumerate(pdf_doc):
    pix = page.get_pixmap(dpi=200)
    png_path = os.path.join(SCRATCH_DIR, f"celsis_table_262017_p{i+1}.png")
    pix.save(png_path)
    print(f"Saved verification PNG: {png_path}")

print("ALL STEPS COMPLETED SUCCESSFULLY!")
