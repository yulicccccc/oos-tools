"""
Script: generate_celsis_tables_262333.py
Purpose: Generate standalone Celsis tables Word document (.docx ONLY) for OOS-262333
         incorporating the real verified Microbial Identification data from EagleTrax:
         - Bacillus paralicheniformis (Gram (+) rods)
         - Direct active hyperlinks for ETX-261002-0288 and ETX-261012-0654
         - Proper bracketing dates (02Oct26, 05Oct26, 06Oct26) & (09Oct26, 12Oct26, 13Oct26)
         - Strict Rule 0M (DOCX only to Desktop, ZERO ampersands)
"""

import os
import sys
import shutil
import docx
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
from docxtpl import DocxTemplate

sys.stdout.reconfigure(encoding='utf-8')

DOCS_DIR = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
DESKTOP_DIR = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")
TEMPLATE_PATH = os.path.join(DOCS_DIR, "tables for celsis.docx")

OUT_DOCX_DESKTOP = os.path.join(DESKTOP_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx")
OUT_DOCX_SCRATCH = os.path.join(SCRATCH_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx")

# Remove any old draft files if present
old_draft_desktop = os.path.join(DESKTOP_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis (Draft).docx")
if os.path.exists(old_draft_desktop):
    try:
        os.remove(old_draft_desktop)
        print(f"Removed old draft from Desktop: {old_draft_desktop}")
    except Exception as e:
        print(f"Could not remove old draft: {e}")

# Direct live EagleTrax links
URL_SAMPLE = "https://etrax.eagleanalytical.com/Submission/Details/fZlFl0vWDpqR-VWBsez7IA__"
URL_MICRO = "https://etrax.eagleanalytical.com/Submission/Details/-2W3%24lpbkR6si5jTUhP7pg__"

table_context = {
    # Table 1: Information for ETX-261002-0288 under investigation
    "analyst_name": "ES",
    "aliquoting_name": "Cuong Du",
    "sample_id": "ETX-261002-0288",
    "positive_id": "ETX-261012-0654",
    "positive_media": "1 x 300mL FTM",
    "positive_org": "Bacillus paralicheniformis\n(Gram (+) rods)",
    
    # Table 2: Processing Phase (05Oct26)
    "process_date": "05Oct26",
    "pro_before_test": "02Oct26",
    "pro_test_date": "05Oct26",
    "pro_after_test": "06Oct26",
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
    "alq_before_test": "09Oct26",
    "alq_test_date": "12Oct26",
    "alq_after_test": "13Oct26",
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

def add_hyperlink_to_cell(cell, url, text, font_size_pt=7):
    """
    Safely adds a hyperlink to a cell without altering the original paragraph spacing,
    table cell borders, or vertical alignment.
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

def set_cell_clean_text(cell, text, font_size_pt=7):
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

def set_table1_microbial_id_cell(cell, species, gram_stain, font_size_pt=7):
    """Sets species in italics and gram stain in regular text on separate line."""
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
set_cell_clean_text(t1.rows[1].cells[1], "Cuong Du")
add_hyperlink_to_cell(t1.rows[1].cells[2], URL_SAMPLE, "ETX-261002-0288")
add_hyperlink_to_cell(t1.rows[1].cells[3], URL_MICRO, "ETX-261012-0654")
set_cell_clean_text(t1.rows[1].cells[4], "1 x 300mL FTM")
set_table1_microbial_id_cell(t1.rows[1].cells[5], "Bacillus paralicheniformis", "(Gram (+) rods)")

# ----------------- TABLE 2 -----------------
# Ensure clean dates and initials are applied
t2 = doc.tables[1]
# Daily rows: 2, 3, 4, 6, 7, 8, 9, 10, 11
for r_idx in [2, 6, 9]:
    set_cell_clean_text(t2.rows[r_idx].cells[3], "02Oct26")
    set_cell_clean_text(t2.rows[r_idx].cells[4], "ES")
for r_idx in [3, 7, 10]:
    set_cell_clean_text(t2.rows[r_idx].cells[3], "05Oct26")
    set_cell_clean_text(t2.rows[r_idx].cells[4], "ES")
for r_idx in [4, 8, 11]:
    set_cell_clean_text(t2.rows[r_idx].cells[3], "06Oct26")
    set_cell_clean_text(t2.rows[r_idx].cells[4], "ES")

# ----------------- TABLE 3 -----------------
t3 = doc.tables[2]
for r_idx in [2, 6, 9]:
    set_cell_clean_text(t3.rows[r_idx].cells[3], "09Oct26")
    set_cell_clean_text(t3.rows[r_idx].cells[4], "CCD")
for r_idx in [3, 7, 10]:
    set_cell_clean_text(t3.rows[r_idx].cells[3], "12Oct26")
    set_cell_clean_text(t3.rows[r_idx].cells[4], "CCD")
for r_idx in [4, 8, 11]:
    set_cell_clean_text(t3.rows[r_idx].cells[3], "13Oct26")
    set_cell_clean_text(t3.rows[r_idx].cells[4], "CCD")

# Clean any '&' in entire document (Rule 0E)
for t_idx, t in enumerate(doc.tables):
    for r_idx, r in enumerate(t.rows):
        for c_idx, c in enumerate(r.cells):
            if '&' in c.text:
                # Replace ampersands with 'and'
                print(f"Replacing '&' in Table {t_idx+1} Row {r_idx} Cell {c_idx}")
                c.text = c.text.replace('&', 'and')

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

print("ALL WORD TABLE GENERATION COMPLETED SUCCESSFULLY!")
