# -*- coding: utf-8 -*-
"""
Script: scratch/rebuild_oos262333_perfect.py
Purpose: Rebuild OOS-262333 deliverables to 100% match the golden standard of
         'OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - RS Reviewed - Signed by - QYC.pdf'
         and all supervisor Robin Seymour & QA review requirements.
"""

import os
import sys
import shutil
import fitz
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import win32com.client
from pypdf import PdfWriter, PdfReader

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = r"scratch"

os.makedirs(SCRATCH_DIR, exist_ok=True)

# 1. Base Blank ZenQMS Form Template specifically for 262333
BLANK_TEMPLATE = os.path.join(SCRATCH_DIR, "qa_celsis_emails", "262333_CORP-FORM-21 Laboratory OOS Investigation Form (v11.1) (5).pdf")
if not os.path.exists(BLANK_TEMPLATE):
    BLANK_TEMPLATE = os.path.join(DOCS_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")

print("Using Blank Form Template:", BLANK_TEMPLATE)

# Deliverable File Names
PDF_NAME = "CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf"
DOCX_NAME = "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx"

DOCX_DESKTOP = os.path.join(DESKTOP_DIR, DOCX_NAME)
DOCX_DOCS = os.path.join(DOCS_DIR, DOCX_NAME)
PDF_DESKTOP = os.path.join(DESKTOP_DIR, PDF_NAME)
PDF_DOCS = os.path.join(DOCS_DIR, PDF_NAME)

SAMPLE_ID = "ETX-261002-0288"
SAMPLE_URL = "https://etrax.eagleanalytical.com/Submission/Details/Q57eK3N9WcR0Qh%24x6N5N6Q__"
SUBCULTURE_ID = "ETX-261012-0654"
ACTIVE_AIR_PLATE_ID = "ETX-261021-0578"
ACTIVE_AIR_PLATE_URL = "https://etrax.eagleanalytical.com/Submission/Details/SPTrZYPqECZmMHXaxXbl9Q__"

# ==============================================================================
# PART 1: GENERATE PERFECT STANDALONE WORD TABLES (2 PAGES)
# ==============================================================================
print("\n=== [PART 1] GENERATING STANDALONE WORD TABLES ===")
doc_word = docx.Document()

# Set standard 0.5 inch margins
sections = doc_word.sections
for s in sections:
    s.top_margin = Inches(0.4)
    s.bottom_margin = Inches(0.4)
    s.left_margin = Inches(0.5)
    s.right_margin = Inches(0.5)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=20, bottom=20, left=35, right=35):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_hyperlink(paragraph, text, url, font_size=Pt(7.0)):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}"><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="{int(font_size.pt * 2)}"/><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t>{text}</w:t></w:r></w:hyperlink>')
    paragraph._p.append(hyperlink)

# --- Table 1: Sample Information ---
p_title1 = doc_word.add_paragraph()
p_title1.paragraph_format.space_before = Pt(0)
p_title1.paragraph_format.space_after = Pt(2)
r_t1 = p_title1.add_run(f"Table 1: Information for {SAMPLE_ID} under investigation")
r_t1.bold = True
r_t1.font.name = 'Times New Roman'
r_t1.font.size = Pt(9.5)

t1 = doc_word.add_table(rows=2, cols=6)
t1.alignment = WD_TABLE_ALIGNMENT.CENTER
t1_headers = ['Processing Analyst', 'Aliquoting Analyst', 'Sample ID', 'Related Microbial ID', 'Media with microbial growth', 'Microbial ID']
t1_widths = [Inches(1.2), Inches(1.2), Inches(1.2), Inches(1.2), Inches(1.2), Inches(1.5)]

for j, h in enumerate(t1_headers):
    c = t1.rows[0].cells[j]
    c.width = t1_widths[j]
    set_cell_background(c, "F2F2F2")
    set_cell_margins(c, top=30, bottom=30, left=35, right=35)
    c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(h)
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(7.0)

t1_data = [
    'Alex Saravia',
    'Cuong Du',
    SAMPLE_ID,
    SUBCULTURE_ID,
    '1 x 300mL FTM',
    'Bacillus paralicheniformis'
]

for j, val in enumerate(t1_data):
    c = t1.rows[1].cells[j]
    c.width = t1_widths[j]
    set_cell_margins(c, top=25, bottom=25, left=35, right=35)
    c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    if j == 2:
        add_hyperlink(p, val, SAMPLE_URL, font_size=Pt(7.0))
    else:
        run = p.add_run(val)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(7.0)
        if j == 5:
            run.italic = True

p_space = doc_word.add_paragraph()
p_space.paragraph_format.space_before = Pt(3)
p_space.paragraph_format.space_after = Pt(0)

# Helper function to build 9-column EM Table matching OOS-262080
def build_em_table(doc, table_title, rows_data, weekly_air_data, is_table3=False):
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(2)
    r_t = p_title.add_run(table_title)
    r_t.bold = True
    r_t.font.name = 'Times New Roman'
    r_t.font.size = Pt(9.5)

    headers = [
        'Environmental\nMonitoring (EM)\nSampling Site',
        'Frequency',
        'Date\n(DDMMMYY)',
        'Analyst (Initials)',
        'Day\n/Week(s)',
        'Observation*',
        'Environmental\nMonitoring Plate\nETX ID',
        'Microbial ID',
        'Notes'
    ]
    col_widths = [Inches(1.5), Inches(0.65), Inches(0.75), Inches(0.6), Inches(0.85), Inches(0.9), Inches(0.9), Inches(1.15), Inches(0.4)]

    # 1 header + 1 banner + 3 rows + 1 banner + 6 rows + 1 banner + 1 row + 1 banner + 1 row = 16 rows
    tbl = doc.add_table(rows=16, cols=9)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Row 0: Headers
    for j, h in enumerate(headers):
        c = tbl.rows[0].cells[j]
        c.width = col_widths[j]
        set_cell_background(c, "F2F2F2")
        set_cell_margins(c, top=25, bottom=25, left=25, right=25)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(7.0)

    # Function to set banner with merged row matching OOS-262080
    def set_banner(row_idx, banner_text):
        row = tbl.rows[row_idx]
        merged = row.cells[0].merge(row.cells[-1])
        set_cell_background(merged, "D9D9D9")
        set_cell_margins(merged, top=15, bottom=15, left=25, right=25)
        merged.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = merged.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.text = banner_text
        p.runs[0].bold = True
        p.runs[0].font.name = 'Times New Roman'
        p.runs[0].font.size = Pt(7.0)

    # Banners matching OOS-262080
    set_banner(1, "Personnel EM Bracketing")
    set_banner(5, "Biological Safety Cabinet EM Bracketing Biological Safety Cabinet (BSC)")
    set_banner(12, "Weekly Active Air Sampling Bracketing")
    set_banner(14, "Surface Sampling of Anteroom and Cleanroom Bracketing")

    # Fill daily rows
    daily_row_indices = [2, 3, 4, 6, 7, 8, 9, 10, 11]
    for idx, r_idx in enumerate(daily_row_indices):
        data = rows_data[idx]
        row = tbl.rows[r_idx]
        for j, val in enumerate(data):
            c = row.cells[j]
            c.width = col_widths[j]
            set_cell_margins(c, top=15, bottom=15, left=25, right=25)
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(7.0)

    # Fill Weekly Active Air (Row 13)
    row_air = tbl.rows[13]
    for j, val in enumerate(weekly_air_data):
        c = row_air.cells[j]
        c.width = col_widths[j]
        set_cell_margins(c, top=15, bottom=15, left=25, right=25)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        if j == 6 and is_table3:
            add_hyperlink(p, val, ACTIVE_AIR_PLATE_URL, font_size=Pt(7.0))
        elif j == 7 and is_table3:
            # 7 pure species separated by blank lines, italic, no CFU, no pending
            species_list = [
                "Moraxella osloensis",
                "Sporisorium graminicola",
                "Staphylococcus haemolyticus",
                "Staphylococcus hominis",
                "Kocuria indica",
                "Ustilago maydis",
                "Kocuria rhizophila/ sp. BT304"
            ]
            for s_idx, sp in enumerate(species_list):
                r_sp = p.add_run(sp)
                r_sp.font.name = 'Times New Roman'
                r_sp.font.size = Pt(7.0)
                r_sp.italic = True
                if s_idx < len(species_list) - 1:
                    p.add_run("\n\n")
        else:
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(7.0)

    # Fill Weekly Surface (Row 15)
    surf_date = "05Oct26" if not is_table3 else "12Oct26"
    weekly_surf_data = [
        'Surface Sampling of Cleanrooms',
        'Weekly',
        surf_date,
        'SMO',
        'Week of Testing',
        'No Growth',
        'Not Applicable',
        'None',
        'None'
    ]
    row_surf = tbl.rows[15]
    for j, val in enumerate(weekly_surf_data):
        c = row_surf.cells[j]
        c.width = col_widths[j]
        set_cell_margins(c, top=15, bottom=15, left=25, right=25)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(val)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(7.0)

    # Footnote
    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(2)
    p_note.paragraph_format.space_after = Pt(0)
    r_fn = p_note.add_run("*Observation is based on CFU counts")
    r_fn.font.name = 'Times New Roman'
    r_fn.font.size = Pt(6.5)

# --- Table 2: Processing EM Data ---
t2_data_rows = [
    ['Personnel (Left Touch and Right Touch)', 'Daily', '02Oct26', 'ES', 'Date Before Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Personnel (Left Touch and Right Touch)', 'Daily', '05Oct26', 'ES', 'Date of Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Personnel (Left Touch and Right Touch)', 'Daily', '06Oct26', 'ES', 'Date After Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of ISO 5 (4 locations)', 'Daily', '02Oct26', 'ALA', 'Date Before Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of ISO 5 (4 locations)', 'Daily', '05Oct26', 'ES', 'Date of Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of ISO 5 (4 locations)', 'Daily', '06Oct26', 'ES', 'Date After Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 (2 locations)', 'Daily', '02Oct26', 'ALA', 'Date Before Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 (2 locations)', 'Daily', '05Oct26', 'ES', 'Date of Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 (2 locations)', 'Daily', '06Oct26', 'ES', 'Date After Testing', 'No Growth', 'Not Applicable', 'None', 'None']
]
t2_air = ['Active Air Sampling of Cleanrooms', 'Weekly', '05Oct26', 'SMO', 'Week of Testing', 'No Growth', 'Not Applicable', 'None', 'None']
build_em_table(doc_word, "Table 2: Environmental Monitoring from Processing Performed on 05Oct26", t2_data_rows, t2_air, is_table3=False)

# Explicit Page Break before Table 3
doc_word.add_page_break()

# --- Table 3: Aliquoting EM Data ---
t3_data_rows = [
    ['Personnel (Left Touch and Right Touch)', 'Daily', '09Oct26', 'CCD', 'Date Before Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Personnel (Left Touch and Right Touch)', 'Daily', '12Oct26', 'CCD', 'Date of Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Personnel (Left Touch and Right Touch)', 'Daily', '13Oct26', 'CCD', 'Date After Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of ISO 5 (4 locations)', 'Daily', '09Oct26', 'ALA', 'Date Before Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of ISO 5 (4 locations)', 'Daily', '12Oct26', 'CCD', 'Date of Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of ISO 5 (4 locations)', 'Daily', '13Oct26', 'CCD', 'Date After Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 (2 locations)', 'Daily', '09Oct26', 'ALA', 'Date Before Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 (2 locations)', 'Daily', '12Oct26', 'CCD', 'Date of Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 (2 locations)', 'Daily', '13Oct26', 'CCD', 'Date After Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None']
]
t3_air = ['Active Air Sampling of Cleanrooms', 'Weekly', '12Oct26', 'SMO', 'Week of Testing', '6 CFUs (ISO 8 114)', ACTIVE_AIR_PLATE_ID, '', 'None']
build_em_table(doc_word, "Table 3: Environmental Monitoring from Aliquoting Performed on 12Oct26", t3_data_rows, t3_air, is_table3=True)

# Save DOCX deliverables
doc_word.save(DOCX_DESKTOP)
doc_word.save(DOCX_DOCS)
print("Saved Standalone Word Tables to Desktop and Docs.")

# Export Standalone Tables PDF via Word COM
TABLES_TEMP_PDF = os.path.abspath(os.path.join(SCRATCH_DIR, "temp_perfect_tables_2p.pdf"))
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
wdoc = word.Documents.Open(os.path.abspath(DOCX_DESKTOP))
wdoc.SaveAs(TABLES_TEMP_PDF, FileFormat=17)
wdoc.Close()
word.Quit()
print("Exported Standalone Tables PDF via Word COM:", TABLES_TEMP_PDF)

p_tables = fitz.open(TABLES_TEMP_PDF)
print(f"VERIFIED STANDALONE TABLES PDF PAGE COUNT: {len(p_tables)} pages.")
assert len(p_tables) == 2, f"Expected strictly 2 pages for tables, got {len(p_tables)}"
p_tables.close()

# ==============================================================================
# PART 2: FILL OFFICIAL 6-PAGE ACROFORM MATCHING OOS-262080 EXACTLY
# ==============================================================================
print("\n=== [PART 2] FILLING OFFICIAL 6-PAGE ACROFORM ===")
doc_acro = fitz.open(BLANK_TEMPLATE)

# Page 3 Narrative (Part 1)
p_p3 = [
    "All analysts involved in the preparation (Mukyung Jang), processing (Alex Saravia), aliquoting (Cuong Du), and reading of the samples (Cuong Du) were interviewed comprehensively. Their responses are documented throughout this report. This investigation was performed as per MICRO-SOP-53 Sterility Test Out-of-Specification (OOS) Investigation Procedure.",
    "Upon arrival on 02Oct26, the sample was stored refrigerated in accordance with the client instructions. Both sample prepping analyst Mukyung Jang and processing analyst (Alex Saravia) inspected the sample during preparation and processing. No leaks or visible turbidity were observed, and no evidence of compromised container integrity was identified.",
    "All reagents and supplies mentioned in the material section above were stored according to supplier recommendations, and their integrity was visually verified before utilization. Moreover, all reagents and supplies had valid expiration dates. The functionality of all equipment was confirmed by reviewing data generated by our comprehensive in-house continuous monitoring system.",
    "The cleanroom suite used for processing procedures (CR114) comprises three interconnected sections: the innermost ISO 7 cleanroom (114B), which connects to the middle ISO 7 buffer room (114A), and then to the outermost ISO 8 anteroom (114). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from 114B through 114A and into 114.",
    "The innermost ISO 7 cleanroom 114B and its ISO 5 BSC E001316 were thoroughly cleaned and prepared before initiating testing by analyst Alex Saravia as per MICRO-SOP-2 (Environmental Monitoring of the Clean Room Facility) and MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology). For Celsis aliquoting, intermediate ISO 7 buffer room 114A and its ISO 5 BSC E001798 were thoroughly cleaned and prepared by analyst (Cuong Du) before initiating testing as per MICRO-SOP-44 and MICRO-SOP-9. Both BSCs, E001316 in 114B and E001798 in 114A, were certified and approved by the Engineering and Quality Assurance teams prior to use.",
    "On 05 Oct 2026, the sample vials for ETX-261002-0288 were received from the Sample Submissions team and brought into the Sterile Microbiology lab. Upon arrival, each vial was sprayed with an acidified bleach disinfectant, placed into pre-disinfected bins, and allowed a 10-minute contact time. Secondary disinfection was performed in the ISO 8 anteroom (114), where the vials were again treated with acidified bleach with a 10-minute contact time before transfer into cleanroom 114B. Inside 114B, the processing analyst (Alex Saravia) performed a final disinfection step with a 10-minute contact time. Once fully disinfected, the vials were transferred into the ISO 5 BSC E001316, placed on the disinfected work surface, aseptically opened, and tested in accordance with MICRO-SOP-44. Direct inoculation was performed by adding 15 mL of sample per media type (divided across ten 3 mL vials), using one 300 mL FTM jar and one 300 mL TSB jar. In accordance with method suitability study ETX-250826-0390, 3 mL of sterile Tween 80 was added into each 300 mL media bottle. Accordingly, a matching negative control jar with the same modifications was made. Following testing, the media bottles were transferred into incubators E001356 and E001357 to initiate incubation."
]

# Page 4 Narrative (Part 2)
p_p4 = [
    "Upon completion of incubation on 12 Oct 2026, the media bottles for ETX-261002-0288 were disinfected and transferred to the middle ISO 7 buffer room (114A) for the aliquoting step per MICRO-SOP-44 (Celsis Sterility Testing). In cleanroom 114A, the media bottles were disinfected one more time before transfer into the ISO 5 BSC E001798 located in cleanroom 114A. In ISO 5 BSC E001798, the sample was aliquoted into assay cuvettes by aliquoting analyst (Cuong Du). Following aliquoting, Celsis Sterility Reading was performed on instrument Celsis Advance 2 E002222 under workload 101226-2222-5.wkl in accordance with MICRO-SOP-44 by reading analyst (Cuong Du).",
    "Following the reading, sample ETX-261002-0288 was found to yield a positive reading in the 300 mL FTM media jar. The confirmed average Relative Luminescence Units (RLU) from the duplicate reading tubes, originating from the positive FTM sample bottle, yielded 3,647,568 RLU (duplicate tube readings of 3,608,783 RLU and 3,686,352 RLU, %CV 1.5%), which exceeded the positive cutoff of 2,596.5 RLU, where the FTM negative control was 866 RLU. The 300 mL TSB sample bottle tested negative at 776 RLU (cutoff 2,326.5 RLU). The %CV from the duplicate reading tubes for the positive FTM bottle was well within the specification (< 30%). Additionally, all Daily Controls, including the Instrument Blank (29 RLU, cutoff 100), Reagent Blank (103 RLU, cutoff 1,000), and Celsis ATP Positive Control (99,973 RLU, cutoff 50,000), were within the defined specifications, each with a %CV below 30%.",
    "Following the OOS result, the positive FTM media jar for ETX-261002-0288 was submitted for subculture and identification to isolate viable organisms, and to determine the potential origin of the contamination. Subculture of the positive FTM bottle under tracking ID ETX-261012-0654 yielded viable microbial growth, and confirmatory Oxford Nanopore sequencing identified the isolate as Bacillus paralicheniformis (Gram-positive spore-forming rod, 85.73% match). The positive reading and identification for the sample are documented in Table 1.",
    "All media used were within their expiration dates. Negative controls were prepared and incubated alongside the samples under the same conditions, and all negative controls passed. Monthly cleaning and disinfection of Cleanroom Suite 114 (including 114, 114A, 114B, and the BSCs E001316 and E001798) were performed on 27Sep26 by Sasha Allen (SAA), as per MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology). Decontamination was performed using HaloMist (Lot 2606A, Exp 06Jun28), and all chemical indicator strips (Lot 260620, Exp 20Jun28) passed, verifying that environmental sanitization remained validated and in control prior to testing. Attached Table 2 and Table 3 present the environmental monitoring results spanning the testing period. All environmental monitoring plates were incubated for no less than 48 hours at 30 to 35°C and no less than an additional 5 days at 20 to 25°C as per MICRO-SOP-2 (Environmental Monitoring of the Clean-room Facility). Table 2 pertains to Environmental Monitoring performed during Celsis Sterility Processing, and Table 3 pertains to Environmental Monitoring performed during Celsis Sterility Aliquoting.",
    "A review of the Environmental Monitoring (EM) results for the relevant Celsis Sterility processing period showed no microbial growth on personnel monitoring plates, settling plates, or ISO 5 BSC E001316 surface sampling plates on the date of testing (05Oct26), the preceding sampling date (02Oct26), or the subsequent sampling date (06Oct26). Weekly active air and surface sampling of the cleanroom suite associated with processing was performed on 05Oct26. No microbial recovery (No Growth) was observed from weekly active air or surface sampling of the cleanroom suite."
]

# Page 5 Narrative (Part 3)
p_p5 = [
    "After reviewing the Environmental Monitoring (EM) results for the relevant Celsis Sterility aliquoting period, no microbial growth was observed on personnel monitoring plates, ISO 5 BSC E001798 surface sampling plates, or settling plates on the date of aliquoting (12Oct26), the preceding sampling date (09Oct26), or the subsequent sampling date (13Oct26). Weekly surface sampling of the cleanroom suite performed on 12Oct26 showed no growth. Weekly active air sampling of the cleanroom suite performed on 12Oct26 showed recovery of 6 CFUs (ETX-261021-0578) in the outermost ISO 8 area (room 114), with zero recovery in the ISO 7 buffer or cleanrooms. Confirmatory Oxford Nanopore sequencing identified the recovered isolates as Moraxella osloensis, Sporisorium graminicola, Staphylococcus haemolyticus, Staphylococcus hominis, Kocuria indica, Ustilago maydis, and Kocuria rhizophila/ sp. BT304.",
    "Evaluation of Environmental Monitoring Results:\r \rIt is important to note that, for both processing and aliquoting activities, the samples are processed within validated ISO 5 BSCs (E001316 and E001798 respectively). Although microbial recovery was observed from weekly active air monitoring in the outermost ISO 8 room 114 background environment in cleanroom suite (CR114), these recoveries were not obtained from the ISO 5 processing environment or the ISO 7 cleanroom environment.",
    "The absence of microbial recovery from both the processing and the aliquoting analyst glove plates, ISO 5 BSC surface monitoring plates, and ISO 5 settling plates indicates that no identifiable contamination transfer pathway existed between the monitored background cleanroom environment and the ISO 5 critical processing areas. Samples were also prepared and transferred into disinfected, lidded containers and transported on carts within the controlled cleanroom environment. Additionally, none of the 7 environmental isolates recovered from the ISO 8 anteroom belong to the genus Bacillus or exhibit spore-forming rod morphology, demonstrating that they are distinctly different from the Bacillus paralicheniformis recovered from the positive product sample. Therefore, while the active air results demonstrate microbial recovery in the surrounding ISO 8 area, there is no evidence that the recovery originated from, or affected, the ISO 5 BSC environment where the test sample was handled and processed.",
    "All analysts confirmed full compliance with cleaning procedures as outlined in MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology) and MICRO-SOP-44 (Celsis Sterility Testing). A review of the available data confirms that the cleanroom and equipment conditions remained within acceptable parameters. No deviations or obvious signs of laboratory error during processing and aliquoting were noted from analysts.",
    "No other sample processed by analyst Alex Saravia on 05 Oct 2026 and aliquoted by analyst Cuong Du on 12 Oct 2026, failed Celsis Sterility testing that day. To assess the potential for sample-to-sample contamination contributing to the positive results, a comprehensive review was conducted of all samples processed in the same batch. All other samples processed by the same analyst on the day of testing were found to test negative for microbial growth. These findings suggest that cross-contamination between samples is unlikely.",
    'Analyzing a 6-month sample history for Optimal Balance Pharmacy indicates, this specific analyte "TESAMORELIN 8 MG/ML (3 ML) Injection" has had no prior failures using the Celsis Sterility testing during this period.',
    "Based on the laboratory investigation, no assignable laboratory cause was identified involving the analyst, instrument, reagents, supplies, or monitored laboratory environment. The initial Celsis OOS result therefore remains valid in accordance with the applicable laboratory OOS procedure."
]

narrative_p3 = "\r \r".join(p_p3)
narrative_p4 = "\r \r".join(p_p4)
narrative_p5 = "\r \r".join(p_p5)

# Page 1 Text Field3 block matching OOS-262080 formatting
analysts_block = (
    "Prepping Analyst:\rMukyung Jang (MJ)\r \r"
    "Processing Analyst:\rAlex Saravia (ES)\r \r"
    "Aliquoting Analyst:\rCuong Du (CCD)\r \r"
    "Reading Analyst:\rCuong Du (CCD)"
)

# Page 2 Materials matching OOS-262080 & Rev_OOS-262316 QA standards
tf22_materials = (
    "Materials and Reagents\r"
    "TSB (300mL): Lot: 08132026-1\r"
    "FTM (300mL): Lot: 08182026-6\r"
    "Tween (Polysorbate) 80: Lot: 0000743947\r \r"
    "Celsis AmpiScreen Reagents: 022601-1400\r"
    "Celsis Daily Wash & Rinse Kit: 012503-1142N\r"
    "Celsis ATP Positive Control: 022601-1483\r \r"
    "TSA and Surface Plates: Please refer to the Environmental Logs uploaded under the submission on EagleTrax\r \r"
    "Monthly Cleaning:\r"
    "HaloMist: Lot: 2606A\r"
    "H2O2 strips: Lot: 260620"
)

tf23_expirations = (
    "Materials and Reagents\r"
    "TSB: Exp: 11Nov2026\r"
    "FTM: Exp: 16Nov2026\r"
    "Tween 80: Exp: 31Jan2030\r \r"
    "Celsis AmpiScreen Reagents:\r"
    "Jan 2027\r"
    "Celsis Daily Wash & Rinse Kit:\r"
    "Jan 2027\r"
    "Celsis ATP Positive Control:\r"
    "Jan 2027\r \r"
    "TSA and Surface Plates: Please refer to the Environmental Logs uploaded under the submission on EagleTrax\r \r"
    "Monthly Cleaning:\r"
    "HaloMist: Exp: 06Jun28\r"
    "H2O2 strips: Exp: 20Jun28"
)

tf43_equipment = (
    "Incubator E001356 (Sensor E001450)\r \r"
    "Incubator E001357 (Sensor E001449)\r \r"
    "Incubator E001034 (Sensor E001501)\r \r"
    "Incubator E001031 (Sensor E001505)\r \r"
    "ISO 5 BSC E001316\r \r"
    "ISO 5 BSC E001798"
)

tf44_cal_due = (
    "Jan 2027\rFeb 2027\r \r"
    "Jan 2027\rFeb 2027\r \r"
    "Aug 2027\rFeb 2027\r \r"
    "Aug 2027\rFeb 2027\r \r"
    "Jun 2027\r \r"
    "Jun 2027"
)

# Comprehensive field mapping 100% aligned with OOS-262080
pdf_fields = {
    # Page 1
    'Text Field57': '262333',
    'Text Field0': 'Cuong Du (Written by Qiyue Chen)',
    'Date Field0': '05-Oct-2026',    # Testing / Inoculation Date
    'Date Field1': '12-Oct-2026',    # Read Date
    'Date Field2': '12-Oct-2026',    # Investigation Initiation Date
    'Date Field3': '12-Oct-2026',    # Manager Notification Date
    'Text Field1': 'Celsis Sterility Test',
    'Text Field2': SAMPLE_ID,        # Sample ID
    'Text Field3': analysts_block,
    'Text Field4': 'TESAMORELIN 8 MG/ML (3 ML) Injection', # Sample Description
    'Text Field5': 'Injection',       # Dosage Form
    'Text Field6': 'LG342011076',    # Lot Number
    'Text Field7': f'On 12 Oct 2026, sample {SAMPLE_ID} was found to be positive for growth after Celsis Sterility Testing.',
    'Text Field8': 'MICRO-SOP-44',
    'Text Field9': '03-Jan-2025',
    'Text Field10': '01',
    'Text Field11': 'Pass/Fail',
    'Text Field12': 'Kathan Parikh',
    'Text Field13': 'Yes, analysts Mukyung Jang, Alex Saravia, and Cuong Du were comprehensively interviewed.',
    'Text Field14': f'Yes, Sample ID {SAMPLE_ID}',
    'Text Field15': 'Yes, as per MICRO-SOP-44',
    'Text Field16': 'Yes, as per MICRO-SOP-44',
    'Text Field17': f'Yes, Information is available in EagleTrax under {SAMPLE_ID}',
    'Text Field18': 'Yes, all analysts are trained and qualified by the quality department to perform the test',
    'Text Field19': 'Not Applicable',
    'Text Field20': 'Not Applicable',
    'Text Field21': "Yes, the sample was stored refrigerated as per client's instructions",

    # Page 2
    'Text Field22': tf22_materials,
    'Text Field23': tf23_expirations,
    'Text Field24': 'Celsis ATP Positive Control',
    'Text Field25': '022601-1483',
    'Text Field26': '31 Jan 2027',
    'Text Field27': 'Not Applicable',
    'Text Field28': 'Not Applicable',
    'Text Field29': 'Not Applicable',
    'Text Field30': 'E002222',
    'Text Field31': 'Dec 2026',
    'Text Field32': 'E001736 (CR 114), BSC E001316\rE001736 (CR 114), BSC E001798',
    'Text Field33': 'Dec 2026, Jun 2027\rDec 2026, Jun 2027',
    'Text Field34': 'E002222',
    'Text Field35': 'Dec 2026',
    'Text Field36': 'Not Applicable',
    'Text Field37': 'Not Applicable',
    'Text Field38': 'Not Applicable',
    'Text Field39': 'Not Applicable',
    'Text Field40': 'See Phase I Summary',
    'Text Field41': 'See Phase I Summary',
    'Text Field42': 'See Phase I Summary',
    'Text Field43': tf43_equipment,
    'Text Field44': tf44_cal_due,
    'Text Field45': 'See Phase I Summary',

    # Page 3
    'Text Field46': 'Not Applicable',
    'Text Field47': 'Not Applicable',
    'Text Field48': 'N/A QYC 22Oct26',  # GDP sign-off
    'Text Field49': narrative_p3,

    # Page 4
    'Text Field50': narrative_p4,

    # Page 5
    'Text Field51': narrative_p5,

    # Page 6
    'Text Field52': 'Not Applicable',
    'Text Field53': 'Qiyue Chen',
    'Text Field54': 'Robin Seymour',
}

# Checkbox Whitelist 100% matched with OOS-262080
checkbox_values = {
    'Check Box0': 'Yes', 'Check Box1': 'Yes', 'Check Box2': 'Yes', 'Check Box3': 'Off',
    'Check Box4': 'Yes', 'Check Box5': 'Off', 'Check Box6': 'Off',
    'Check Box7': 'Yes', 'Check Box8': 'Off', 'Check Box9': 'Off',
    'Check Box10': 'Yes', 'Check Box11': 'Off', 'Check Box12': 'Off',
    'Check Box13': 'Yes', 'Check Box14': 'Off', 'Check Box15': 'Off',
    'Check Box16': 'Yes', 'Check Box17': 'Off', 'Check Box18': 'Off',
    'Check Box19': 'Yes', 'Check Box20': 'Off', 'Check Box21': 'Off',
    'Check Box22': 'Off', 'Check Box23': 'Off', 'Check Box24': 'Yes',
    'Check Box25': 'Off', 'Check Box26': 'Off', 'Check Box27': 'Yes',
    'Check Box28': 'Yes', 'Check Box29': 'Off', 'Check Box30': 'Off',
    'Check Box31': 'Off', 'Check Box32': 'Yes', 'Check Box33': 'Off',
    'Check Box34': 'Yes', 'Check Box35': 'Off', 'Check Box36': 'Off', # CB36 strictly OFF!
    'Check Box37': 'Off', 'Check Box38': 'Yes', 'Check Box39': 'Off',
    'Check Box40': 'Off', 'Check Box41': 'Off', 'Check Box42': 'Yes',
    'Check Box43': 'Yes', 'Check Box44': 'Off', 'Check Box45': 'Off',
    'Check Box46': 'Off', 'Check Box47': 'Off', 'Check Box48': 'Yes',
    'Check Box49': 'Off', 'Check Box50': 'Off', 'Check Box51': 'Yes',
    'Check Box52': 'Yes', 'Check Box53': 'Off', 'Check Box54': 'Off',
    'Check Box55': 'Yes', 'Check Box56': 'Off', 'Check Box57': 'Off',
    'Check Box58': 'Yes', 'Check Box59': 'Off', 'Check Box60': 'Off',
    'Check Box61': 'Off', 'Check Box62': 'Off', 'Check Box63': 'Yes',
    'Check Box64': 'Off', 'Check Box65': 'Off', 'Check Box66': 'Yes',
    'Check Box67': 'Yes', 'Check Box68': 'Off', 'Check Box69': 'Off', # CB69 strictly OFF!
    'Check Box70': 'Yes', 'Check Box71': 'Off', 'Check Box72': 'Off',
    'Check Box73': 'Yes', 'Check Box74': 'Off', 'Check Box75': 'Off',
    'Check Box76': 'Off', 'Check Box77': 'Off', 'Check Box78': 'Yes',
    'Check Box79': 'Yes', 'Check Box80': 'Off', 'Check Box81': 'Off',
    'Check Box82': 'Off', 'Check Box83': 'Off', 'Check Box84': 'Off', # CB84 strictly OFF!
    'Check Box85': 'Off', 'Check Box86': 'Off', 'Check Box87': 'Off',
    'Check Box88': 'Yes', 'Check Box89': 'Off', 'Check Box90': 'Off',
    'Check Box91': 'Off', 'Check Box92': 'Off',
}

# Typography Font Sizes
font_sizes = {
    'Text Field57': 8.5,
    'Date Field0': 8.5, 'Date Field1': 8.5, 'Date Field2': 8.5, 'Date Field3': 8.5,
    'Text Field0': 7.5, 'Text Field1': 8.5, 'Text Field2': 8.5, 'Text Field3': 6.0,
    'Text Field4': 8.5, 'Text Field5': 8.5, 'Text Field6': 8.5, 'Text Field7': 8.0,
    'Text Field8': 7.0, 'Text Field9': 7.0, 'Text Field10': 7.0, 'Text Field11': 8.5,
    'Text Field12': 8.5, 'Text Field13': 5.2, 'Text Field14': 8.5, 'Text Field15': 8.5,
    'Text Field16': 8.5, 'Text Field17': 6.5, 'Text Field18': 5.2, 'Text Field19': 8.5,
    'Text Field20': 8.5, 'Text Field21': 5.5,
    'Text Field22': 5.0, 'Text Field23': 5.0, 'Text Field24': 8.5, 'Text Field25': 8.5,
    'Text Field26': 8.5, 'Text Field27': 8.5, 'Text Field28': 8.5, 'Text Field29': 8.5,
    'Text Field30': 8.5, 'Text Field31': 8.5, 'Text Field32': 5.5, 'Text Field33': 5.5,
    'Text Field34': 8.5, 'Text Field35': 8.5, 'Text Field36': 8.5, 'Text Field37': 8.5,
    'Text Field38': 8.5, 'Text Field39': 8.5, 'Text Field40': 8.5, 'Text Field41': 8.5,
    'Text Field42': 8.5, 'Text Field43': 5.5, 'Text Field44': 5.5, 'Text Field45': 8.5,
    'Text Field46': 8.5, 'Text Field47': 8.5, 'Text Field48': 8.5, 'Text Field49': 7.4,
    'Text Field50': 7.4, 'Text Field51': 7.2,
    'Text Field52': 8.5, 'Text Field53': 8.5, 'Text Field54': 8.5,
}

for page_idx in range(len(doc_acro)):
    page = doc_acro[page_idx]
    for widget in page.widgets():
        name = widget.field_name
        if name in pdf_fields:
            widget.field_value = pdf_fields[name]
            widget.text_fontsize = font_sizes.get(name, 7.5)
            widget.text_font = "Helv"
            widget.update()
        elif name in checkbox_values:
            widget.field_value = checkbox_values[name]
            widget.update()
        elif name in font_sizes:
            widget.text_fontsize = font_sizes[name]
            widget.text_font = "Helv"
            widget.update()

TEMP_FILLED_ACRO = os.path.join(SCRATCH_DIR, "temp_perfect_filled_acro.pdf")
doc_acro.save(TEMP_FILLED_ACRO)
doc_acro.close()
print("Saved filled AcroForm to:", TEMP_FILLED_ACRO)

# Verify Pages 3, 4, 5 for text fit & zero overflow
doc_check = fitz.open(TEMP_FILLED_ACRO)
for pidx, fn in [(2, 'Text Field49'), (3, 'Text Field50'), (4, 'Text Field51')]:
    p = doc_check[pidx]
    w = [widget for widget in p.widgets() if widget.field_name == fn][0]
    words = [wd for wd in p.get_text('words') if wd[1] >= w.rect.y0 - 2 and wd[3] <= w.rect.y1 + 5]
    last_y = max([wd[3] for wd in words]) if words else 0
    rem = w.rect.y1 - last_y
    last_word = words[-1][4] if words else ''
    print(f"  Page {pidx+1} ({fn}): clearance={rem:.1f}pt, last word='{last_word}'")
    assert rem > 0, f"OVERFLOW DETECTED ON PAGE {pidx+1}!"
doc_check.close()

# ==============================================================================
# PART 3: COMBINE 6-PAGE ACROFORM WITH 2-PAGE TABLES INTO 8-PAGE MASTER DELIVERABLE
# ==============================================================================
print("\n=== [PART 3] COMBINING INTO 8-PAGE MASTER DELIVERABLE (VIA PYMUPDF) ===")

doc_master = fitz.open(TEMP_FILLED_ACRO)
assert len(doc_master) == 6

doc_tables = fitz.open(TABLES_TEMP_PDF)
assert len(doc_tables) == 2

doc_master.insert_pdf(doc_tables)
assert len(doc_master) == 8, f"Expected 8 pages master PDF, got {len(doc_master)}"

# Add Hyperlinks to Pages 7 & 8
p7 = doc_master[6] # Table 1 & Table 2
p8 = doc_master[7] # Table 3

rects_p7 = p7.search_for(SAMPLE_ID)
if rects_p7:
    print(f"Adding hyperlink on Page 7 for {SAMPLE_ID} at rect: {rects_p7[0]}")
    p7.insert_link({
        'kind': fitz.LINK_URI,
        'from': rects_p7[0],
        'uri': SAMPLE_URL
    })

rects_p8 = p8.search_for(ACTIVE_AIR_PLATE_ID)
if rects_p8:
    print(f"Adding hyperlink on Page 8 for {ACTIVE_AIR_PLATE_ID} at rect: {rects_p8[0]}")
    p8.insert_link({
        'kind': fitz.LINK_URI,
        'from': rects_p8[0],
        'uri': ACTIVE_AIR_PLATE_URL
    })

# Save to Desktop and Documents
doc_master.save(PDF_DESKTOP)
doc_master.save(PDF_DOCS)
doc_master.close()

print(f"\n==================================================================")
print(f"PERFECT MASTER DELIVERABLES DELIVERED!")
print(f"DOCX: {DOCX_DESKTOP}")
print(f"PDF:  {PDF_DESKTOP}")
print(f"Total Master PDF Pages: 8 PAGES (P1-6 AcroForm, P7 Table 1+2, P8 Table 3)")
print(f"==================================================================")
