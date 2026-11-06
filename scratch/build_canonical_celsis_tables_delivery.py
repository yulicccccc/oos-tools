"""
scratch/build_canonical_celsis_tables_delivery.py
Build canonical OOS-262333 deliverables using the user's requested table structure:
- Table 1 (6 columns): Processing Analyst | Reading Analyst | Sample ID | Related Microbial ID | Media with microbial growth | Microbial ID
- Table 2 (9 columns): Standard EM grid with 4 #D9D9D9 section banners
- Table 3 (9 columns): Standard EM grid with 4 #D9D9D9 section banners, 7 Nanopore sequenced species in italic
"""

import os, sys, shutil, docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client
import fitz
from pypdf import PdfReader, PdfWriter

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = r"scratch"

BLANK_ACROFORM_PDF = r"scratch/qa_celsis_emails/262333_CORP-FORM-21 Laboratory OOS Investigation Form (v11.1) (5).pdf"

DOCX_NAME = "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx"
PDF_NAME = "CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf"

DOCX_DESKTOP = os.path.join(DESKTOP_DIR, DOCX_NAME)
DOCX_DOCS = os.path.join(DOCUMENTS_DIR, DOCX_NAME)
PDF_DESKTOP = os.path.join(DESKTOP_DIR, PDF_NAME)
PDF_DOCS = os.path.join(DOCUMENTS_DIR, PDF_NAME)

SAMPLE_ID = "ETX-261002-0288"
SAMPLE_URL = "https://etrax.eagleanalytical.com/Submission/Details/Q57eK3N9WcR0Qh%24x6N5N6Q__"
SUBCULTURE_ID = "ETX-261012-0654"
ACTIVE_AIR_PLATE_ID = "ETX-261021-0578"
ACTIVE_AIR_PLATE_URL = "https://etrax.eagleanalytical.com/SubmissionTest/Details/Cw%24QZhYAogGUwoutnOavpw__"

PRO_ANALYST = "Alex Saravia"
READ_ANALYST = "Cuong Du"
MEDIA = "1 x 300mL FTM"
MICROBIAL_ID = "Bacillus paralicheniformis"
MICROBIAL_STAIN = "Gram-positive rod"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=20, bottom=20, left=35, right=35):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def add_hyperlink(paragraph, url, text, font_size=Pt(7), color_rgb="0000FF"):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w")} r:id="{r_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
    rPr.append(parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'))
    rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{int(font_size.pt * 2)}"/>'))
    rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{int(font_size.pt * 2)}"/>'))
    rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="{color_rgb}"/>'))
    rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>'))
    new_run.append(rPr)
    t = parse_xml(f'<w:t {nsdecls("w")}>{text}</w:t>')
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def format_cell(cell, text, bold=False, italic=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER, v_align=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = v_align
    set_cell_margins(cell)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = font_size
    r.bold = bold
    r.italic = italic
    return r

# ==============================================================================
# BUILD DOCX TABLES
# ==============================================================================
print("\n=== [1] BUILDING CANONICAL WORD TABLES ===")
doc = docx.Document()
for section in doc.sections:
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

# --- Table 1: Information for Sample under investigation ---
p_t1 = doc.add_paragraph()
p_t1.paragraph_format.space_before = Pt(0)
p_t1.paragraph_format.space_after = Pt(2)
r_t1 = p_t1.add_run(f"Table 1: Information for {SAMPLE_ID} under investigation")
r_t1.font.name = "Times New Roman"
r_t1.font.size = Pt(8.5)
r_t1.underline = True

t1 = doc.add_table(rows=2, cols=6)
t1.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(t1)

t1_headers = [
    "Processing Analyst",
    "Reading Analyst",
    "Sample ID",
    "Related Microbial ID",
    "Media with microbial growth",
    "Microbial ID"
]
t1_widths = [Inches(1.10), Inches(1.10), Inches(1.25), Inches(1.10), Inches(1.50), Inches(1.45)]

for j, h in enumerate(t1_headers):
    format_cell(t1.cell(0, j), h, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

format_cell(t1.cell(1, 0), PRO_ANALYST, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
format_cell(t1.cell(1, 1), READ_ANALYST, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

# Sample ID Hyperlink
c_sid = t1.cell(1, 2)
c_sid.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_cell_margins(c_sid)
p_sid = c_sid.paragraphs[0]
p_sid.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sid.paragraph_format.space_before = Pt(0)
p_sid.paragraph_format.space_after = Pt(0)
add_hyperlink(p_sid, SAMPLE_URL, SAMPLE_ID, font_size=Pt(7))

format_cell(t1.cell(1, 3), SUBCULTURE_ID, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
format_cell(t1.cell(1, 4), MEDIA, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

# Microbial ID with Gram stain
c_mid = t1.cell(1, 5)
c_mid.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_cell_margins(c_mid)
p_mid = c_mid.paragraphs[0]
p_mid.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_mid.paragraph_format.space_before = Pt(0)
p_mid.paragraph_format.space_after = Pt(0)
p_mid.paragraph_format.line_spacing = 1.0
r_m1 = p_mid.add_run(MICROBIAL_ID)
r_m1.font.name = 'Times New Roman'
r_m1.font.size = Pt(7)
r_m1.italic = True
r_m_br = p_mid.add_run("\n")
r_m_br.font.name = 'Times New Roman'
r_m_br.font.size = Pt(7)
r_m2 = p_mid.add_run(f"({MICROBIAL_STAIN})")
r_m2.font.name = 'Times New Roman'
r_m2.font.size = Pt(7)
r_m2.italic = False

for row in t1.rows:
    for j, w in enumerate(t1_widths):
        row.cells[j].width = w

# Spacing between Table 1 and Table 2
p_space = doc.add_paragraph()
p_space.paragraph_format.space_before = Pt(6)
p_space.paragraph_format.space_after = Pt(0)

# --- Table 2: Environmental Monitoring from Testing Performed on 05Oct26 ---
p_t2 = doc.add_paragraph()
p_t2.paragraph_format.space_before = Pt(0)
p_t2.paragraph_format.space_after = Pt(2)
r_t2 = p_t2.add_run("Table 2: Environmental Monitoring from Testing Performed on 05Oct26")
r_t2.font.name = "Times New Roman"
r_t2.font.size = Pt(8.5)
r_t2.underline = True

t2 = doc.add_table(rows=1, cols=9)
t2.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(t2)

t2_headers = [
    "Environmental\nMonitoring (EM)\nSampling Site",
    "Frequency",
    "Date\n(DDMMMYY)",
    "Analyst",
    "Day /Week(s)",
    "Observation",
    "EM Plate ETX Number",
    "Microbial ID",
    "Notes"
]
t2_widths = [
    Inches(1.50), # Site
    Inches(0.55), # Freq
    Inches(0.60), # Date
    Inches(0.50), # Analyst
    Inches(0.85), # Day/Week
    Inches(0.85), # Obs
    Inches(0.85), # ETX
    Inches(1.35), # Micro ID
    Inches(0.45)  # Notes
]

for j, h in enumerate(t2_headers):
    format_cell(t2.cell(0, j), h, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

def add_section_divider(table, title):
    row = table.add_row()
    set_cell_background(row.cells[0], "D9D9D9")
    set_cell_margins(row.cells[0], top=18, bottom=18, left=35, right=35)
    p = row.cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(title)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(7)
    r.bold = True
    for k in range(1, 9):
        row.cells[k]._tc.get_or_add_tcPr().append(parse_xml(f'<w:hMerge {nsdecls("w")} w:val="continue"/>'))
    row.cells[0]._tc.get_or_add_tcPr().append(parse_xml(f'<w:hMerge {nsdecls("w")} w:val="restart"/>'))
    return row

def add_data_row(table, site, freq, date_str, analyst, day_week, obs, etx, micro_content_fn, notes="None", etx_url=None):
    row = table.add_row()
    format_cell(row.cells[0], site, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(row.cells[1], freq, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(row.cells[2], date_str, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(row.cells[3], analyst, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(row.cells[4], day_week, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(row.cells[5], obs, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    c_etx = row.cells[6]
    c_etx.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(c_etx)
    p_etx = c_etx.paragraphs[0]
    p_etx.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_etx.paragraph_format.space_before = Pt(0)
    p_etx.paragraph_format.space_after = Pt(0)
    if etx_url and etx != "Not Applicable":
        add_hyperlink(p_etx, etx_url, etx, font_size=Pt(7))
    else:
        r = p_etx.add_run(etx)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(7)
        
    c_micro = row.cells[7]
    c_micro.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(c_micro)
    p_micro = c_micro.paragraphs[0]
    p_micro.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_micro.paragraph_format.space_before = Pt(0)
    p_micro.paragraph_format.space_after = Pt(0)
    p_micro.paragraph_format.line_spacing = 1.0
    
    if callable(micro_content_fn):
        micro_content_fn(p_micro)
    else:
        r = p_micro.add_run(str(micro_content_fn))
        r.font.name = 'Times New Roman'
        r.font.size = Pt(7)
        
    format_cell(row.cells[8], notes, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    return row

# Table 2 Rows
add_section_divider(t2, "Personnel EM")
add_data_row(t2, "Personnel (Left Touch and Right Touch)", "Daily", "02Oct26", "ES", "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t2, "Personnel (Left Touch and Right Touch)", "Daily", "05Oct26", "ES", "Date of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t2, "Personnel (Left Touch and Right Touch)", "Daily", "06Oct26", "ES", "Date After Testing", "No Growth", "Not Applicable", "Not Applicable", "None")

add_section_divider(t2, "Biological Safety Cabinet EM for E001316")
add_data_row(t2, "Surface Sampling of ISO 5 (4 locations)", "Daily", "02Oct26", "ALA", "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t2, "Surface Sampling of ISO 5 (4 locations)", "Daily", "05Oct26", "ES", "Date of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t2, "Surface Sampling of ISO 5 (4 locations)", "Daily", "06Oct26", "ES", "Date After Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t2, "Settling Sampling of ISO 5 (2 locations)", "Daily", "02Oct26", "ALA", "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t2, "Settling Sampling of ISO 5 (2 locations)", "Daily", "05Oct26", "ES", "Date of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t2, "Settling Sampling of ISO 5 (2 locations)", "Daily", "06Oct26", "ES", "Date After Testing", "No Growth", "Not Applicable", "Not Applicable", "None")

add_section_divider(t2, "Weekly Active Air Sampling of Cleanroom Suite 114 (E001737) with Processing BSC")
add_data_row(t2, "Active Air Sampling of Cleanrooms", "Weekly", "05Oct26", "SMO", "Week of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")

add_section_divider(t2, "Weekly Surface Sampling of Cleanroom Suite 114 (E001737) with Processing BSC")
add_data_row(t2, "Surface Sampling of Cleanrooms", "Weekly", "05Oct26", "SMO", "Week of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")

for row in t2.rows:
    for j, w in enumerate(t2_widths):
        row.cells[j].width = w

# --- Table 3: Environmental Monitoring from Aliquoting Performed on 12Oct26 ---
# Explicit pageBreakBefore puts Table 3 on Page 2
p_t3 = doc.add_paragraph()
pPr3 = p_t3._p.get_or_add_pPr()
pPr3.append(parse_xml(f'<w:pageBreakBefore {nsdecls("w")}/>'))
p_t3.paragraph_format.space_before = Pt(0)
p_t3.paragraph_format.space_after = Pt(2)
r_t3 = p_t3.add_run("Table 3: Environmental Monitoring from Aliquoting Performed on 12Oct26")
r_t3.font.name = "Times New Roman"
r_t3.font.size = Pt(8.5)
r_t3.underline = True

t3 = doc.add_table(rows=1, cols=9)
t3.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(t3)

for j, h in enumerate(t2_headers):
    format_cell(t3.cell(0, j), h, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

add_section_divider(t3, "Personnel EM")
add_data_row(t3, "Personnel (Left Touch and Right Touch)", "Daily", "09Oct26", "CCD", "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t3, "Personnel (Left Touch and Right Touch)", "Daily", "12Oct26", "CCD", "Date of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t3, "Personnel (Left Touch and Right Touch)", "Daily", "13Oct26", "CCD", "Date After Testing", "No Growth", "Not Applicable", "Not Applicable", "None")

add_section_divider(t3, "Biological Safety Cabinet EM for E001798")
add_data_row(t3, "Surface Sampling of ISO 5 (4 locations)", "Daily", "09Oct26", "ALA", "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t3, "Surface Sampling of ISO 5 (4 locations)", "Daily", "12Oct26", "CCD", "Date of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t3, "Surface Sampling of ISO 5 (4 locations)", "Daily", "13Oct26", "CCD", "Date After Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t3, "Settling Sampling of ISO 5 (2 locations)", "Daily", "09Oct26", "ALA", "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t3, "Settling Sampling of ISO 5 (2 locations)", "Daily", "12Oct26", "CCD", "Date of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
add_data_row(t3, "Settling Sampling of ISO 5 (2 locations)", "Daily", "13Oct26", "CCD", "Date After Testing", "No Growth", "Not Applicable", "Not Applicable", "None")

add_section_divider(t3, "Weekly Active Air Sampling of Cleanroom Suite 114 (E001737) with Processing BSC")

def fill_active_air_7species(p):
    species = [
        "Moraxella osloensis",
        "Sporisorium graminicola",
        "Staphylococcus haemolyticus",
        "Staphylococcus hominis",
        "Kocuria indica",
        "Ustilago maydis",
        "Kocuria rhizophila"
    ]
    for idx, sp in enumerate(species):
        r = p.add_run(sp)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(6.5)
        r.italic = True
        if idx < len(species) - 1:
            p.add_run("\n\n")

add_data_row(t3, "Active Air Sampling of Cleanrooms", "Weekly", "12Oct26", "SMO", "Week of Testing", "6 CFUs (ISO 8 114)", ACTIVE_AIR_PLATE_ID, fill_active_air_7species, "None", etx_url=ACTIVE_AIR_PLATE_URL)

add_section_divider(t3, "Weekly Surface Sampling of Cleanroom Suite 114 (E001737) with Processing BSC")
add_data_row(t3, "Surface Sampling of Cleanrooms", "Weekly", "12Oct26", "SMO", "Week of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")

for row in t3.rows:
    for j, w in enumerate(t2_widths):
        row.cells[j].width = w

# Save DOCX deliverables
doc.save(DOCX_DESKTOP)
doc.save(DOCX_DOCS)
print(f"Saved DOCX deliverable: {DOCX_DESKTOP}")
print(f"Saved DOCX deliverable: {DOCX_DOCS}")

# Export Standalone Tables PDF via Word
TABLES_TEMP_PDF = os.path.abspath(os.path.join(SCRATCH_DIR, "temp_tables_canonical_usp.pdf"))
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
wdoc = word.Documents.Open(os.path.abspath(DOCX_DESKTOP))
wdoc.SaveAs(TABLES_TEMP_PDF, FileFormat=17)
wdoc.Close()
word.Quit()
print(f"Exported Standalone Tables PDF: {TABLES_TEMP_PDF}")

p_tables = fitz.open(TABLES_TEMP_PDF)
print(f"VERIFIED STANDALONE TABLES PDF PAGE COUNT: {len(p_tables)} pages.")
assert len(p_tables) == 2, f"Expected 2 pages tables, got {len(p_tables)}"

# ==============================================================================
# COMBINE WITH 6-PAGE ACROFORM INTO MASTER DELIVERABLE PDF
# ==============================================================================
print("\n=== [2] COMBINING WITH 6-PAGE ACROFORM ===")
FILLED_ACRO_6P = os.path.join(SCRATCH_DIR, "temp_filled_acro_6p.pdf")
assert os.path.exists(FILLED_ACRO_6P), f"Missing {FILLED_ACRO_6P}"

writer = PdfWriter()
reader_form = PdfReader(FILLED_ACRO_6P)
assert len(reader_form.pages) == 6
for p in reader_form.pages:
    writer.add_page(p)

reader_tables = PdfReader(TABLES_TEMP_PDF)
assert len(reader_tables.pages) == 2
for p in reader_tables.pages:
    writer.add_page(p)

PRE_LINK_PDF = os.path.join(SCRATCH_DIR, "temp_canonical_pre_link.pdf")
with open(PRE_LINK_PDF, "wb") as f_out:
    writer.write(f_out)

# Add hyperlinks on Page 7 (Table 1) and Page 8 (Table 3)
doc_master = fitz.open(PRE_LINK_PDF)
assert len(doc_master) == 8, f"Expected 8 pages master PDF, got {len(doc_master)}"

p7 = doc_master[6] # Page 7: Table 1 & Table 2
p8 = doc_master[7] # Page 8: Table 3

# Page 7 Hyperlink for ETX-261002-0288
rects_p7 = p7.search_for("ETX-261002-0288")
if rects_p7:
    print(f"Adding hyperlink on Page 7 for ETX-261002-0288 at rect: {rects_p7[0]}")
    p7.insert_link({
        'kind': fitz.LINK_URI,
        'from': rects_p7[0],
        'uri': SAMPLE_URL
    })

# Page 8 Hyperlink for ETX-261021-0578
rects_p8 = p8.search_for("ETX-261021-0578")
if rects_p8:
    print(f"Adding hyperlink on Page 8 for ETX-261021-0578 at rect: {rects_p8[0]}")
    p8.insert_link({
        'kind': fitz.LINK_URI,
        'from': rects_p8[0],
        'uri': ACTIVE_AIR_PLATE_URL
    })

# Save Master PDF deliverables
doc_master.save(PDF_DESKTOP)
doc_master.save(PDF_DOCS)
doc_master.close()

print(f"\n==================================================================")
print(f"CANONICAL DELIVERABLES COMPLETE AND SYNCHRONIZED!")
print(f"DOCX Desktop: {DOCX_DESKTOP}")
print(f"PDF Desktop:  {PDF_DESKTOP}")
print(f"Total Master PDF Pages: 8 PAGES (P1-6 AcroForm, P7 Table 1+2, P8 Table 3)")
print(f"==================================================================")
