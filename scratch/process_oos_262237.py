"""
Script: process_oos_262237.py
Purpose: Complete generation of Standalone Tables (Word & PDF) and investigation pipeline
         for Scan RDI OOS-262237.
Client: GoGoMeds Select (E10747)
Sample ID: ETX-260921-0148
Sample Name: Semaglutide R 2.5mg/ml Benzyl Alcohol 0.9%
Lot: 20066
Test Date: 24Sep26
Processing Analyst: Guanchen Li (GL) (27th sample processed)
Processing BSC: BSC E001312 (Room 116B / Suite 116)
Changeover Analyst: Muralidhar Bythatagari (MRB)
Changeover BSC: BSC E001937 (L-Suite / CR144)
"""

import os
import sys
import shutil
from datetime import datetime
import fitz  # PyMuPDF
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"

OOS_ID = "262237"
CLIENT_NAME = "GoGoMeds Select (E10747)"
SAMPLE_ID = "ETX-260921-0148"
SAMPLE_URL = "https://etrax.eagleanalytical.com/Submission/Details/08VuIoFsBaXGehnOxJ3PMQ__"
SAMPLE_NAME = "Semaglutide R 2.5mg/ml Benzyl Alcohol 0.9%"
LOT_NUMBER = "20066"
DOSAGE_FORM = "Injectable"
TEST_DATE = "24Sep26"
TEST_DATE_LONG = "24-Sep-2026"

PROCESSING_ANALYST = "Guanchen Li"
PROCESSING_INITIAL = "GL"
PROCESSING_BSC = "1312"
CHANGEOVER_ANALYST = "Muralidhar Bythatagari"
CHANGEOVER_INITIAL = "MRB"
CHANGEOVER_BSC = "1937"

# Default placeholders to be updated upon user confirmation
READING_ANALYST = "[Pending]"
EVENT_COUNT = "[Pending]"
CONFIRMED_COUNT = "[Pending]"
MORPHOLOGY = "Curved rod-shaped morphology"

# -------------------------------------------------------------
# XML Helper Functions for Word Document
# -------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=50, bottom=50, left=60, right=60):
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

def add_hyperlink(paragraph, url, text, font_size=Pt(7), color_rgb="2980B9"):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w")} r:id="{r_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
    rPr.append(parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'))
    rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{int(font_size.pt * 2)}"/>'))
    rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="{color_rgb}"/>'))
    rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>'))
    new_run.append(rPr)
    t = parse_xml(f'<w:t {nsdecls("w")}>{text}</w:t>')
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def format_cell(cell, text, bold=False, italic=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex=None):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(cell)
    if fill_hex:
        set_cell_background(cell, fill_hex)
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

# -------------------------------------------------------------
# Build Standalone Tables Word Document
# -------------------------------------------------------------
def build_standalone_tables_doc(
    reading_analyst=READING_ANALYST,
    event_count=EVENT_COUNT,
    confirmed_count=CONFIRMED_COUNT,
    pers_obs="No Growth",
    surf_obs="No Growth",
    sett_obs="No Growth",
    air_date="25Sep26",
    air_obs="No Growth",
    air_etx="Not Applicable",
    air_id="Not Applicable",
    air_url=None,
    room_date="25Sep26",
    room_obs="No Growth",
    room_etx="Not Applicable",
    room_id="Not Applicable",
    room_url=None
):
    doc = docx.Document()
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # ---------------- Table 1 ----------------
    p_t1 = doc.add_paragraph()
    p_t1.paragraph_format.space_before = Pt(0)
    p_t1.paragraph_format.space_after = Pt(4)
    r_t1 = p_t1.add_run(f"Table 1: Information for {SAMPLE_ID} under investigation")
    r_t1.font.name = "Times New Roman"
    r_t1.font.size = Pt(9)
    r_t1.underline = True

    t1 = doc.add_table(rows=2, cols=6)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t1)

    t1_headers = [
        "Processing Analyst",
        "Reading Analyst",
        "Sample ID",
        "Events",
        "Confirmed Microbial\nEvents",
        "Morphology Description"
    ]
    t1_widths = [Inches(1.25), Inches(1.25), Inches(1.1), Inches(0.7), Inches(1.2), Inches(2.0)]

    for j, h in enumerate(t1_headers):
        cell = t1.cell(0, j)
        format_cell(cell, h, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

    format_cell(t1.cell(1, 0), PROCESSING_ANALYST, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(t1.cell(1, 1), reading_analyst, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    
    c_sid = t1.cell(1, 2)
    c_sid.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(c_sid)
    p_sid = c_sid.paragraphs[0]
    p_sid.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sid.paragraph_format.space_before = Pt(0)
    p_sid.paragraph_format.space_after = Pt(0)
    add_hyperlink(p_sid, SAMPLE_URL, SAMPLE_ID, font_size=Pt(7))

    format_cell(t1.cell(1, 3), str(event_count), bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(t1.cell(1, 4), str(confirmed_count), bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(t1.cell(1, 5), MORPHOLOGY, bold=True, italic=True, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

    for row in t1.rows:
        for j, w in enumerate(t1_widths):
            row.cells[j].width = w

    # Spacing between tables
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(12)
    p_space.paragraph_format.space_after = Pt(0)

    # ---------------- Table 2 ----------------
    p_t2 = doc.add_paragraph()
    p_t2.paragraph_format.space_before = Pt(0)
    p_t2.paragraph_format.space_after = Pt(4)
    r_t2 = p_t2.add_run(f"Table 2: Environmental Monitoring from Testing Performed on {TEST_DATE}")
    r_t2.font.name = "Times New Roman"
    r_t2.font.size = Pt(9)
    r_t2.underline = True

    t2 = doc.add_table(rows=0, cols=9)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t2)

    t2_widths = [
        Inches(1.05), # Sampling site
        Inches(0.65), # Frequency
        Inches(0.70), # Date (DDMMMYY)
        Inches(0.55), # Analyst
        Inches(1.15), # Day / Week(s)
        Inches(1.10), # Observation
        Inches(0.85), # EM Plate ETX Number
        Inches(0.85), # Microbial ID
        Inches(0.60)  # Notes
    ]

    hdr_row = t2.add_row()
    t2_headers = [
        "Environmental\nMonitoring (EM)\nSampling Site",
        "Frequency",
        "Date\n(DDMMMYY)",
        "Analyst",
        "Day /Week(s)",
        "Observation",
        "EM Plate ETX\nNumber",
        "Microbial ID",
        "Notes"
    ]
    for j, h in enumerate(t2_headers):
        format_cell(hdr_row.cells[j], h, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

    def add_section_divider(table, title_text):
        row = table.add_row()
        cell = row.cells[0]
        for c in row.cells[1:]:
            cell.merge(c)
        format_cell(cell, title_text, bold=True, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.LEFT, fill_hex="E8E8E8")
        return row

    def add_data_row(table, site, freq, date_str, analyst, day_week, obs, etx_num, micro_id, notes, etx_url=None):
        row = table.add_row()
        format_cell(row.cells[0], site, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(row.cells[1], freq, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(row.cells[2], date_str, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(row.cells[3], analyst, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(row.cells[4], day_week, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(row.cells[5], obs, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        
        if etx_url:
            c = row.cells[6]
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(c)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            add_hyperlink(p, etx_url, etx_num, font_size=Pt(7))
        else:
            format_cell(row.cells[6], etx_num, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
            
        format_cell(row.cells[7], micro_id, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(row.cells[8], notes, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        return row

    # --- Section 1: Personnel EM ---
    add_section_divider(t2, f"Personnel EM for {TEST_DATE}")
    add_data_row(t2, "Personal (Left Touch\nand Right Touch)", "Daily", TEST_DATE, PROCESSING_INITIAL, "Date of Testing\n(Scan)", pers_obs, "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, "Personal (Left Touch\nand Right Touch)", "Daily", TEST_DATE, CHANGEOVER_INITIAL, "Date of Testing\n(Scan C/O)", pers_obs, "Not Applicable", "Not Applicable", "None")

    # --- Section 2: BSC EM ---
    add_section_divider(t2, f"Biological Safety Cabinet (BSC) EM for BSC E00{PROCESSING_BSC} and BSC E00{CHANGEOVER_BSC} for {TEST_DATE}")
    add_data_row(t2, f"Surface Sampling of\nISO 5 BSC E00{PROCESSING_BSC}\n(4 Locations)", "Daily", TEST_DATE, PROCESSING_INITIAL, "Date of Testing\n(Scan)", surf_obs, "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, f"Surface Sampling of\nISO 5 BSC E00{CHANGEOVER_BSC}\n(4 Locations)", "Daily", TEST_DATE, CHANGEOVER_INITIAL, "Date of Testing\n(Scan C/O)", surf_obs, "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, f"Settling Sampling of\nISO 5 BSC E00{PROCESSING_BSC}", "Daily", TEST_DATE, PROCESSING_INITIAL, "Date of Testing\n(Scan)", sett_obs, "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, f"Settling Sampling of\nISO 5 BSC E00{CHANGEOVER_BSC}", "Daily", TEST_DATE, CHANGEOVER_INITIAL, "Date of Testing\n(Scan C/O)", sett_obs, "Not Applicable", "Not Applicable", "None")

    # --- Section 3: Weekly Active Air ---
    add_section_divider(t2, f"Weekly Active Air Sampling of Cleanroom Suite 116 (E001738) with Processing BSC for {TEST_DATE}")
    add_data_row(t2, "Active Air Sampling\nof Cleanrooms", "Weekly", air_date, "[TBD]", "Week of Testing", air_obs, air_etx, air_id, "None", etx_url=air_url)

    # --- Section 4: Weekly Surface ---
    add_section_divider(t2, f"Weekly Surface Sampling of Cleanroom Suite 116 (E001738) with Processing BSC for {TEST_DATE}")
    add_data_row(t2, "Surface Sampling of\nCleanrooms", "Weekly", room_date, "[TBD]", "Week of Testing", room_obs, room_etx, room_id, "None", etx_url=room_url)

    for row in t2.rows:
        if len(row.cells) == 9:
            for j, w in enumerate(t2_widths):
                row.cells[j].width = w

    return doc, t2

def export_docx_to_pdf(docx_path, pdf_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17) # 17 = wdFormatPDF
        doc.Close()
    finally:
        word.Quit()
    print(f"Exported to PDF: {pdf_path}")

if __name__ == "__main__":
    print("=== Generating Standalone Tables for OOS-262237 ===")
    tables_doc, _ = build_standalone_tables_doc()

    tables_docx_desktop = os.path.join(DESKTOP_DIR, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.docx")
    tables_pdf_desktop = os.path.join(DESKTOP_DIR, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
    tables_docx_docs = os.path.join(DOCUMENTS_DIR, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.docx")
    tables_pdf_docs = os.path.join(DOCUMENTS_DIR, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")

    tables_doc.save(tables_docx_docs)
    try:
        tables_doc.save(tables_docx_desktop)
        print(f"Saved Desktop DOCX: {tables_docx_desktop}")
    except Exception as e:
        print(f"Desktop DOCX lock note: {e}")

    temp_tbl_pdf = os.path.join(SCRIPT_DIR, "temp_tables_export_262237.pdf")
    export_docx_to_pdf(tables_docx_docs, temp_tbl_pdf)
    shutil.copy2(temp_tbl_pdf, tables_pdf_docs)
    try:
        shutil.copy2(temp_tbl_pdf, tables_pdf_desktop)
        print(f"Saved Desktop PDF: {tables_pdf_desktop}")
    except Exception as e:
        print(f"Desktop PDF lock note: {e}")

    print("\nTables OOS-262237 generated successfully in Word and PDF!")
