import os
import sys
import shutil
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import win32com.client
import fitz

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"

OOS_ID = "262329"
CLIENT_NAME = "GoGoMeds Select (E10747)"
TEST_NAME = "USP71"
SAMPLE_ID = "ETX-260624-0251 (BUD-T90)"
SAMPLE_URL = "https://etrax.eagleanalytical.com/Submission/Details/nIsCJb3y7xy7m1FK8luTrg__"
SUBCULTURE_ID = "ETX-261012-0367"
SUBCULTURE_URL = "https://etrax.eagleanalytical.com/Submission/Details/e6Jo3DyLpUEd3QnYm5Uttw__"
EM_PLATE_ID = "ETX-261001-0625"
EM_PLATE_URL = "https://etrax.eagleanalytical.com/Submission/Details/yp3RLi-9BGnGhzxnYnqk4Q__"

ANALYTE = "Sermorelin 0.9mg/ml"
LOT_NUMBER = "17009"
MEDIA = "TSB"
MICROBIAL_ID = "Bacillus megaterium"

PRO_DATE = "23Sep26"
PRO_ANALYST_FULL = "Guanchen Li"
PRO_INITIAL = "GL"

READ_ANALYST_FULL = "Andrew Carrillo"
READ_INITIAL = "ANC"

AIR_ANALYST = "SMO"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=30, bottom=30, left=45, right=45):
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

def format_cell(cell, text, bold=False, italic=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER, fill_hex=None, v_align=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = v_align
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

def build_tables_doc():
    doc = docx.Document()
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # =========================================================================
    # Table 1: Information for Sample under investigation
    # =========================================================================
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
    t1_widths = [Inches(1.20), Inches(1.20), Inches(1.25), Inches(1.10), Inches(1.15), Inches(1.60)]

    for j, h in enumerate(t1_headers):
        format_cell(t1.cell(0, j), h, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

    format_cell(t1.cell(1, 0), PRO_ANALYST_FULL, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(t1.cell(1, 1), READ_ANALYST_FULL, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

    # Sample ID Hyperlink
    c_sid = t1.cell(1, 2)
    c_sid.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(c_sid)
    p_sid = c_sid.paragraphs[0]
    p_sid.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sid.paragraph_format.space_before = Pt(0)
    p_sid.paragraph_format.space_after = Pt(0)
    add_hyperlink(p_sid, SAMPLE_URL, SAMPLE_ID, font_size=Pt(7))

    # Subculture ID Hyperlink
    c_sub = t1.cell(1, 3)
    c_sub.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_margins(c_sub)
    p_sub = c_sub.paragraphs[0]
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(0)
    add_hyperlink(p_sub, SUBCULTURE_URL, SUBCULTURE_ID, font_size=Pt(7))

    format_cell(t1.cell(1, 4), MEDIA, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
    # Rule 0P: "有Microbial ID就不用staining了", binomial italicized
    format_cell(t1.cell(1, 5), MICROBIAL_ID, bold=False, italic=True, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

    for row in t1.rows:
        for j, w in enumerate(t1_widths):
            row.cells[j].width = w

    # Spacing between tables
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(6)
    p_space.paragraph_format.space_after = Pt(0)

    # =========================================================================
    # Table 2: Environmental Monitoring from Processing Performed on 23Sep26
    # =========================================================================
    p_t2 = doc.add_paragraph()
    p_t2.paragraph_format.space_before = Pt(0)
    p_t2.paragraph_format.space_after = Pt(2)
    r_t2 = p_t2.add_run(f"Table 2: Environmental Monitoring from Processing Performed on {PRO_DATE}")
    r_t2.font.name = "Times New Roman"
    r_t2.font.size = Pt(8.5)
    r_t2.underline = True

    t2 = doc.add_table(rows=0, cols=9)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t2)

    t2_widths = [
        Inches(1.20), # Sampling site
        Inches(0.55), # Frequency
        Inches(0.60), # Date (DDMMMYY)
        Inches(0.50), # Analyst
        Inches(0.95), # Day / Week(s)
        Inches(0.95), # Observation
        Inches(0.95), # EM Plate ETX Number
        Inches(1.20), # Microbial ID
        Inches(0.60)  # Notes
    ]

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

    hdr2 = t2.add_row()
    for j, h in enumerate(t2_headers):
        format_cell(hdr2.cells[j], h, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

    def add_section_divider(table, title_text):
        row = table.add_row()
        cell = row.cells[0]
        for c in row.cells[1:]:
            cell.merge(c)
        format_cell(cell, title_text, bold=True, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.LEFT, fill_hex="E8E8E8")
        return row

    def add_data_row(table, site, freq, date_str, analyst, day_week, obs, etx_num, micro_id, notes, micro_italic=False, etx_url=None):
        row = table.add_row()
        format_cell(row.cells[0], site, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(row.cells[1], freq, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(row.cells[2], date_str, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(row.cells[3], analyst, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(row.cells[4], day_week, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(row.cells[5], obs, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        
        # Cell 6: ETX number (with hyperlink if url provided)
        c_etx = row.cells[6]
        c_etx.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(c_etx)
        p_etx = c_etx.paragraphs[0]
        p_etx.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_etx.paragraph_format.space_before = Pt(0)
        p_etx.paragraph_format.space_after = Pt(0)
        if etx_url and etx_num not in ["Not Applicable", "N/A"]:
            add_hyperlink(p_etx, etx_url, etx_num, font_size=Pt(7))
        else:
            r = p_etx.add_run(etx_num)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(7)
            
        format_cell(row.cells[7], micro_id, bold=False, italic=micro_italic, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(row.cells[8], notes, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        return row

    # --- Section 1: Personnel EM ---
    add_section_divider(t2, "Personnel EM")
    add_data_row(t2, "Personnel (Left Touch and Right Touch)", "Daily", "22Sep26", PRO_INITIAL, "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, "Personnel (Left Touch and Right Touch)", "Daily", "23Sep26", PRO_INITIAL, "Date of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, "Personnel (Left Touch and Right Touch)", "Daily", "24Sep26", PRO_INITIAL, "Date After Testing", "No Growth", "Not Applicable", "Not Applicable", "None")

    # --- Section 2: BSC EM ---
    add_section_divider(t2, "Biological Safety Cabinet EM for E001314")
    add_data_row(t2, "Surface Sampling of ISO 5 (4 locations)", "Daily", "22Sep26", PRO_INITIAL, "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, "Surface Sampling of ISO 5 (4 locations)", "Daily", "23Sep26", PRO_INITIAL, "Date of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, "Surface Sampling of ISO 5 (4 locations)", "Daily", "24Sep26", PRO_INITIAL, "Date After Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, "Settling Sampling of ISO 5 (2 locations)", "Daily", "22Sep26", PRO_INITIAL, "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, "Settling Sampling of ISO 5 (2 locations)", "Daily", "23Sep26", PRO_INITIAL, "Date of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")
    add_data_row(t2, "Settling Sampling of ISO 5 (2 locations)", "Daily", "24Sep26", PRO_INITIAL, "Date After Testing", "No Growth", "Not Applicable", "Not Applicable", "None")

    # --- Section 3: Cleanroom Active Air ---
    add_section_divider(t2, "Weekly Active Air Sampling of Cleanroom Suite 115 (E001737) with Processing BSC")
    add_data_row(t2, "Active Air Sampling of Cleanrooms", "Weekly", "23Sep26", AIR_ANALYST, "Week of Testing", "6 CFUs (ISO 8 115)", EM_PLATE_ID, "Pending Identification", "None", etx_url=EM_PLATE_URL)

    # --- Section 4: Cleanroom Surface ---
    add_section_divider(t2, "Weekly Surface Sampling of Cleanroom Suite 115 (E001737) with Processing BSC")
    add_data_row(t2, "Surface Sampling of Cleanrooms", "Weekly", "23Sep26", AIR_ANALYST, "Week of Testing", "No Growth", "Not Applicable", "Not Applicable", "None")

    for row in t2.rows:
        for j, w in enumerate(t2_widths):
            row.cells[j].width = w

    # Save documents
    filename = f"Tables OOS-{OOS_ID} {CLIENT_NAME} - {TEST_NAME}.docx"
    doc_path_local = os.path.join(DOCUMENTS_DIR, filename)
    doc_path_desktop = os.path.join(DESKTOP_DIR, filename)

    doc.save(doc_path_local)
    doc.save(doc_path_desktop)
    print(f"Saved DOCX locally: {doc_path_local}")
    print(f"Saved DOCX to Desktop: {doc_path_desktop}")

    # Convert to PDF via Word COM to check exact page count and verify visual appearance
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        wb = word.Documents.Open(os.path.abspath(doc_path_local))
        pdf_temp = os.path.abspath(f"scratch/verify_tables_{OOS_ID}.pdf")
        wb.SaveAs(pdf_temp, FileFormat=17) # wdFormatPDF = 17
        wb.Close()
        
        pdf_doc = fitz.open(pdf_temp)
        page_count = len(pdf_doc)
        print(f"PDF Page count: {page_count}")
        
        # Render image of page 1 to verify layout visually
        page = pdf_doc[0]
        pix = page.get_pixmap(dpi=150)
        preview_img = f"scratch/preview_tables_{OOS_ID}_p1.png"
        pix.save(preview_img)
        print(f"Saved page 1 preview: {preview_img}")
        
        if page_count > 1:
            print("WARNING: Table exceeded 1 page! Further compression needed.")
            for p_idx in range(1, page_count):
                pix_extra = pdf_doc[p_idx].get_pixmap(dpi=150)
                pix_extra.save(f"scratch/preview_tables_{OOS_ID}_p{p_idx+1}.png")
        else:
            print("SUCCESS: Table fits perfectly on exactly 1 single page!")
            
    finally:
        word.Quit()

if __name__ == "__main__":
    build_tables_doc()
