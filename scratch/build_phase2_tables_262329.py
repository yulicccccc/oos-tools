import os, sys, shutil, fitz, datetime
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
SCRATCH_DIR = os.path.join(DOCUMENTS_DIR, "scratch")

OOS_ID = "262329"
CLIENT_NAME = "GoGoMeds Select (E10747)"
TEST_NAME = "USP71"

ORIGINAL_ID = "ETX-260624-0251 (BUD-T90)"
RETEST_ID = "ETX-261021-0445"

ORIGINAL_URL = "https://etrax.eagleanalytical.com/Submission/Details/QY-h1u_rYkWQY-98oK0eUA__"
RETEST_URL = "https://etrax.eagleanalytical.com/Submission"

def set_cell_margins(cell, top=40, bottom=40, left=60, right=60):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name, border_props in kwargs.items():
        node = OxmlElement(f'w:{border_name}')
        for key, val in border_props.items():
            node.set(qn(f'w:{key}'), str(val))
        tcBorders.append(node)
    tcPr.append(tcBorders)

def add_hyperlink(paragraph, url, text, font_size_pt=7, color_rgb="0000FF"):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w")} r:id="{r_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
    rPr.append(parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'))
    rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{int(font_size_pt * 2)}"/>'))
    rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{int(font_size_pt * 2)}"/>'))
    rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="{color_rgb}"/>'))
    rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>'))
    new_run.append(rPr)
    t = parse_xml(f'<w:t {nsdecls("w")}>{text}</w:t>')
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def format_cell(cell, text="", bold=False, italic=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER, fill=None):
    set_cell_margins(cell, top=35, bottom=35, left=50, right=50)
    border_kw = {'val': 'single', 'sz': '4', 'space': '0', 'color': '000000'}
    set_cell_border(cell, top=border_kw, bottom=border_kw, left=border_kw, right=border_kw)
    if fill:
        set_cell_background(cell, fill)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    if text:
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic
        r.font.name = 'Times New Roman'
        r.font.size = font_size
    return p

def build_phase2_tables_doc():
    doc = docx.Document()
    
    # 0.75 inch top margin for whitespace ("留点白"), 0.5 inch for sides
    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.5)
        sec.left_margin = Inches(0.5)
        sec.right_margin = Inches(0.5)

    # Title 1
    p_t1 = doc.add_paragraph()
    p_t1.paragraph_format.space_before = Pt(12)
    p_t1.paragraph_format.space_after = Pt(4)
    r = p_t1.add_run(f"Table 1: Information for {ORIGINAL_ID} and {RETEST_ID} under investigation")
    r.font.name = 'Times New Roman'; r.font.size = Pt(8); r.bold = False

    # Table 1: 3 rows, 7 cols
    t1 = doc.add_table(rows=3, cols=7)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1_widths = [Inches(1.0), Inches(0.9), Inches(1.5), Inches(1.3), Inches(1.3), Inches(0.7), Inches(0.8)]
    
    headers_t1 = [
        "Processing Analyst", "Reading Analyst", "Sample ID", 
        "Related Microbial ID", "Media with microbial growth", "Sterility Result", "Microbial ID"
    ]
    for j, h in enumerate(headers_t1):
        format_cell(t1.rows[0].cells[j], h, bold=False, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        
    # Row 1: Original Test
    format_cell(t1.rows[1].cells[0], "Guanchen Li", font_size=Pt(7))
    format_cell(t1.rows[1].cells[1], "Andrew Carrillo", font_size=Pt(7))
    p_orig = format_cell(t1.rows[1].cells[2], "", font_size=Pt(7))
    add_hyperlink(p_orig, ORIGINAL_URL, ORIGINAL_ID, font_size_pt=7)
    
    p_sub = format_cell(t1.rows[1].cells[3], "", font_size=Pt(7))
    add_hyperlink(p_sub, "https://etrax.eagleanalytical.com/Submission/Details/e6Jo3DyLpUEd3QnYm5Uttw__", "ETX-261012-0367", font_size_pt=7)
    
    format_cell(t1.rows[1].cells[4], "2 x 200mL TSB (Original and Subculture Bottles)", font_size=Pt(7))
    format_cell(t1.rows[1].cells[5], "Fail", bold=False, font_size=Pt(7))
    
    p_mic1 = format_cell(t1.rows[1].cells[6], "", font_size=Pt(7))
    r_bm = p_mic1.add_run("Bacillus megaterium\n")
    r_bm.font.name = 'Times New Roman'; r_bm.font.size = Pt(7); r_bm.italic = True
    r_bm_g = p_mic1.add_run("(Gram (+) rods)")
    r_bm_g.font.name = 'Times New Roman'; r_bm_g.font.size = Pt(7); r_bm_g.italic = False

    # Row 2: Phase II Retest
    format_cell(t1.rows[2].cells[0], "Ayomide Odugbesi", font_size=Pt(7))
    format_cell(t1.rows[2].cells[1], "Andrew Carrillo", font_size=Pt(7))
    p_ret = format_cell(t1.rows[2].cells[2], "", font_size=Pt(7))
    add_hyperlink(p_ret, RETEST_URL, RETEST_ID, font_size_pt=7)
    
    format_cell(t1.rows[2].cells[3], "Not Applicable", font_size=Pt(7))
    format_cell(t1.rows[2].cells[4], "No Growth", font_size=Pt(7))
    format_cell(t1.rows[2].cells[5], "Pass", bold=False, font_size=Pt(7))
    format_cell(t1.rows[2].cells[6], "None", font_size=Pt(7))

    for row in t1.rows:
        for j, w in enumerate(t1_widths):
            row.cells[j].width = w

    # Space between tables
    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(16)
    p_space.paragraph_format.space_after = Pt(4)
    r_t2 = p_space.add_run("Table 2: Environmental Monitoring from Testing Performed on 23Sep26")
    r_t2.font.name = 'Times New Roman'; r_t2.font.size = Pt(8); r_t2.bold = False

    # Table 2: Environmental Monitoring (Original Testing on 23Sep26)
    t2 = doc.add_table(rows=0, cols=9)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2_widths = [Inches(1.5), Inches(0.6), Inches(0.6), Inches(0.5), Inches(1.0), Inches(0.9), Inches(1.0), Inches(1.1), Inches(0.4)]

    def add_sec_divider(tbl, title_text):
        r = tbl.add_row()
        cell = r.cells[0]
        for c in r.cells[1:]:
            cell.merge(c)
        format_cell(cell, title_text, bold=True, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.LEFT, fill="D9D9D9")
        set_cell_margins(cell, top=25, bottom=25, left=40, right=40)

    def add_em_row(tbl, site, freq, date_str, analyst, day_week, obs, plate_id, micro_fn, notes="None", etx_url=None):
        r = tbl.add_row()
        format_cell(r.cells[0], site, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(r.cells[1], freq, font_size=Pt(7))
        format_cell(r.cells[2], date_str, font_size=Pt(7))
        format_cell(r.cells[3], analyst, font_size=Pt(7))
        format_cell(r.cells[4], day_week, font_size=Pt(7))
        format_cell(r.cells[5], obs, font_size=Pt(7))
        
        p_etx = format_cell(r.cells[6], "", font_size=Pt(7))
        if etx_url and plate_id not in ["Not Applicable", "N/A"]:
            add_hyperlink(p_etx, etx_url, plate_id, font_size_pt=7)
        else:
            r_etx = p_etx.add_run(plate_id)
            r_etx.font.name = 'Times New Roman'; r_etx.font.size = Pt(7)
            
        p_micro = format_cell(r.cells[7], "", font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        if callable(micro_fn):
            micro_fn(p_micro)
        else:
            r_m = p_micro.add_run(str(micro_fn))
            r_m.font.name = 'Times New Roman'; r_m.font.size = Pt(7)
            
        format_cell(r.cells[8], notes, font_size=Pt(7))

    # Banner headers for Table 2
    r_hdr = t2.add_row()
    headers_t2 = [
        "Environmental Monitoring (EM) Sampling Site", "Frequency", "Date (DDMMMYY)", "Analyst",
        "Day /Week(s)", "Observation", "EM Plate ETX Number", "Microbial ID", "Notes"
    ]
    for j, h in enumerate(headers_t2):
        format_cell(r_hdr.cells[j], h, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)

    # Section 1: Personnel EM
    add_sec_divider(t2, "Personnel EM")
    add_em_row(t2, "Personnel (Left Touch and Right Touch)", "Daily", "22Sep26", "GL", "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable")
    add_em_row(t2, "Personnel (Left Touch and Right Touch)", "Daily", "23Sep26", "GL", "Date of Testing", "No Growth", "Not Applicable", "Not Applicable")
    add_em_row(t2, "Personnel (Left Touch and Right Touch)", "Daily", "24Sep26", "GL", "Date After Testing", "No Growth", "Not Applicable", "Not Applicable")

    # Section 2: BSC EM
    add_sec_divider(t2, "Biological Safety Cabinet EM for E001314")
    add_em_row(t2, "Surface Sampling of ISO 5 (4 locations)", "Daily", "22Sep26", "AOD", "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable")
    add_em_row(t2, "Surface Sampling of ISO 5 (4 locations)", "Daily", "23Sep26", "GL", "Date of Testing", "No Growth", "Not Applicable", "Not Applicable")
    add_em_row(t2, "Surface Sampling of ISO 5 (4 locations)", "Daily", "24Sep26", "AOD", "Date After Testing", "No Growth", "Not Applicable", "Not Applicable")
    add_em_row(t2, "Settling Sampling of ISO 5 (2 locations)", "Daily", "22Sep26", "AOD", "Date Before Testing", "No Growth", "Not Applicable", "Not Applicable")
    add_em_row(t2, "Settling Sampling of ISO 5 (2 locations)", "Daily", "23Sep26", "GL", "Date of Testing", "No Growth", "Not Applicable", "Not Applicable")
    add_em_row(t2, "Settling Sampling of ISO 5 (2 locations)", "Daily", "24Sep26", "AOD", "Date After Testing", "No Growth", "Not Applicable", "Not Applicable")

    # Section 3: Active Air
    add_sec_divider(t2, "Weekly Active Air Sampling of Cleanroom Suite 115 (E001737) with Processing BSC")
    def fill_air_micro(p):
        r1 = p.add_run("Micrococcus luteus\n\n")
        r1.font.name = 'Times New Roman'; r1.font.size = Pt(7); r1.italic = True
        r2a = p.add_run("Kocuria indica\n\n")
        r2a.font.name = 'Times New Roman'; r2a.font.size = Pt(7); r2a.italic = True
        r2b = p.add_run("Kocuria rhizophila\n\n")
        r2b.font.name = 'Times New Roman'; r2b.font.size = Pt(7); r2b.italic = True
        r3 = p.add_run("Pending FunITS (Hyphae)")
        r3.font.name = 'Times New Roman'; r3.font.size = Pt(7); r3.italic = False

    add_em_row(t2, "Active Air Sampling of Cleanrooms", "Weekly", "23Sep26", "SMO", "Week of Testing", "6 CFUs (ISO 8 115)", "ETX-261001-0625", fill_air_micro, etx_url="https://etrax.eagleanalytical.com/Submission/Details/yp3RLi-9BGnGhzxnYnqk4Q__")

    # Section 4: Surface
    add_sec_divider(t2, "Weekly Surface Sampling of Cleanroom Suite 115 (E001737) with Processing BSC")
    add_em_row(t2, "Surface Sampling of Cleanrooms", "Weekly", "23Sep26", "SMO", "Week of Testing", "No Growth", "Not Applicable", "Not Applicable")

    for row in t2.rows:
        for j, w in enumerate(t2_widths):
            row.cells[j].width = w

    # Uniform XML 14 half-points font size
    for p in doc.paragraphs:
        for r in p.runs:
            r.font.name = 'Times New Roman'
    for tbl in [t1, t2]:
        for row in tbl.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    for r in p.runs:
                        r.font.name = 'Times New Roman'
                        r.font.size = Pt(7)

    # Save Word doc
    filename = f"Tables OOS-{OOS_ID} {CLIENT_NAME} - {TEST_NAME}.docx"
    doc_path_local = os.path.join(DOCUMENTS_DIR, filename)
    doc_path_desktop = os.path.join(DESKTOP_DIR, filename)

    doc.save(doc_path_local)
    try:
        shutil.copy2(doc_path_local, doc_path_desktop)
        print("Saved Phase II Tables Word doc to Documents and Desktop.")
    except Exception as e:
        print("Notice copying Word doc to Desktop:", e)

    # Convert to PDF
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    pdf_temp_path = os.path.join(SCRATCH_DIR, f"verify_tables_{OOS_ID}.pdf")
    try:
        wdoc = word.Documents.Open(os.path.abspath(doc_path_local))
        wdoc.SaveAs(os.path.abspath(pdf_temp_path), FileFormat=17)
        wdoc.Close()
        shutil.copy2(pdf_temp_path, os.path.join(SCRATCH_DIR, f"verify_tables_phase2_{OOS_ID}.pdf"))
    finally:
        word.Quit()

    pdf_doc = fitz.open(pdf_temp_path)
    page_count = len(pdf_doc)
    print(f"Phase II Tables PDF generated: {page_count} page(s).")
    pdf_doc.close()

if __name__ == "__main__":
    build_phase2_tables_doc()
