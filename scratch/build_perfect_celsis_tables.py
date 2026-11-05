import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import win32com.client
import fitz

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCX_OUT = os.path.join(DESKTOP_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx")
PDF_OUT = os.path.join(r"scratch", "temp_tables_oos262333_3p.pdf")

doc = docx.Document()

# Set standard margins (0.5 in left/right, 0.5 in top/bottom for clean table fit)
for section in doc.sections:
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)

def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_shading(cell, color_hex):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = OxmlElement(tag)
            element.set(qn('w:val'), edge_data.get('val', 'single'))
            element.set(qn('w:sz'), str(edge_data.get('sz', 4)))
            element.set(qn('w:space'), '0')
            element.set(qn('w:color'), edge_data.get('color', '000000'))
            tcBorders.append(element)
    tcPr.append(tcBorders)

def add_hyperlink(paragraph, url, text, font_size=7.0, bold=False, italic=False):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{r_id}"/>')
    new_run = parse_xml(
        f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'{"<w:b/>" if bold else ""}'
        f'{"<w:i/>" if italic else ""}'
        f'<w:color w:val="0563C1"/>'
        f'<w:u w:val="single"/>'
        f'<w:sz w:val="{int(font_size*2)}"/>'
        f'<w:szCs w:val="{int(font_size*2)}"/>'
        f'</w:rPr>'
        f'<w:t>{text}</w:t>'
        f'</w:r>'
    )
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

# ==========================================
# PAGE 1: TABLE 1
# ==========================================
# Some spacing before Table 1
for _ in range(8):
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(0)
    p_sp.paragraph_format.space_after = Pt(0)

p_t1_title = doc.add_paragraph()
p_t1_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_t1_title.paragraph_format.space_after = Pt(12)
run_t1_title = p_t1_title.add_run("Table 1: Reading analyst, Days of incubation, and Microbial ID for Sample under Investigation")
run_t1_title.font.name = "Times New Roman"
run_t1_title.font.size = Pt(10.0)
run_t1_title.bold = True

t1 = doc.add_table(rows=2, cols=6)
t1.alignment = WD_TABLE_ALIGNMENT.CENTER
t1_col_widths = [Inches(1.52), Inches(1.04), Inches(1.31), Inches(1.38), Inches(1.25), Inches(1.12)]

t1_headers = ["Reading Analyst", "Day of Incubation", "Sample ID", "Related Microbial ID", "Media with growth", "Microbial ID"]
for col_idx, h_text in enumerate(t1_headers):
    cell = t1.cell(0, col_idx)
    cell.width = t1_col_widths[col_idx]
    set_cell_shading(cell, "E8E8E8")
    set_cell_border(cell, top={'sz': 4, 'val': 'single', 'color': '000000'},
                          bottom={'sz': 4, 'val': 'single', 'color': '000000'},
                          left={'sz': 4, 'val': 'single', 'color': '000000'},
                          right={'sz': 4, 'val': 'single', 'color': '000000'})
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(h_text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(9.0)
    r.bold = True

# T1 Row 1 data
row1_cells = t1.rows[1].cells
for c in row1_cells:
    set_cell_border(c, top={'sz': 4, 'val': 'single', 'color': '000000'},
                       bottom={'sz': 4, 'val': 'single', 'color': '000000'},
                       left={'sz': 4, 'val': 'single', 'color': '000000'},
                       right={'sz': 4, 'val': 'single', 'color': '000000'})
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

# Col 0: Reading Analyst
r0 = row1_cells[0].paragraphs[0].add_run("CCD")
r0.font.name = "Times New Roman"
r0.font.size = Pt(8.5)

# Col 1: Day of Incubation
r1 = row1_cells[1].paragraphs[0].add_run("Day 7")
r1.font.name = "Times New Roman"
r1.font.size = Pt(8.5)

# Col 2: Sample ID (hyperlinked)
add_hyperlink(row1_cells[2].paragraphs[0],
              "https://etrax.eagleanalytical.com/Submission/Details/Q57eK3N9WcR0Qh%24x6N5N6Q__",
              "ETX-261002-0288", font_size=8.5, bold=False)

# Col 3: Related Microbial ID
r3 = row1_cells[3].paragraphs[0].add_run("ETX-261012-0654")
r3.font.name = "Times New Roman"
r3.font.size = Pt(8.5)

# Col 4: Media with growth
r4 = row1_cells[4].paragraphs[0].add_run("1 x 300mL FTM")
r4.font.name = "Times New Roman"
r4.font.size = Pt(8.5)

# Col 5: Microbial ID
p5 = row1_cells[5].paragraphs[0]
r5_sp = p5.add_run("Bacillus paralicheniformis\n")
r5_sp.font.name = "Times New Roman"
r5_sp.font.size = Pt(8.5)
r5_sp.italic = True
r5_gr = p5.add_run("(Gram-positive rod)")
r5_gr.font.name = "Times New Roman"
r5_gr.font.size = Pt(8.0)

# End of Page 1 -> Explicit Page Break
doc.add_page_break()

# ==========================================
# PAGE 2: TABLE 2 (Processing EM)
# ==========================================
p_t2_title = doc.add_paragraph()
p_t2_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_t2_title.paragraph_format.space_before = Pt(0)
p_t2_title.paragraph_format.space_after = Pt(6)
run_t2_title = p_t2_title.add_run("Table 2: Environmental Monitoring from Celsis Processing Performed on 05Oct2026")
run_t2_title.font.name = "Times New Roman"
run_t2_title.font.size = Pt(9.5)
run_t2_title.bold = True

t2 = doc.add_table(rows=16, cols=10)
t2.alignment = WD_TABLE_ALIGNMENT.CENTER
t2_widths = [Inches(1.05), Inches(0.85), Inches(0.65), Inches(0.55), Inches(0.95), Inches(0.95), Inches(0.85), Inches(0.95), Inches(0.95), Inches(0.50)]

em_headers = ['Environmental Monitoring (EM) Sampling Site', 'Frequency', 'Date', 'Analyst', 'Day /Week(s)', 'Day /Week(s)', 'Observation', 'EM Plate ETX Number', 'Microbial ID', 'Notes']

def format_row(row, values, bg_color=None, bold=False, font_size=7.0, italic_col=None, url_col=None, is_banner=False):
    for c_idx, cell in enumerate(row.cells):
        cell.width = t2_widths[c_idx]
        if bg_color:
            set_cell_shading(cell, bg_color)
        set_cell_border(cell, top={'sz': 4, 'val': 'single', 'color': '000000'},
                              bottom={'sz': 4, 'val': 'single', 'color': '000000'},
                              left={'sz': 4, 'val': 'single', 'color': '000000'},
                              right={'sz': 4, 'val': 'single', 'color': '000000'})
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(cell, top=40, bottom=40, left=40, right=40)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0

    if is_banner:
        # Merge all cells across the row
        first_cell = row.cells[0]
        for c in row.cells[1:]:
            first_cell.merge(c)
        p = first_cell.paragraphs[0]
        p.text = values[0]
        r = p.runs[0]
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size)
        r.bold = bold
        return

    for c_idx, text_val in enumerate(values):
        cell = row.cells[c_idx]
        p = cell.paragraphs[0]
        p.text = ""
        if url_col and c_idx == url_col[0]:
            add_hyperlink(p, url_col[1], text_val, font_size=font_size, bold=bold)
        elif italic_col and c_idx == italic_col:
            # Species lines
            lines = text_val.split('\n')
            for l_idx, line in enumerate(lines):
                if line.strip():
                    r = p.add_run(line)
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(font_size)
                    r.italic = True
                    r.bold = bold
                if l_idx < len(lines) - 1:
                    p.add_run('\n')
        else:
            lines = text_val.split('\n')
            for l_idx, line in enumerate(lines):
                r = p.add_run(line)
                r.font.name = "Times New Roman"
                r.font.size = Pt(font_size)
                r.bold = bold
                if l_idx < len(lines) - 1:
                    p.add_run('\n')

# Row 0: Header
format_row(t2.rows[0], em_headers, bg_color="E8E8E8", bold=True, font_size=7.0)
# Merge Day/Week(s) header columns 4 and 5
t2.rows[0].cells[4].merge(t2.rows[0].cells[5])

# Row 1: Banner 1
format_row(t2.rows[1], ["Personnel EM Bracketing 05Oct2026"], bg_color="E8E8E8", bold=True, font_size=7.0, is_banner=True)

# Rows 2-4: Personnel EM
t2_data = [
    ['Personal (Left Touch and Right Touch)', 'Daily', '02Oct2026', 'ES', 'Date Before Testing', 'Date Before Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Personal (Left Touch and Right Touch)', 'Daily', '05Oct2026', 'ES', 'Date of Testing', 'Date of Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Personal (Left Touch and Right Touch)', 'Daily', '06Oct2026', 'ES', 'Date After Testing', 'Date After Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
]
for idx, d in enumerate(t2_data):
    format_row(t2.rows[idx+2], d, font_size=7.0)
    t2.rows[idx+2].cells[4].merge(t2.rows[idx+2].cells[5])

# Row 5: Banner 2
format_row(t2.rows[5], ["Biological Safety Cabinet EM Bracketing E001316 for 05Oct2026"], bg_color="E8E8E8", bold=True, font_size=7.0, is_banner=True)

# Rows 6-11: BSC EM
t2_bsc = [
    ['Surface Sampling of ISO 5 BSC E001316 (4 Locations)', 'Daily', '02Oct2026', 'ALA', 'Date Before Testing', 'Date Before Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of ISO 5 BSC E001316 (4 Locations)', 'Daily', '05Oct2026', 'ES', 'Date of Testing', 'Date of Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of ISO 5 BSC E001316 (4 Locations)', 'Daily', '06Oct2026', 'ES', 'Date After Testing', 'Date After Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 BSC E001316', 'Daily', '02Oct2026', 'ALA', 'Date Before Testing', 'Date Before Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 BSC E001316', 'Daily', '05Oct2026', 'ES', 'Date of Testing', 'Date of Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 BSC E001316', 'Daily', '06Oct2026', 'ES', 'Date After Testing', 'Date After Testing', 'No Growth', 'Not Applicable', 'None', 'None'],
]
for idx, d in enumerate(t2_bsc):
    format_row(t2.rows[idx+6], d, font_size=7.0)
    t2.rows[idx+6].cells[4].merge(t2.rows[idx+6].cells[5])

# Row 12: Banner 3
format_row(t2.rows[12], ["Weekly Active Air and Surface Sampling of Cleanroom Suite 114 with Processing BSC For 05Oct2026"], bg_color="E8E8E8", bold=True, font_size=7.0, is_banner=True)

# Row 13: Subheader
format_row(t2.rows[13], ['Environmental Monitoring (EM) Sampling Site', 'Frequency', 'Date', 'Analyst', 'Observation', 'Observation', 'EM Plate ETX Number', 'EM Plate ETX Number', 'Microbial ID', 'Notes'], bg_color="E8E8E8", bold=True, font_size=7.0)
t2.rows[13].cells[4].merge(t2.rows[13].cells[5])
t2.rows[13].cells[6].merge(t2.rows[13].cells[7])

# Rows 14-15: Weekly Sampling
t2_weekly = [
    ['Active Air Sampling of Cleanrooms', 'Weekly', '05Oct2026', 'SMO', 'No Growth', 'No Growth', 'Not Applicable', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of Cleanrooms', 'Weekly', '05Oct2026', 'SMO', 'No Growth', 'No Growth', 'Not Applicable', 'Not Applicable', 'None', 'None']
]
for idx, d in enumerate(t2_weekly):
    format_row(t2.rows[idx+14], d, font_size=7.0)
    t2.rows[idx+14].cells[4].merge(t2.rows[idx+14].cells[5])
    t2.rows[idx+14].cells[6].merge(t2.rows[idx+14].cells[7])

# End of Page 2 -> Explicit Page Break
doc.add_page_break()

# ==========================================
# PAGE 3: TABLE 3 (Aliquoting EM)
# ==========================================
p_t3_title = doc.add_paragraph()
p_t3_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_t3_title.paragraph_format.space_before = Pt(0)
p_t3_title.paragraph_format.space_after = Pt(6)
run_t3_title = p_t3_title.add_run("Table 3: Environmental Monitoring from Celsis Aliquoting Performed on 12Oct2026")
run_t3_title.font.name = "Times New Roman"
run_t3_title.font.size = Pt(9.5)
run_t3_title.bold = True

t3 = doc.add_table(rows=16, cols=10)
t3.alignment = WD_TABLE_ALIGNMENT.CENTER

# Row 0: Header
format_row(t3.rows[0], em_headers, bg_color="E8E8E8", bold=True, font_size=7.0)
t3.rows[0].cells[4].merge(t3.rows[0].cells[5])

# Row 1: Banner 1
format_row(t3.rows[1], ["Personnel EM Bracketing 12Oct2026"], bg_color="E8E8E8", bold=True, font_size=7.0, is_banner=True)

# Rows 2-4: Personnel EM (CCD)
t3_data = [
    ['Personal (Left Touch and Right Touch)', 'Daily', '09Oct2026', 'CCD', 'Date Before Aliquoting', 'Date Before Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Personal (Left Touch and Right Touch)', 'Daily', '12Oct2026', 'CCD', 'Date of Aliquoting', 'Date of Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Personal (Left Touch and Right Touch)', 'Daily', '13Oct2026', 'CCD', 'Date After Aliquoting', 'Date After Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
]
for idx, d in enumerate(t3_data):
    format_row(t3.rows[idx+2], d, font_size=7.0)
    t3.rows[idx+2].cells[4].merge(t3.rows[idx+2].cells[5])

# Row 5: Banner 2
format_row(t3.rows[5], ["Biological Safety Cabinet EM for 12Oct2026"], bg_color="E8E8E8", bold=True, font_size=7.0, is_banner=True)

# Rows 6-11: BSC EM (E001798)
t3_bsc = [
    ['Surface Sampling of ISO 5 BSC E001798 (4 Locations)', 'Daily', '09Oct2026', 'CCD', 'Date Before Aliquoting', 'Date Before Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of ISO 5 BSC E001798 (4 Locations)', 'Daily', '12Oct2026', 'CCD', 'Date of Aliquoting', 'Date of Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Surface Sampling of ISO 5 BSC E001798 (4 Locations)', 'Daily', '13Oct2026', 'CCD', 'Date After Aliquoting', 'Date After Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 BSC E001798', 'Daily', '09Oct2026', 'CCD', 'Date Before Aliquoting', 'Date Before Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 BSC E001798', 'Daily', '12Oct2026', 'CCD', 'Date of Aliquoting', 'Date of Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
    ['Settling Sampling of ISO 5 BSC E001798', 'Daily', '13Oct2026', 'CCD', 'Date After Aliquoting', 'Date After Aliquoting', 'No Growth', 'Not Applicable', 'None', 'None'],
]
for idx, d in enumerate(t3_bsc):
    format_row(t3.rows[idx+6], d, font_size=7.0)
    t3.rows[idx+6].cells[4].merge(t3.rows[idx+6].cells[5])

# Row 12: Banner 3
format_row(t3.rows[12], ["Weekly Active Air and Surface Sampling of Cleanroom Suite 114 with Processing BSC For 12Oct2026"], bg_color="E8E8E8", bold=True, font_size=7.0, is_banner=True)

# Row 13: Subheader
format_row(t3.rows[13], ['Environmental Monitoring (EM) Sampling Site', 'Frequency', 'Date', 'Analyst', 'Observation', 'Observation', 'EM Plate ETX Number', 'EM Plate ETX Number', 'Microbial ID', 'Notes'], bg_color="E8E8E8", bold=True, font_size=7.0)
t3.rows[13].cells[4].merge(t3.rows[13].cells[5])
t3.rows[13].cells[6].merge(t3.rows[13].cells[7])

# Row 14: Active Air (6 CFUs with Oxford Nanopore confirmed species)
microbial_species = (
    "Moraxella osloensis\n\n"
    "Sporisorium graminicola\n\n"
    "Staphylococcus haemolyticus\n\n"
    "Staphylococcus hominis\n\n"
    "Kocuria indica\n\n"
    "Ustilago maydis\n\n"
    "Kocuria rhizophila"
)
d_active_air = ['Active Air Sampling of Cleanrooms', 'Weekly', '12Oct2026', 'SMO', '6 CFUs\n(ISO 8 114)', '6 CFUs\n(ISO 8 114)', 'ETX-261021-0578', 'ETX-261021-0578', microbial_species, 'None']
format_row(t3.rows[14], d_active_air, font_size=7.0, italic_col=8, url_col=(6, "https://etrax.eagleanalytical.com/SubmissionTest/Details/Cw%24QZhYAogGUwoutnOavpw__"))
t3.rows[14].cells[4].merge(t3.rows[14].cells[5])
t3.rows[14].cells[6].merge(t3.rows[14].cells[7])

# Row 15: Surface Sampling
d_surface = ['Surface Sampling of Cleanrooms', 'Weekly', '12Oct2026', 'SMO', 'No Growth', 'No Growth', 'Not Applicable', 'Not Applicable', 'None', 'None']
format_row(t3.rows[15], d_surface, font_size=7.0)
t3.rows[15].cells[4].merge(t3.rows[15].cells[5])
t3.rows[15].cells[6].merge(t3.rows[15].cells[7])

# Save Document
doc.save(DOCX_OUT)
print(f"Saved DOCX to: {DOCX_OUT}")

# Convert to PDF via Word COM
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
try:
    wdoc = word.Documents.Open(os.path.abspath(DOCX_OUT))
    wdoc.SaveAs(os.path.abspath(PDF_OUT), FileFormat=17) # 17 = wdFormatPDF
    wdoc.Close()
finally:
    word.Quit()

print(f"Converted PDF to: {PDF_OUT}")

# Verify Page Count of Generated PDF
doc_pdf = fitz.open(PDF_OUT)
print(f"VERIFIED STANDALONE TABLES PDF PAGE COUNT: {len(doc_pdf)} pages.")
for pno in range(len(doc_pdf)):
    txt = doc_pdf[pno].get_text()[:60].strip().replace('\n', ' ')
    print(f"  Page {pno+1}: {txt}")
assert len(doc_pdf) == 3, f"Expected strictly 3 pages, got {len(doc_pdf)}"
doc_pdf.close()
print("SUCCESS: Standalone Word Table is strictly 3 pages!")
