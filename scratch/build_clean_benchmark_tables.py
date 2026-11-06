import docx
import os
import fitz
import win32com.client
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BENCHMARK_DOCX = r"scratch/qa_celsis_emails/262403_PGRrx EM TABLE.docx"
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
TARGET_DOCX = os.path.join(DESKTOP_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx")
TARGET_PDF = r"scratch/temp_tables_oos262333_3p_clean.pdf"

doc = docx.Document(BENCHMARK_DOCX)

# Set 0.45 in top/bottom, 0.5 in left/right margins to guarantee generous breathing room
for s in doc.sections:
    s.top_margin = Inches(0.45)
    s.bottom_margin = Inches(0.45)
    s.left_margin = Inches(0.5)
    s.right_margin = Inches(0.5)

# Remove all empty paragraphs in body
body = doc.element.body
for p_el in list(body.xpath('w:p')):
    p = docx.text.paragraph.Paragraph(p_el, doc)
    if not p.text.strip():
        body.remove(p_el)

def insert_page_break_before(element):
    p_pb = parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>')
    element.addprevious(p_pb)

for p_el in body.xpath('w:p'):
    p = docx.text.paragraph.Paragraph(p_el, doc)
    if "Table 2:" in p.text:
        insert_page_break_before(p_el)
    elif "Table 3:" in p.text:
        insert_page_break_before(p_el)

def set_cell_simple_text(cell, text, font_size_pt=7.0, bold=False, italic=False):
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    p = docx.text.paragraph.Paragraph(parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>'), cell)
    lines = text.split('\n')
    for idx, line in enumerate(lines):
        r = p.add_run(line)
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size_pt)
        r.bold = bold
        r.italic = italic
        if idx < len(lines) - 1:
            p.add_run('\n')
    tc.append(p._p)

def set_cell_hyperlink(cell, url, text, font_size_pt=7.0, bold=False):
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    p = docx.text.paragraph.Paragraph(parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>'), cell)
    part = cell.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{r_id}"/>')
    new_run = parse_xml(
        f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'{"<w:b/>" if bold else ""}'
        f'<w:color w:val="0563C1"/>'
        f'<w:u w:val="single"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'<w:t>{text}</w:t>'
        f'</w:r>'
    )
    hyperlink.append(new_run)
    p._p.append(hyperlink)
    tc.append(p._p)

def set_cell_multispecies(cell, species_list, font_size_pt=7.0):
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    for i, sp in enumerate(species_list):
        p = parse_xml(
            f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:pPr>'
            f'<w:jc w:val="center"/>'
            f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'</w:pPr>'
            f'<w:r>'
            f'<w:rPr>'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'<w:i/>'
            f'<w:sz w:val="{int(font_size_pt*2)}"/>'
            f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
            f'</w:rPr>'
            f'<w:t>{sp}</w:t>'
            f'</w:r>'
            f'</w:p>'
        )
        tc.append(p)
        if i < len(species_list) - 1:
            p_blank = parse_xml(
                f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:pPr>'
                f'<w:jc w:val="center"/>'
                f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
                f'</w:pPr>'
                f'</w:p>'
            )
            tc.append(p_blank)

# 1. Update Titles
for p in doc.paragraphs:
    if "Table 2:" in p.text:
        p.text = "Table 2: Environmental Monitoring from Celsis Processing Performed on 05Oct2026"
        p.runs[0].font.name = "Times New Roman"
        p.runs[0].font.size = Pt(10.0)
        p.runs[0].bold = True
        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
    elif "Table 3:" in p.text:
        p.text = "Table 3: Environmental Monitoring from Celsis Aliquoting Performed on 12Oct2026"
        p.runs[0].font.name = "Times New Roman"
        p.runs[0].font.size = Pt(10.0)
        p.runs[0].bold = True
        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)

# 2. Table 1
t1 = doc.tables[0]
r1 = t1.rows[1]
set_cell_simple_text(r1.cells[0], "CCD", font_size_pt=8.5)
set_cell_simple_text(r1.cells[1], "Day 7", font_size_pt=8.5)
set_cell_hyperlink(r1.cells[2], "https://etrax.eagleanalytical.com/Submission/Details/Q57eK3N9WcR0Qh%24x6N5N6Q__", "ETX-261002-0288", font_size_pt=8.5)
set_cell_simple_text(r1.cells[3], "ETX-261012-0654", font_size_pt=8.5)
set_cell_simple_text(r1.cells[4], "1 x 300mL FTM", font_size_pt=8.5)

tc_m = r1.cells[5]._tc
for p in tc_m.xpath('w:p'):
    tc_m.remove(p)
p_m = docx.text.paragraph.Paragraph(parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>'), r1.cells[5])
r_sp = p_m.add_run("Bacillus paralicheniformis\n")
r_sp.font.name = "Times New Roman"
r_sp.font.size = Pt(8.5)
r_sp.italic = True
r_gr = p_m.add_run("(Gram-positive rod)")
r_gr.font.name = "Times New Roman"
r_gr.font.size = Pt(8.0)
tc_m.append(p_m._p)

# 3. Table 2
t2 = doc.tables[1]
set_cell_simple_text(t2.rows[1].cells[0], "Personnel EM Bracketing 05Oct2026", font_size_pt=7.0, bold=True)
set_cell_simple_text(t2.rows[2].cells[2], "02Oct2026", font_size_pt=7.0)
set_cell_simple_text(t2.rows[3].cells[2], "05Oct2026", font_size_pt=7.0)
set_cell_simple_text(t2.rows[4].cells[2], "06Oct2026", font_size_pt=7.0)

set_cell_simple_text(t2.rows[5].cells[0], "Biological Safety Cabinet EM Bracketing E001316 for 05Oct2026", font_size_pt=7.0, bold=True)
set_cell_simple_text(t2.rows[6].cells[2], "02Oct2026", font_size_pt=7.0)
set_cell_simple_text(t2.rows[6].cells[3], "ALA", font_size_pt=7.0)
set_cell_simple_text(t2.rows[7].cells[2], "05Oct2026", font_size_pt=7.0)
set_cell_simple_text(t2.rows[7].cells[3], "ES", font_size_pt=7.0)
set_cell_simple_text(t2.rows[8].cells[2], "06Oct2026", font_size_pt=7.0)
set_cell_simple_text(t2.rows[8].cells[3], "ES", font_size_pt=7.0)

set_cell_simple_text(t2.rows[9].cells[2], "02Oct2026", font_size_pt=7.0)
set_cell_simple_text(t2.rows[9].cells[3], "ALA", font_size_pt=7.0)
set_cell_simple_text(t2.rows[10].cells[2], "05Oct2026", font_size_pt=7.0)
set_cell_simple_text(t2.rows[10].cells[3], "ES", font_size_pt=7.0)
set_cell_simple_text(t2.rows[11].cells[2], "06Oct2026", font_size_pt=7.0)
set_cell_simple_text(t2.rows[11].cells[3], "ES", font_size_pt=7.0)

set_cell_simple_text(t2.rows[12].cells[0], "Weekly Active Air and Surface Sampling of Cleanroom Suite 114 with Processing BSC For 05Oct2026", font_size_pt=7.0, bold=True)
set_cell_simple_text(t2.rows[14].cells[2], "05Oct2026", font_size_pt=7.0)
set_cell_simple_text(t2.rows[14].cells[3], "SMO", font_size_pt=7.0)
set_cell_simple_text(t2.rows[14].cells[4], "No Growth", font_size_pt=7.0)
set_cell_simple_text(t2.rows[14].cells[5], "Not Applicable", font_size_pt=7.0)
set_cell_simple_text(t2.rows[14].cells[7], "Not Applicable", font_size_pt=7.0)
set_cell_simple_text(t2.rows[14].cells[9], "None", font_size_pt=7.0)

set_cell_simple_text(t2.rows[15].cells[2], "05Oct2026", font_size_pt=7.0)
set_cell_simple_text(t2.rows[15].cells[3], "SMO", font_size_pt=7.0)
set_cell_simple_text(t2.rows[15].cells[4], "No Growth", font_size_pt=7.0)
set_cell_simple_text(t2.rows[15].cells[5], "Not Applicable", font_size_pt=7.0)
set_cell_simple_text(t2.rows[15].cells[7], "Not Applicable", font_size_pt=7.0)
set_cell_simple_text(t2.rows[15].cells[9], "None", font_size_pt=7.0)

# 4. Table 3
t3 = doc.tables[2]
set_cell_simple_text(t3.rows[1].cells[0], "Personnel EM Bracketing 12Oct2026", font_size_pt=7.0, bold=True)
set_cell_simple_text(t3.rows[2].cells[2], "09Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[2].cells[4], "Date Before Aliquoting", font_size_pt=7.0)
set_cell_simple_text(t3.rows[3].cells[2], "12Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[3].cells[4], "Date of Aliquoting", font_size_pt=7.0)
set_cell_simple_text(t3.rows[4].cells[2], "13Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[4].cells[4], "Date After Aliquoting", font_size_pt=7.0)

set_cell_simple_text(t3.rows[5].cells[0], "Biological Safety Cabinet EM for 12Oct2026", font_size_pt=7.0, bold=True)
set_cell_simple_text(t3.rows[6].cells[2], "09Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[6].cells[4], "Date Before Aliquoting", font_size_pt=7.0)
set_cell_simple_text(t3.rows[7].cells[2], "12Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[7].cells[4], "Date of Aliquoting", font_size_pt=7.0)
set_cell_simple_text(t3.rows[8].cells[2], "13Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[8].cells[4], "Date After Aliquoting", font_size_pt=7.0)

set_cell_simple_text(t3.rows[9].cells[2], "09Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[9].cells[4], "Date Before Aliquoting", font_size_pt=7.0)
set_cell_simple_text(t3.rows[10].cells[2], "12Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[10].cells[4], "Date of Aliquoting", font_size_pt=7.0)
set_cell_simple_text(t3.rows[11].cells[2], "13Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[11].cells[4], "Date After Aliquoting", font_size_pt=7.0)

set_cell_simple_text(t3.rows[12].cells[0], "Weekly Active Air and Surface Sampling of Cleanroom Suite 114 with Processing BSC For 12Oct2026", font_size_pt=7.0, bold=True)

# Row 14: Weekly Active Air (6 CFUs under ETX-261021-0578 on 12Oct2026)
set_cell_simple_text(t3.rows[14].cells[2], "12Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[14].cells[3], "SMO", font_size_pt=7.0)
set_cell_simple_text(t3.rows[14].cells[4], "6 CFUs\n(ISO 8 114)", font_size_pt=7.0)
set_cell_hyperlink(t3.rows[14].cells[5], "https://etrax.eagleanalytical.com/SubmissionTest/Details/Cw%24QZhYAogGUwoutnOavpw__", "ETX-261021-0578", font_size_pt=7.0)

species_list = [
    "Moraxella osloensis",
    "Sporisorium graminicola",
    "Staphylococcus haemolyticus",
    "Staphylococcus hominis",
    "Kocuria indica",
    "Ustilago maydis",
    "Kocuria rhizophila"
]
set_cell_multispecies(t3.rows[14].cells[7], species_list, font_size_pt=7.0)
set_cell_simple_text(t3.rows[14].cells[9], "None", font_size_pt=7.0)

# Row 15: Weekly Surface
set_cell_simple_text(t3.rows[15].cells[2], "12Oct2026", font_size_pt=7.0)
set_cell_simple_text(t3.rows[15].cells[3], "SMO", font_size_pt=7.0)
set_cell_simple_text(t3.rows[15].cells[4], "No Growth", font_size_pt=7.0)
set_cell_simple_text(t3.rows[15].cells[5], "Not Applicable", font_size_pt=7.0)
set_cell_simple_text(t3.rows[15].cells[7], "Not Applicable", font_size_pt=7.0)
set_cell_simple_text(t3.rows[15].cells[9], "None", font_size_pt=7.0)

doc.save(TARGET_DOCX)
print(f"Saved DOCX to: {TARGET_DOCX}")

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
try:
    wdoc = word.Documents.Open(os.path.abspath(TARGET_DOCX))
    wdoc.SaveAs(os.path.abspath(TARGET_PDF), FileFormat=17)
    wdoc.Close()
finally:
    word.Quit()

print(f"Exported PDF to: {TARGET_PDF}")

pdf = fitz.open(TARGET_PDF)
print(f"VERIFIED STANDALONE TABLES PDF PAGE COUNT: {len(pdf)} pages.")
for pno in range(len(pdf)):
    txt = pdf[pno].get_text()[:60].strip().replace('\n', ' ')
    links = list(pdf[pno].get_links())
    print(f"  Page {pno+1}: \"{txt}\", links={len(links)}")

assert len(pdf) == 3, f"Expected strictly 3 pages, got {len(pdf)}"
pdf.close()
print("SUCCESS: 3-page standalone tables verified!")
