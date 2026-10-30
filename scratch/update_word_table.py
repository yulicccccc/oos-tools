import os
import sys
import docx
from docx.shared import Pt
from docx.oxml import parse_xml
import win32com.client
import fitz

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCX_PATH = os.path.join(DESKTOP_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx")
PDF_PATH = os.path.join(r"scratch", "temp_tables_oos262333.pdf")

doc = docx.Document(DOCX_PATH)
t3 = doc.tables[2]
row13 = t3.rows[13]

def set_cell_multiline(cell, lines, font_size_pt=7.0):
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
        
    for i, line in enumerate(lines):
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
            f'<w:t>{line}</w:t>'
            f'</w:r>'
            f'</w:p>'
        )
        tc.append(parse_xml(p_xml))
        if i < len(lines) - 1:
            p_empty = (
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
                f'</w:p>'
            )
            tc.append(parse_xml(p_empty))
            
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER

# Update Observation
cell_obs = row13.cells[5]
tc_obs = cell_obs._tc
for p in tc_obs.xpath('w:p'):
    tc_obs.remove(p)
p_obs = (
    f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    f'<w:pPr>'
    f'<w:jc w:val="center"/>'
    f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
    f'<w:rPr>'
    f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    f'<w:sz w:val="14"/>'
    f'<w:szCs w:val="14"/>'
    f'</w:rPr>'
    f'</w:pPr>'
    f'<w:r>'
    f'<w:rPr>'
    f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    f'<w:sz w:val="14"/>'
    f'<w:szCs w:val="14"/>'
    f'</w:rPr>'
    f'<w:t>6 CFUs (ISO 8 114)</w:t>'
    f'</w:r>'
    f'</w:p>'
)
tc_obs.append(parse_xml(p_obs))
cell_obs.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER

# Update Microbial ID
micro_lines = [
    "Gram (+) cocci (3 CFUs)",
    "Gram-variable coccobacilli (2 CFUs)",
    "Budding yeast (1 CFU)",
    "(Pending Microbial ID)"
]
set_cell_multiline(row13.cells[7], micro_lines, font_size_pt=7.0)

doc.save(DOCX_PATH)
print("Updated Desktop DOCX successfully.", flush=True)

# Export via Word COM
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
try:
    abs_docx = os.path.abspath(DOCX_PATH)
    abs_pdf = os.path.abspath(PDF_PATH)
    if os.path.exists(abs_pdf):
        os.remove(abs_pdf)
    w_doc = word.Documents.Open(abs_docx, ReadOnly=1)
    w_doc.SaveAs2(abs_pdf, FileFormat=17)
    w_doc.Close(SaveChanges=0)
finally:
    word.Quit()

p_doc = fitz.open(PDF_PATH)
page_count = len(p_doc)
print(f"Exported PDF page count: {page_count} (Target: strictly 2 pages)", flush=True)
assert page_count == 2, f"Expected strictly 2 pages, got {page_count}!"
p_doc.close()
print("Tables Word document strictly 2 pages verified!", flush=True)
