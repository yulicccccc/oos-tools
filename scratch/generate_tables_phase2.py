import os, sys, shutil, datetime
import docx
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
from docxtpl import DocxTemplate
import win32com.client
from pypdf import PdfReader
import fitz

OOS_ROOT = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
sys.path.insert(0, OOS_ROOT)
os.chdir(OOS_ROOT)

OUTPUT_DIR = os.path.join(OOS_ROOT, "scratch")
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"

# Base template
TPL_BASE = os.path.join(OUTPUT_DIR, "Tables OOS-262098 Solyn LLC (E75000) - USP71.docx")
if not os.path.exists(TPL_BASE):
    TPL_BASE = os.path.join(DESKTOP_DIR, "Tables OOS-262098 Solyn LLC (E75000) - USP71.docx")

doc = docx.Document(TPL_BASE)
t1 = doc.tables[0]

# Add row for retest in Table 1
r_new = t1.add_row()

# Define URLs
URL_INIT = "https://etrax.eagleanalytical.com/Submission/Details/%247Ydn%24JOiYgmsnaKjbxn-g__"
URL_INIT_ID = "https://etrax.eagleanalytical.com/Submission/Details?id=oZa2TD1UI-0eAJOFysBJ3A__"
URL_RETEST = "https://etrax.eagleanalytical.com/Submission/Details/18rFOYVWqrO0Afk1OwB2hg__"
URL_RETEST_ID = "https://etrax.eagleanalytical.com/Submission/Details/7K%24JgQe6bduNBFY3pBa03w__"

def set_cell_clean_text(cell, text, font_size_pt=7, italic_match=None):
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
        f'</w:p>'
    )
    new_p = parse_xml(p_xml)
    tc.append(new_p)
    
    # split text by newlines
    lines = text.split('\n')
    for l_idx, line in enumerate(lines):
        if l_idx > 0:
            br_r = parse_xml(
                f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:br/></w:r>'
            )
            new_p.append(br_r)
        
        is_it = False
        if italic_match and italic_match.lower() in line.lower():
            is_it = True
        
        it_xml = '<w:i/><w:iCs/>' if is_it else ''
        r_xml = (
            f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:rPr>'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'{it_xml}'
            f'<w:sz w:val="{int(font_size_pt*2)}"/>'
            f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
            f'</w:rPr>'
            f'<w:t>{line}</w:t>'
            f'</w:r>'
        )
        new_p.append(parse_xml(r_xml))

def set_cell_hyperlink(cell, display_text, url, font_size_pt=7):
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    part = cell.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hl_xml = (
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
        f'<w:hyperlink r:id="{r_id}">'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:color w:val="0000FF"/>'
        f'<w:u w:val="single"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'<w:t>{display_text}</w:t>'
        f'</w:r>'
        f'</w:hyperlink>'
        f'</w:p>'
    )
    tc.append(parse_xml(hl_xml))

# Populate Row 2 (Retest) in Table 1
cells_r2 = r_new.cells
set_cell_clean_text(cells_r2[0], "Abayomi Odugbesi", font_size_pt=7)
set_cell_clean_text(cells_r2[1], "Andrew Carrillo", font_size_pt=7)
set_cell_hyperlink(cells_r2[2], "ETX-260914-0470", URL_RETEST, font_size_pt=7)
set_cell_hyperlink(cells_r2[3], "ETX-260921-0498", URL_RETEST_ID, font_size_pt=7)
set_cell_clean_text(cells_r2[4], "1 x 100mL TSB bottle", font_size_pt=7)
set_cell_clean_text(cells_r2[5], "Microbacterium sp.\n(Gram (+) rods)", font_size_pt=7, italic_match="Microbacterium")

out_table_docx = os.path.join(OUTPUT_DIR, "Tables OOS-262098 Solyn LLC (E75000) - Phase II.docx")
out_table_pdf = os.path.join(OUTPUT_DIR, "Tables OOS-262098 Solyn LLC (E75000) - Phase II.pdf")
doc.save(out_table_docx)
print("Saved Phase II Table docx:", out_table_docx)

# Convert Table docx to pdf
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc_com = word.Documents.Open(os.path.abspath(out_table_docx))
    doc_com.SaveAs(os.path.abspath(out_table_pdf), FileFormat=17)
    doc_com.Close()
finally:
    word.Quit()
print("Saved Phase II Table pdf:", out_table_pdf)
