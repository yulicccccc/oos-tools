# -*- coding: utf-8 -*-
import os
import sys
import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client
import fitz

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCX_PATH = os.path.join(DESKTOP_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx")
PDF_PATH = os.path.join(r"scratch", "temp_tables_oos262333.pdf")

doc = docx.Document(DOCX_PATH)
t3 = doc.tables[2]
row13 = t3.rows[13]

def set_cell_multiline(cell, lines, font_size_pt=7.0, italic=True):
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
        
    for i, line in enumerate(lines):
        rPr_italic = '<w:i/>' if italic else ''
        p_xml = (
            f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:pPr>'
            f'<w:jc w:val="center"/>'
            f'<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:rPr>'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'{rPr_italic}'
            f'<w:sz w:val="{int(font_size_pt*2)}"/>'
            f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
            f'</w:rPr>'
            f'</w:pPr>'
            f'<w:r>'
            f'<w:rPr>'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'{rPr_italic}'
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

species_list = [
    "Moraxella osloensis",
    "Sporisorium graminicola",
    "Staphylococcus haemolyticus",
    "Staphylococcus hominis",
    "Kocuria indica",
    "Ustilago maydis",
    "Kocuria rhizophila/ sp. BT304"
]

set_cell_multiline(row13.cells[7], species_list, font_size_pt=7.0, italic=True)

# Also ensure plate hyperlink is intact
cell_plate = row13.cells[6]
tc_plate = cell_plate._tc
hlinks = tc_plate.xpath('.//w:hyperlink')
if not hlinks:
    for p in tc_plate.xpath('w:p'):
        tc_plate.remove(p)
    p = cell_plate.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    part = doc.part
    url = 'https://etrax.eagleanalytical.com/Submission/Details/SPTrZYPqECZmMHXaxXbl9Q__'
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w")} r:id="{r_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="14"/><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t>ETX-261021-0578</w:t></w:r>')
    hyperlink.append(new_run)
    p._p.append(hyperlink)
    cell_plate.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER

doc.save(DOCX_PATH)
print("Updated Desktop DOCX with finalized Microbial ID species list successfully.")

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
print(f"Exported PDF page count: {page_count} (Target: strictly 2 pages)")
assert page_count == 2, f"Expected strictly 2 pages, got {page_count}!"
p_doc.close()
print("Tables Word document strictly 2 pages verified!")
