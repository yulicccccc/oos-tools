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

DOCX_PATH = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx"
PDF_PATH = os.path.join(r"scratch", "temp_tables_oos262333_test.pdf")

doc = docx.Document(DOCX_PATH)
t3 = doc.tables[2]
cell_plate = t3.rows[13].cells[6]

# Clear cell
tc = cell_plate._tc
for p in tc.xpath('w:p'):
    tc.remove(p)

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
print("Saved Word doc with plate hyperlink!")

# Check page count
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
print(f"Page count: {len(p_doc)}")
p8_links = p_doc[1].get_links()
print(f"Page 2 (Table 3) links: {len(p8_links)}")
for l in p8_links:
    print("  ", l.get("uri"))
p_doc.close()
