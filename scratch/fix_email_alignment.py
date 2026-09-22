import os
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client as win32

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"

for fname in ["OOS-261242 Review Email.docx", "OOS-261185 Review Email.docx"]:
    docx_path = os.path.join(desktop_dir, fname)
    if not os.path.exists(docx_path):
        continue
    doc = Document(docx_path)
    tbl = doc.tables[0]
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT

    tblPr = tbl._element.xpath('w:tblPr')
    if tblPr:
        for jc in tblPr[0].xpath('w:jc'):
            tblPr[0].remove(jc)
        tblPr[0].append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
        
        for ind in tblPr[0].xpath('w:tblInd'):
            tblPr[0].remove(ind)
        tblPr[0].append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="0" w:type="dxa"/>'))

    doc.save(docx_path)
    print(f"Fixed {fname} to strictly LEFT-ALIGNED (w:val='left', w:tblInd=0)")

# Also update HTML file
html_path = os.path.join(desktop_dir, "OOS-261242_Review_Email_Robin_Format.html")
if os.path.exists(html_path):
    with open(html_path, "r", encoding="utf-8") as f:
        html_content = f.read()
    html_content = html_content.replace(
        'table { border-collapse: collapse; width: 100%;',
        'table { border-collapse: collapse; width: 100%; margin-left: 0; margin-right: auto;'
    )
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print("Fixed HTML table margin to left: 0")

# Copy OOS-261242 Review Email to clipboard via Word COM
target_docx = os.path.join(desktop_dir, "OOS-261242 Review Email.docx")
word_app = win32.Dispatch("Word.Application")
word_app.Visible = False
try:
    d = word_app.Documents.Open(os.path.abspath(target_docx))
    d.Content.Copy()
    d.Close(False)
    print("SUCCESS: Copied LEFT-ALIGNED email to Windows Clipboard via Word COM!")
finally:
    word_app.Quit()
