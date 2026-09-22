import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client as win32

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docx_path = os.path.join(desktop_dir, "OOS-261242 Review Email.docx")
html_path = os.path.join(desktop_dir, "OOS-261242_Review_Email_Robin_Format.html")

# --- 1. REBUILD CLEAN DOCX ---
doc = Document()
for s in doc.sections:
    s.top_margin = Inches(0.8)
    s.bottom_margin = Inches(0.8)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)

# Greeting
p_g = doc.add_paragraph()
r_g = p_g.add_run("Good morning @Simin Mohammad,")
r_g.font.name = "Calibri"
r_g.font.size = Pt(11)
p_g.paragraph_format.space_before = Pt(0)
p_g.paragraph_format.space_after = Pt(6)

# Intro
p_i = doc.add_paragraph()
r_i = p_i.add_run("I have reviewed OOS-261242. Please find the following edits made in the summary table below:")
r_i.font.name = "Calibri"
r_i.font.size = Pt(11)
p_i.paragraph_format.space_before = Pt(0)
p_i.paragraph_format.space_after = Pt(12)

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('shd'):
            tcPr.remove(child)
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_borders(cell):
    tcPr = cell._element.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcBorders'):
            tcPr.remove(child)
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcMar'):
            tcPr.remove(child)
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

table_data = [
    ["Section", "Unreviewed Version", "Reviewed Version"],
    [
        "Equipment/Facility Description",
        "Only lists **CR115 (E001737)** in equipment section. In Section D, incubator IDs and calibration dates ran together without proper line spacing: **\"Incubator E001034(Sensor E001501)Incubator E001031 (Sensor E001505)\"** and **\"Aug-2026 Feb-2027 Aug-2026 Feb-2027\"**.",
        "Added full cleanroom facility identifiers for **Suite 115 (ISO 8 Anteroom)**, **Suite 115A (ISO 7 Buffer room)**, and **Suite 115B (ISO 7 Cleanroom)** with Sensor E001737. Formatted **Incubator E001034 (Sensor E001501)** and **Incubator E001031 (Sensor E001505)** with clean line breaks and calibration dates (**Aug 2026 / Feb 2027**)."
    ],
    [
        "Cleanroom Suite Description",
        "Inconsistently mixed cleanroom suite nomenclature with **\"CR115\"**, **\"ISO 8 115, and ISO 7 115A, and ISO 7 115B\"**, and capitalized **\"The CR115 was thoroughly cleaned\"** in the narrative.",
        "Standardized to clean **Suite 115 (ISO 8)**, **Suite 115A (ISO 7)**, and **Suite 115B (ISO 7)** cleanroom suite nomenclature across all narrative sections and corrected capitalization."
    ],
    [
        "Reading Analyst, Processing Analyst & SOP Description",
        "Combined analyst names on a single line: **\"Simin Mohammad(Weekly Active air Sampling Plate Setup)Sophia Santamaria(Weekly Active air Sampling Plate Readers\"** (missing closing parenthesis and plural typo). Cited legacy **\"MICRO-SOP-2\"** Rev **15** (05-AUG-2025). Stated **\"analysts Simin Mohammad, & Sophia Santamaria\"** with redundant ampersand. In Section D, misspelled reader as **\"Sophia Sanatamaria\"**, contained grammar error **\"incubators were verified\"**, and erroneously cited **\"MICRO-SOP-2 (Cleaning and Disinfecting Procedure for Microbiology)\"**.",
        "Delineated setup analyst **Simin Mohammad (Weekly Active Air Sampling Plate Setup)** and reader **Sophia Santamaria (Weekly Active Air Sampling Plate Reader)**. Updated procedure to current **MICRO-SOP-2** Rev **16** (Effective Date: **23-Jul-2026**). Revised interview comment to **\"Simin Mohammad and Sophia Santamaria\"**. Corrected spelling to **Sophia Santamaria**, grammar to **incubator was verified**, and corrected cleaning procedure reference to **MICRO-SOP-9**."
    ],
    [
        "Limits / Specification",
        "Listed as **\"Action level: >10CFU/plate\"** without proper spacing or cGMP equality symbol.",
        "Standardized to **\"Action level: >= 10 CFU/Plate\"**."
    ],
    [
        "EM hits, Cleanroom Bracketing & BSC Hits Assessment",
        "Contained duplicate date in monthly cleaning: **\"performed on 31-May-2026 as per MICRO-SOP-9 ... on 31-May-2026 It was documented that all H₂O₂ indicators passed\"** (missing period). Note on concurrent sterility test hits contained grammatical errors (**\"there was two celsis sterility test sample that was tested positive\"**, lowercase **\"on 11-May-2026\"**), and lacked explicit taxonomic differentiation explaining why these hits are unrelated to the active air OOS.",
        "Reconciled monthly cleaning sentence to remove duplicate date, inserted period before **\"It was documented that all H₂O₂ indicators passed.\"** Corrected grammar to **\"there were two sterility test samples that tested positive\"**, capitalized sentence start, and explicitly detailed that the recoveries in BSC E001314 (**Staphylococcus lugdunensis** and **Cladosporium halotolerans**) exhibit distinct taxonomy from the OOS recovery (**Candida orthopsilosis**), confirming no common source or cleanroom facility contamination."
    ],
    [
        "Phase I Summary – Root Cause Statement & Continuation Pages",
        "The Phase I Root Cause Statement was placed at the bottom of Page 4, leaving Page 5 marked as **\"N/A SMO 18-AUG-2026\"**, causing Page 4 to be overcrowded while leaving Page 5 blank.",
        "Balanced narrative across continuation pages: Page 4 details bracketing summary, analyst interview, and monthly cleaning; Page 5 houses the evaluation of concurrent cleanroom BSC samples, taxonomic differentiation, and the finalized Phase I Root Cause Statement confirming that the recovery represents an isolated, transient environmental event with effective routine disinfection."
    ],
    [
        "EM Table Attachment & Formatting",
        "The EM Table was omitted from the PDF submission. In the draft Word document, Table 1 date had excessive spacing (**\"14MAY   2026\"**), CFU count lacked spacing (**\"11CFUs\"**), Table 2 Row 7 had plural typo (**\"1 CFUs on table  in 115 ISO 8 and 1 CFUs on cart in 115 ISO 8\"**), and multi-organism entries lacked line breaks.",
        "Amended both tables to correct spacing (**14MAY 2026**, **11 CFUs**), corrected plural typos (**1 CFU on table in 115 ISO 8 and 1 CFU on cart in 115 ISO 8**), formatted multi-organism line breaks, standardized bracketing section headers, exported as a pixel-perfect 1-page PDF, and attached as **Page 8** to complete the finalized 8-page package."
    ]
]

col_widths = [1.8, 2.6, 2.8]
table = doc.add_table(rows=len(table_data), cols=3)
table.alignment = WD_TABLE_ALIGNMENT.LEFT
table.autofit = False

tblPr = table._element.xpath('w:tblPr')
if tblPr:
    for jc in tblPr[0].xpath('w:jc'):
        tblPr[0].remove(jc)
    tblPr[0].append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
    
    for ind in tblPr[0].xpath('w:tblInd'):
        tblPr[0].remove(ind)
    tblPr[0].append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="0" w:type="dxa"/>'))

def add_formatted_runs(paragraph, text, align=WD_ALIGN_PARAGRAPH.LEFT):
    paragraph.alignment = align
    parts = text.split('**')
    for idx, part in enumerate(parts):
        if not part:
            continue
        is_bold = (idx % 2 == 1)
        r = paragraph.add_run(part)
        r.font.name = "Calibri"
        r.font.size = Pt(10)
        if is_bold:
            r.bold = True

for r_idx, row_content in enumerate(table_data):
    row = table.rows[r_idx]
    is_header = (r_idx == 0)
    
    for c_idx, cell_text in enumerate(row_content):
        cell = row.cells[c_idx]
        cell.width = Inches(col_widths[c_idx])
        set_cell_borders(cell)
        set_cell_margins(cell, top=120, bottom=120, left=150, right=150)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER if (is_header or c_idx == 0) else WD_ALIGN_VERTICAL.TOP
        
        if is_header:
            set_cell_background(cell, "FFFF00")
        else:
            set_cell_background(cell, "FFFFFF")
            
        p = cell.paragraphs[0]
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        
        if is_header:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(cell_text)
            r.font.name = "Calibri"
            r.font.size = Pt(10.5)
            r.bold = True
            r.font.color.rgb = RGBColor(0, 0, 0)
        else:
            lines = cell_text.split('\n')
            align = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            for l_idx, line in enumerate(lines):
                if l_idx > 0:
                    p = cell.add_paragraph()
                    p.paragraph_format.left_indent = Inches(0)
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.15
                add_formatted_runs(p, line, align=align)

# Exact Closing as requested: "please review and sign the updated OOS form."
p_close = doc.add_paragraph()
p_close.paragraph_format.space_before = Pt(14)
p_close.paragraph_format.space_after = Pt(0)
r_close = p_close.add_run("please review and sign the updated OOS form.")
r_close.font.name = "Calibri"
r_close.font.size = Pt(11)

# NO ROBIN SIGNATURE! End cleanly here.
doc.save(docx_path)
print(f"Saved clean Word email (no Robin signature): {docx_path}")

# --- 2. REBUILD CLEAN HTML ---
def html_bold(text):
    lines = text.split('\n')
    formatted_lines = []
    for line in lines:
        parts = line.split('**')
        out_line = ""
        for idx, part in enumerate(parts):
            if not part:
                continue
            if idx % 2 == 1:
                out_line += f'<b>{part}</b>'
            else:
                out_line += part
        formatted_lines.append(out_line)
    return '<br>'.join(formatted_lines)

html_rows = ""
for r_idx, row in enumerate(table_data[1:]):
    sec, unrev, rev = row
    sec_html = html_bold(sec)
    unrev_html = html_bold(unrev)
    rev_html = html_bold(rev)
    
    html_rows += f"""
    <tr>
      <td style="border: 1px solid #000000; padding: 8px 10px; text-align: center; vertical-align: middle;">{sec_html}</td>
      <td style="border: 1px solid #000000; padding: 8px 10px; vertical-align: top; text-align: left;">{unrev_html}</td>
      <td style="border: 1px solid #000000; padding: 8px 10px; vertical-align: top; text-align: left;">{rev_html}</td>
    </tr>
    """

full_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body {{ font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000; line-height: 1.4; }}
  table {{ border-collapse: collapse; width: 100%; font-family: Calibri, Arial, sans-serif; font-size: 10pt; margin: 12px 0 16px 0; margin-left: 0; margin-right: auto; }}
  th {{ background-color: #FFFF00 !important; color: #000; font-weight: bold; text-align: center; border: 1px solid #000000; padding: 8px 10px; }}
  td {{ border: 1px solid #000000; padding: 8px 10px; }}
</style>
</head>
<body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000; line-height: 1.4;">
<!--StartFragment-->
<p style="margin: 0 0 8px 0; font-family: Calibri, sans-serif; font-size: 11pt;">Good morning @Simin Mohammad,</p>
<p style="margin: 0 0 12px 0; font-family: Calibri, sans-serif; font-size: 11pt;">I have reviewed OOS-261242. Please find the following edits made in the summary table below:</p>
<table border="1" bordercolor="#000000" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Calibri, sans-serif; font-size: 10pt; margin: 12px 0 16px 0; margin-left: 0; margin-right: auto;">
  <thead>
    <tr>
      <th bgcolor="#FFFF00" style="background-color: #FFFF00 !important; width: 25%; color: #000; font-weight: bold; text-align: center; border: 1px solid #000000; padding: 8px 10px;">Section</th>
      <th bgcolor="#FFFF00" style="background-color: #FFFF00 !important; width: 37%; color: #000; font-weight: bold; text-align: center; border: 1px solid #000000; padding: 8px 10px;">Unreviewed Version</th>
      <th bgcolor="#FFFF00" style="background-color: #FFFF00 !important; width: 38%; color: #000; font-weight: bold; text-align: center; border: 1px solid #000000; padding: 8px 10px;">Reviewed Version</th>
    </tr>
  </thead>
  <tbody>
    {html_rows}
  </tbody>
</table>
<p style="margin: 14px 0 0 0; font-family: Calibri, sans-serif; font-size: 11pt;">please review and sign the updated OOS form.</p>
<!--EndFragment-->
</body>
</html>
"""

with open(html_path, "w", encoding="utf-8") as f:
    f.write(full_html)
print(f"Saved clean HTML (no Robin signature): {html_path}")

# --- 3. COPY DIRECTLY TO WINDOWS CLIPBOARD VIA WORD COM ---
word_app = win32.Dispatch("Word.Application")
word_app.Visible = False
try:
    d = word_app.Documents.Open(os.path.abspath(docx_path))
    d.Content.Copy()
    d.Close(False)
    print("SUCCESS: Copied clean email (ending with 'please review and sign the updated OOS form.') to Windows Clipboard via Word COM!")
finally:
    word_app.Quit()
