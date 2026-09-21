import os
import win32clipboard
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docx_out = os.path.join(desktop_dir, "OOS-261185 Review Email.docx")
html_out = os.path.join(desktop_dir, "OOS-261185_Review_Email_Robin_Format.html")

# --- 1. GENERATE WORD DOCX (EXACT 3-COLUMN ROBIN SEYMOUR FORMAT) ---
doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

# Greeting
p1 = doc.add_paragraph()
p1.paragraph_format.left_indent = Inches(0)
p1.paragraph_format.space_before = Pt(0)
p1.paragraph_format.space_after = Pt(6)
r1 = p1.add_run("Good morning @Simin Mohammad,")
r1.font.name = "Calibri"
r1.font.size = Pt(11)

p2 = doc.add_paragraph()
p2.paragraph_format.left_indent = Inches(0)
p2.paragraph_format.space_before = Pt(0)
p2.paragraph_format.space_after = Pt(12)
r2 = p2.add_run(
    "I have reviewed this OOS and made my edits to this and summary of the major ones as below. "
    "Please also note that the Highlighted sections in the table will need to be amended:"
)
r2.font.name = "Calibri"
r2.font.size = Pt(11)

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

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
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

# EXACT 3 COLUMNS: Section | Unreviewed Version | Reviewed Version
table_data = [
    ["Section", "Unreviewed Version", "Reviewed Version"],
    [
        "EM Table Attachment & Formatting",
        "The EM Table was omitted from the PDF submission. In the draft Word document, Table 1 listed ETX ID as **ETX-260526-0461** and Microbial ID as **Talaromyces purpurogenus**, Table 2 contained unclosed parentheses in timing descriptions, and multi-organism entries ran together without line breaks.",
        "Amended both tables to reflect the correct ETX submission ID **ETX-260518-0250**, updated the microbial identification to **Penicillium decumbens, Cladosporium tenuissimum, Cladosporium langeronii, Cladosporium halotolerans**, corrected all timing descriptions, formatted multi-organism line breaks, and attached the amended 1-page table as **Page 7** to complete the 7-page package."
    ],
    [
        "Initiator Name",
        "Listed as **\"Simin Mohammad (written by Simin Mohammad)\"**",
        "Revised to **\"Simin Mohammad\"** to remove redundant parenthetical phrasing."
    ],
    [
        "Name of Analyst who Performed the Test",
        "Listed as **\"Simin Mohammad(weekly Active air Sampling Plate Setup)Maraya Chukwumerije & Sophia Santamaria(Weekly active air Sampling Plate Readers)\"**",
        "Revised to separate entries with proper role titles and line breaks:\n**Simin Mohammad (Weekly Active Air Sampling Plate Setup)**\n**Maraya Chukwumerije (Weekly Active Air Sampling Plate Reader)**\n**Sophia Santamaria (Weekly Active Air Sampling Plate Reader)**"
    ],
    [
        "SOP Reference & Effective Date",
        "Listed as **\"20600.002\"** Rev **15**, Effective Date **05AUG2025**",
        "Revised to **\"MICRO-SOP-2\"** Rev **16**, Effective Date **23-Jul-2026**"
    ],
    [
        "Limits / Specification",
        "Listed as **\"Action level >10\"**",
        "Revised to **\"Action level: >= 10 CFU/Plate\"**"
    ],
    [
        "Section B SOP Citations",
        "Comments listed **\"SOP 2.600.00\"** missing trailing digit",
        "Revised to **\"MICRO-SOP-2\"**"
    ],
    [
        "Analyst Interview Comment",
        "Stated **\"Yes, analysts Simin Mohammad, Maraya Chukwumerije, & Sophia Santamaria were comprehensively interviewed.\"**",
        "Revised ampersand to \"and\": **\"Yes, analysts Simin Mohammad, Maraya Chukwumerije, and Sophia Santamaria were comprehensively interviewed.\"**"
    ],
    [
        "Equipment & Calibration Details",
        "Equipment IDs and calibration dates running together without spacing:\n**\"Incubator E001034(Sensor E001501)Incubator E001031 (Sensor E001505)\"**\n**\"Aug 2026 Feb 2027  Aug 2026 Feb 2027\"**",
        "Revised with clean line breaks:\n**Incubator E001034 (Sensor E001501)**\n**Incubator E001031 (Sensor E001505)**\n\n**Aug 2026 / Feb 2027**"
    ],
    [
        "Phase I Summary – Setup & Incubation Narrative",
        "Contained spelling error in reader's name (**\"Sophia Sanatamaria\"**), grammatical error (**\"incubators were verified\"**), and inaccurate reference to biosafety cabinet locations instead of cleanroom suites.",
        "Corrected name to **Sophia Santamaria**, corrected grammar to **incubator was verified**, updated locations to **Suite 115 (ISO 8), Suite 115A (ISO 7), and Suite 115B (ISO 7)**, and cited **MICRO-SOP-2** and **MICRO-SOP-9**."
    ],
    [
        "Phase I Summary – Monthly Cleaning & Disinfection Narrative",
        "Contained contradictory monthly cleaning dates of **\"22 Feb 2026\"** and **\"26APR2026\"** within the same sentence, a missing period, and legacy SOP citations.",
        "Reconciled cleaning date to **26-Apr-2026** performed by Tamiru Kotisso and Cuong Du, inserted missing period before the H2O2 indicators statement, and updated procedure reference to **MICRO-SOP-9**."
    ],
    [
        "Phase I Summary – Root Cause Statement",
        "Stated that the OOS result observed for the Environmental Monitoring (EM) **\"Settling Sampling plate\"** may be attributed to a potential analyst error.",
        "Revised plate type to **Active Air Sampling plate** to accurately reflect the test performed, while maintaining the determination that the result may be attributed to a potential analyst error, noting that the contamination was transient in nature and no corrective actions are necessary."
    ]
]

col_widths = [1.8, 2.5, 2.7]
table = doc.add_table(rows=len(table_data), cols=3)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False

tblPr = table._element.xpath('w:tblPr')
if tblPr:
    for ind in tblPr[0].xpath('w:tblInd'):
        tblPr[0].remove(ind)
    tblPr[0].append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="0" w:type="dxa"/>'))

def add_formatted_runs(paragraph, text, align=WD_ALIGN_PARAGRAPH.LEFT):
    paragraph.alignment = align
    # Parse **bold** markers
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
        set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
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

# Closing
p_close = doc.add_paragraph()
p_close.paragraph_format.space_before = Pt(14)
p_close.paragraph_format.space_after = Pt(4)
r_close = p_close.add_run("@Simin Mohammad please review and sign the attached finalized 7-page PDF package.")
r_close.font.name = "Calibri"
r_close.font.size = Pt(11)

p_sig = doc.add_paragraph()
p_sig.paragraph_format.space_before = Pt(6)
p_sig.paragraph_format.space_after = Pt(0)
p_sig.paragraph_format.line_spacing = 1.15
sig_text = (
    ".\n"
    "Thanks,\n"
    "Robin Seymour\n"
    "Microbiology Supervisor (Sterile Lab)\n"
    "Eagle Analytical Services\n"
    "rseymour@eagleanalytical.com\n"
    "11111 S. Wilcrest Dr. # S1000\n"
    "Houston, Texas 77099"
)
r_sig = p_sig.add_run(sig_text)
r_sig.font.name = "Calibri"
r_sig.font.size = Pt(10)

doc.save(docx_out)
print(f"Saved Word review email: {docx_out}")

# --- 2. GENERATE HTML FILE (EXACT 3-COLUMN ROBIN FORMAT) ---
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
  table {{ border-collapse: collapse; width: 100%; font-family: Calibri, Arial, sans-serif; font-size: 10pt; margin: 12px 0 16px 0; }}
  th {{ background-color: #FFFF00 !important; color: #000; font-weight: bold; text-align: center; border: 1px solid #000000; padding: 8px 10px; }}
  td {{ border: 1px solid #000000; padding: 8px 10px; }}
</style>
</head>
<body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000; line-height: 1.4;">
<!--StartFragment-->
<p style="margin: 0 0 8px 0; font-family: Calibri, sans-serif; font-size: 11pt;">Good morning @Simin Mohammad,</p>
<p style="margin: 0 0 12px 0; font-family: Calibri, sans-serif; font-size: 11pt;">I have reviewed this OOS and made my edits to this and summary of the major ones as below. Please also note that the Highlighted sections in the table will need to be amended:</p>
<table border="1" bordercolor="#000000" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Calibri, sans-serif; font-size: 10pt; margin: 12px 0 16px 0;">
  <thead>
    <tr>
      <th bgcolor="#FFFF00" style="background-color: #FFFF00 !important; width: 25%; color: #000; font-weight: bold; text-align: center; border: 1px solid #000000; padding: 8px 10px;">Section</th>
      <th bgcolor="#FFFF00" style="background-color: #FFFF00 !important; width: 35%; color: #000; font-weight: bold; text-align: center; border: 1px solid #000000; padding: 8px 10px;">Unreviewed Version</th>
      <th bgcolor="#FFFF00" style="background-color: #FFFF00 !important; width: 40%; color: #000; font-weight: bold; text-align: center; border: 1px solid #000000; padding: 8px 10px;">Reviewed Version</th>
    </tr>
  </thead>
  <tbody>
    {html_rows}
  </tbody>
</table>
<p style="margin: 14px 0 6px 0; font-family: Calibri, sans-serif; font-size: 11pt;">@Simin Mohammad please review and sign the attached finalized 7-page PDF package.</p>
<p style="margin: 6px 0 0 0; font-family: Calibri, sans-serif; font-size: 10pt; line-height: 1.2;">
.<br>
Thanks,<br>
Robin Seymour<br>
Microbiology Supervisor (Sterile Lab)<br>
Eagle Analytical Services<br>
rseymour@eagleanalytical.com<br>
11111 S. Wilcrest Dr. # S1000<br>
Houston, Texas 77099
</p>
<!--EndFragment-->
</body>
</html>
"""

with open(html_out, "w", encoding="utf-8") as f:
    f.write(full_html)
print(f"Saved clean HTML: {html_out}")

# --- 3. COPY TO WINDOWS CLIPBOARD DIRECTLY ---
header_tmpl = 'Version:0.9\r\nStartHTML:{:08d}\r\nEndHTML:{:08d}\r\nStartFragment:{:08d}\r\nEndFragment:{:08d}\r\n'
dummy = header_tmpl.format(0, 0, 0, 0)
header_len = len(dummy.encode('utf-8'))

start_html = header_len
end_html = header_len + len(full_html.encode('utf-8'))

start_frag_idx = full_html.find('<!--StartFragment-->') + len('<!--StartFragment-->')
start_frag = header_len + len(full_html[:start_frag_idx].encode('utf-8'))

end_frag_idx = full_html.find('<!--EndFragment-->')
end_frag = header_len + len(full_html[:end_frag_idx].encode('utf-8'))

final_header = header_tmpl.format(start_html, end_html, start_frag, end_frag)
payload = final_header.encode('utf-8') + full_html.encode('utf-8')

plain_text = f"""Good morning @Simin Mohammad,

I have reviewed this OOS and made my edits to this and summary of the major ones as below. Please also note that the Highlighted sections in the table will need to be amended:

Section | Unreviewed Version | Reviewed Version
---------------------------------------------------------------------------------------------------------
"""
for r in table_data[1:]:
    plain_text += f"{r[0]}:\n  Unreviewed Version: {r[1]}\n  Reviewed Version: {r[2]}\n\n"

plain_text += (
    "@Simin Mohammad please review and sign the attached finalized 7-page PDF package.\n\n"
    ".\n"
    "Thanks,\n"
    "Robin Seymour\n"
    "Microbiology Supervisor (Sterile Lab)\n"
    "Eagle Analytical Services\n"
    "rseymour@eagleanalytical.com\n"
    "11111 S. Wilcrest Dr. # S1000\n"
    "Houston, Texas 77099\n"
)

win32clipboard.OpenClipboard()
try:
    win32clipboard.EmptyClipboard()
    CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
    win32clipboard.SetClipboardData(CF_HTML, payload)
    win32clipboard.SetClipboardText(plain_text)
    print("SUCCESS: Copied Robin Seymour exact 3-column email directly into Windows Clipboard!")
finally:
    win32clipboard.CloseClipboard()
