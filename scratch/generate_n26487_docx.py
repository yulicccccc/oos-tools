import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

doc = Document()

# Set standard 0.5 inch margins
for section in doc.sections:
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

def set_cell_shading(cell, color_hex):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_borders(cell, top="0.5pt", bottom="0.5pt", left="0.5pt", right="0.5pt", color="CCCCCC"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{color}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

# Header Table
tbl_hdr = doc.add_table(rows=2, cols=4)
tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_hdr.autofit = False

hdr_data = [
    [("EAGLE", True, 12, "003366"), ("Category: Form Corporate\nName: Nonconformance Report Form", False, 8.5, "333333"), ("ETX ID:\n3.100.007.F01", False, 8.5, "333333"), ("Effective:\n10-JUL-2026", False, 8.5, "333333")],
    [("CORP-FORM-5 v9.1", True, 9, "003366"), ("Location: Houston (Corporate)", False, 8.5, "333333"), ("Printed from https://app.zenqms.com", False, 8, "777777"), ("NCR#: N 26487", True, 11, "003366")]
]

for r_idx, row in enumerate(tbl_hdr.rows):
    for c_idx, cell in enumerate(row.cells):
        set_cell_borders(cell, color="AAAAAA")
        set_cell_shading(cell, "F8F9FA")
        txt, bold, sz, col = hdr_data[r_idx][c_idx]
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(txt)
        run.bold = bold
        run.font.size = Pt(sz)
        run.font.color.rgb = RGBColor.from_string(col)

doc.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 1: IDENTIFICATION
p_sec1 = doc.add_paragraph()
r1 = p_sec1.add_run("IDENTIFICATION")
r1.bold = True
r1.font.size = Pt(10)
r1.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
p_sec1.paragraph_format.space_before = Pt(4)
p_sec1.paragraph_format.space_after = Pt(2)

tbl_id = doc.add_table(rows=4, cols=4)
tbl_id.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_id.autofit = False

id_data = [
    [("Initiator:", True), ("Qiyue Chen", False), ("Date:", True), ("18-Sep-2026", False)],
    [("Department:", True), ("Microbiology", False), ("Equipment ID #:", True), ("E001230 (Scan RDI)", False)],
    [("Sample ID #:", True), ("ETX-260915-0653\nETX-260916-0044\nETX-260916-0374", False), ("Sample Name / Description:", True), ("Semaglutide/Cyanocobalamin\nTirzepatide 22mg/mL-Pyridoxine HCl 4mg/mL\nSEMAGLUTIDE/ B12 1.2 MG/ 500 MCG/ML Inj", False)],
    [("Lot #:", True), ("0216202602\n22666\nLG342010613", False), ("Client Info:", True), ("ANG Labs (E73555)\nSouthend Pharmacy (E19207)\nOptimal Balance Pharmacy (E19193)", False)]
]

for r_idx, row in enumerate(tbl_id.rows):
    for c_idx, cell in enumerate(row.cells):
        set_cell_borders(cell, color="CBD5E0")
        txt, bold = id_data[r_idx][c_idx]
        if bold:
            set_cell_shading(cell, "EDF2F7")
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(txt)
        run.bold = bold
        run.font.size = Pt(8.5)

doc.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 2: NONCONFORMITY DESCRIPTION
p_sec2 = doc.add_paragraph()
r2 = p_sec2.add_run("NONCONFORMITY DESCRIPTION (Who, What, When, Where, Why, & How)")
r2.bold = True
r2.font.size = Pt(10)
r2.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
p_sec2.paragraph_format.space_before = Pt(4)
p_sec2.paragraph_format.space_after = Pt(2)

tbl_desc = doc.add_table(rows=1, cols=1)
tbl_desc.alignment = WD_TABLE_ALIGNMENT.CENTER
cell_d = tbl_desc.rows[0].cells[0]
set_cell_borders(cell_d, color="CBD5E0")
set_cell_shading(cell_d, "FFFFFF")

desc_paragraphs = [
    "On 17-Sep-2026, during routine Scan RDI sterility testing operations, an unexpected software crash occurred on the Scan RDI instrument (Equipment ID: E001230) while running session '17 Sep 2026 - 3'. Third-shift analyst JOC immediately notified Project Microbiologist QYC regarding the incident and software failure. JOC attempted to re-open the application and recover the testing session data; however, the instrument repeatedly displayed error messages stating, 'Unable to Create New Session Unknown Error' and 'Unable to Load Session Unknown Error', preventing access to the session.",
    "The following day, on 18-Sep-2026, QYC attempted to re-open the application and recover the testing session data again; the session '17 Sep 2026 - 3' could be opened, but no data was present within it. Work Order WO-260470 was promptly issued to Engineering/IT by RS to formally investigate the software failure and initiate technical remediation.",
    "On 21-Sep-2026, QYC proactively reached out to Eagle IT Coordinator KT (who was previously unaware of the incident) and attempted to restore and retrieve the session data. Following technical review, KT confirmed that the session file was corrupted, exporting as only 7 KB, and the corresponding raw acquisition dataset within the D:\\ directory was completely empty and unrecoverable.",
    "Three ScanRDI test sample submissions, ETX-260915-0653, ETX-260916-0044 and ETX-260916-0374 were actively processed in this corrupted session and impacted by data loss. No valid test results could be generated or verified for these submissions from the corrupted run."
]

for idx, txt in enumerate(desc_paragraphs):
    p = cell_d.paragraphs[0] if idx == 0 else cell_d.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(txt)
    run.font.size = Pt(8.5)

doc.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 3: CAUSE(S) ANALYSIS
p_sec3 = doc.add_paragraph()
r3 = p_sec3.add_run("CAUSE(S) ANALYSIS")
r3.bold = True
r3.font.size = Pt(10)
r3.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
p_sec3.paragraph_format.space_before = Pt(4)
p_sec3.paragraph_format.space_after = Pt(2)

tbl_cause = doc.add_table(rows=2, cols=1)
tbl_cause.alignment = WD_TABLE_ALIGNMENT.CENTER
c_c0 = tbl_cause.rows[0].cells[0]
set_cell_borders(c_c0, color="CBD5E0")
p = c_c0.paragraphs[0]
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(2)
r = p.add_run("Description of Cause(s): ")
r.bold = True
r.font.size = Pt(8.5)
r_val = p.add_run("Scan RDI software crash and session file corruption on instrument E001230, resulting in unrecoverable raw acquisition data in the D:\\ directory.")
r_val.font.size = Pt(8.5)

c_c1 = tbl_cause.rows[1].cells[0]
set_cell_borders(c_c1, color="CBD5E0")
set_cell_shading(c_c1, "F8F9FA")
p = c_c1.paragraphs[0]
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(2)
p.add_run("Category:  [X] Equipment/System    [ ] Process/Method    [ ] Personnel    [ ] External Phenomena    [ ] Other: N/A QYC 22Sep26").font.size = Pt(8.5)

doc.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 4: INTERIM CORRECTIVE ACTION(S)
p_sec4 = doc.add_paragraph()
r4 = p_sec4.add_run("INTERIM CORRECTIVE ACTION(S)")
r4.bold = True
r4.font.size = Pt(10)
r4.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
p_sec4.paragraph_format.space_before = Pt(4)
p_sec4.paragraph_format.space_after = Pt(2)

tbl_ca = doc.add_table(rows=1, cols=1)
tbl_ca.alignment = WD_TABLE_ALIGNMENT.CENTER
c_ca = tbl_ca.rows[0].cells[0]
set_cell_borders(c_ca, color="CBD5E0")
ca_text = (
    "Work Order WO-260470 was issued to Engineering and IT to investigate and resolve the software crash, clear corrupted temporary files, and evaluate database integrity on instrument E001230. "
    "Client Care was instructed to request additional sample vials for retesting for ETX-260915-0653, ETX-260916-0044 and ETX-260916-0374."
)
p = c_ca.paragraphs[0]
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(2)
p.add_run(ca_text).font.size = Pt(8.5)

doc.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 5: QUALITY EVALUATION & RISK ASSESSMENT
p_sec5 = doc.add_paragraph()
r5 = p_sec5.add_run("QUALITY EVALUATION & CLOSING")
r5.bold = True
r5.font.size = Pt(10)
r5.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
p_sec5.paragraph_format.space_before = Pt(4)
p_sec5.paragraph_format.space_after = Pt(2)

tbl_qe = doc.add_table(rows=3, cols=2)
tbl_qe.alignment = WD_TABLE_ALIGNMENT.CENTER
qe_rows = [
    [("Was action effective in addressing root cause?", True), ("[X] Yes    [ ] No    [ ] N/A", False)],
    [("Additional Comments / Risk Assessment:", True), ("Impact is isolated to the 3 sample submissions processed in the corrupted session (ETX-260915-0653, ETX-260916-0044, ETX-260916-0374). No impact to other testing sessions or alternate operational Scan RDI instruments (E002017, E002225). Additional vials requested / retesting initiated. Instrument E001230 remains under observation and repair under WO-260470.", False)],
    [("Impact Scoring Matrix (Quality Unit):", True), ("Frequency: 2  |  Severity: 2  |  Magnitude: 2  -->  Total Risk Score: 8\nImpact Level: [X] Major (3 to 17)    [ ] Low    [ ] Critical", False)]
]

for r_idx, row in enumerate(tbl_qe.rows):
    for c_idx, cell in enumerate(row.cells):
        set_cell_borders(cell, color="CBD5E0")
        txt, bold = qe_rows[r_idx][c_idx]
        if bold:
            set_cell_shading(cell, "EDF2F7")
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(txt)
        run.bold = bold
        run.font.size = Pt(8.5)

doc.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 6: ACKNOWLEDGEMENT & SIGNATURES
tbl_sig = doc.add_table(rows=2, cols=2)
tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
sig_cells = [
    [("Prepared By (Print):\nQiyue Chen, Ph.D. | Project Microbiologist\nDate: 18-Sep-2026", "Signature:\n\n__________________________________"),
     ("Reviewed By (Print):\nRobin Seymour | Microbiology Supervisor\nDate: ____________", "Signature:\n\n__________________________________")]
]

for r_idx, row in enumerate(tbl_sig.rows):
    for c_idx, cell in enumerate(row.cells):
        set_cell_borders(cell, color="CBD5E0")
        set_cell_shading(cell, "F8F9FA")
        txt = sig_cells[0][c_idx] if r_idx == 0 else ""
        if r_idx == 0:
            left_txt, right_txt = txt
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.add_run(left_txt + "\n\n" + right_txt).font.size = Pt(8.5)

# Save
docx_docs = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\N26487 - Scan RDI Instrument E001230 Software Crash.docx"
docx_oos = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\N26487 - Scan RDI Instrument E001230 Software Crash.docx"

doc.save(docx_docs)
doc.save(docx_oos)
print("Saved Word Document to:")
print(" -", docx_docs)
print(" -", docx_oos)
