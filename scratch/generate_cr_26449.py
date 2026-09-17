import os, sys, shutil
from datetime import datetime
import fitz
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT

DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"
DESKTOP_DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\Documents"
OUTPUT_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\scratch"

TEMPLATE_PDF = os.path.join(DESKTOP_DOCS_DIR, "3-100-004-F01-rev10.pdf")
if not os.path.exists(TEMPLATE_PDF):
    TEMPLATE_PDF = r"C:\Users\qchen\.gemini\antigravity\brain\0bf0eb5b-c315-434e-bf5a-249ac4f142ec\media__1783442157468.pdf"

print(f"Using CR Template: {TEMPLATE_PDF}")

# Metadata
cr_id = "CR-26449"
sample_id = "ETX-260805-0189"
transposed_sample_id = "ETX-260804-0101"
client_name = "GoGoMeds Select (E10747)"
sample_name = "Tirzepatide/B12 25/1mg/mL Injectable"
lot_number = "18468"
test_date = "07-Aug-2026"
oos_id = "OOS-261814"
ncr_id = "N26414"
date_requested = datetime.now().strftime("%d-%b-%Y")

# Concise, structured text for Page 1 Reference Box (w=142.8, h=87)
ref_text = (
    f"Sample: {sample_id}\n"
    f"Client: {client_name}\n"
    f"Analyte: {sample_name}\n"
    f"Lot #: {lot_number} | Date: {test_date}\n"
    f"Test: Scan RDI Sterility Testing\n\n"
    f"Related Quality Records:\n"
    f"• {oos_id} (Phase I OOS Investigation)\n"
    f"• NCR {ncr_id} (Sample Mix-up)\n"
    f"• MICRO-SOP-12 | CORP-SOP-15"
)

# Concise, structured text for Page 1 Description Box (w=372, h=87)
desc_text = (
    f"1. Incident & NCR {ncr_id}: On {test_date}, sample {sample_id} failed Scan RDI sterility testing "
    f"(4 CFUs rod morphology). Due to a clerical transposition during data entry, it was inadvertently marked "
    f"and approved as 'Passing/Complete' in EagleTrax, while passing sample {transposed_sample_id} was placed on hold "
    f"(documented under closed NCR {ncr_id}).\n"
    f"2. Client Notification: On {test_date}, Client Care and Lab Management promptly notified the client by phone and email "
    f"to place Lot {lot_number} on quarantine, preventing unintended product release.\n"
    f"3. OOS Investigation: Phase I investigation {oos_id} confirmed the failing result is valid and not due to laboratory "
    f"contamination or analytical error. Original test failure is deemed valid.\n"
    f"4. Action in EagleTrax: Correct result from 'Pass' to 'Fail' (Positive, 4 CFUs, rod morphology) and update final test "
    f"status for {sample_id} in EagleTrax, coupled with closed {oos_id}."
)

# Justification of Impact (placed in the left box under Section 3, avoiding the colored matrix on the right)
just_text = (
    "Justification of Impact:\n"
    f"The transposition occurred during manual data entry on {test_date}. Client was notified immediately on the same day "
    f"to quarantine Lot {lot_number}, preventing release of affected product. Investigation {oos_id} confirmed test validity. "
    f"This CR authorizes correcting EagleTrax records from Pass to Fail. Retraining was completed under closed NCR {ncr_id}. "
    "Overall product quality risk is fully controlled."
)

# 1. Generate Filled PDF Form using fitz (PyMuPDF)
doc = fitz.open(TEMPLATE_PDF)
p1 = doc[0]
p2 = doc[1]

# Page 1 Field Values & Font Sizes
p1_vals = {
    'Text Field34': '26449',
    'Text Field0': 'Qiyue Chen',
    'Text Field1': 'Microbiology',
    'Date Field0': date_requested,
    'Check Box6': 'Yes', # Other:
    'Text Field2': 'EagleTrax Test Result Correction (Pass to Fail)',
    'Check Box8': 'Yes', # OOS #
    'Text Field3': oos_id,
    'Check Box12': 'Yes', # Other
    'Text Field6': f"NCR {ncr_id}",
    'Text Field7': ref_text,
    'Text Field8': desc_text,
    'Text Field9': '1',  # Frequency
    'Text Field10': '3', # Severity
    'Text Field11': '1', # Magnitude
    'Text Field12': '3', # Impact Score
    'Check Box14': 'Yes',# Major (Score 3 to 17)
}

font_sizes_p1 = {
    'Text Field34': 9.0,
    'Text Field0': 8.5,
    'Text Field1': 8.5,
    'Date Field0': 8.5,
    'Text Field2': 7.8,
    'Text Field3': 8.5,
    'Text Field6': 7.5,
    'Text Field7': 6.8,
    'Text Field8': 6.2,
    'Text Field9': 9.0,
    'Text Field10': 9.0,
    'Text Field11': 9.0,
    'Text Field12': 9.0,
}

for w in p1.widgets():
    if w.field_name in p1_vals:
        w.field_value = p1_vals[w.field_name]
        if w.field_name in font_sizes_p1:
            w.text_fontsize = font_sizes_p1[w.field_name]
        w.update()

# Insert Impact Justification text cleanly into the Section 3 left column space on Page 1 (x: 46.2 to 195.0, avoiding middle legend and matrix)
rect_just = fitz.Rect(46.2, 528.0, 195.0, 626.0)
p1.insert_textbox(rect_just, just_text, fontsize=6.2, fontname='helv', color=(0, 0, 0))

# Page 2 Field Values & Font Sizes
p2_vals = {
    'Text Field34': '26449',
    'Text Field13': 'Microbiology',
    'Text Field14': 'Robin Seymour',
    'Text Field15': 'Microbiology Supervisor (Sterile Lab)',
    'Date Field1': date_requested,
    'Text Field16': 'Microbiology',
    'Text Field17': 'Kathan Parikh',
    'Text Field18': 'Associate Director of Microbiology',
    'Date Field2': date_requested,
    'Check Box18': 'Yes', # N/A for training (handled in NCR N26414)
    'Text Field33': (
        f"\nResult correction in EagleTrax coupled with closed {oos_id} and closed NCR {ncr_id}. "
        "Re-training on MICRO-SOP-12 and CORP-SOP-15 completed under NCR N26414 in ZenQMS."
    ),
}

font_sizes_p2 = {
    'Text Field34': 9.0,
    'Text Field13': 8.0,
    'Text Field14': 8.0,
    'Text Field15': 7.5,
    'Date Field1': 8.0,
    'Text Field16': 8.0,
    'Text Field17': 8.0,
    'Text Field18': 7.5,
    'Date Field2': 8.0,
    'Text Field33': 7.5,
}

for w in p2.widgets():
    if w.field_name in p2_vals:
        w.field_value = p2_vals[w.field_name]
        if w.field_name in font_sizes_p2:
            w.text_fontsize = font_sizes_p2[w.field_name]
        w.update()

out_pdf_name = f"CR-26449 - Result Correction Pass to Fail - {sample_id} ({oos_id}).pdf"
out_pdf_scratch = os.path.join(OUTPUT_DIR, out_pdf_name)
doc.save(out_pdf_scratch)
doc.close()
print(f"Generated CR PDF: {out_pdf_scratch}")

# Render PNGs for visual inspection
doc_check = fitz.open(out_pdf_scratch)
pix1 = doc_check[0].get_pixmap(dpi=150)
pix1.save(os.path.join(OUTPUT_DIR, "cr_page1.png"))
pix2 = doc_check[1].get_pixmap(dpi=150)
pix2.save(os.path.join(OUTPUT_DIR, "cr_page2.png"))
doc_check.close()
print("Rendered cr_page1.png and cr_page2.png for visual verification.")

# 2. Generate Professional Word (.docx) Companion Document
doc_word = docx.Document()

# Page setup: Margins 0.7 in
for section in doc_word.sections:
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

def add_heading(text, level=1):
    h = doc_word.add_paragraph()
    h.paragraph_format.space_before = Pt(10)
    h.paragraph_format.space_after = Pt(3)
    run = h.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(11.5 if level == 1 else 10)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    return h

def add_para(text, bold_prefix=None, space_after=Pt(3)):
    p = doc_word.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = space_after
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Arial'
        r_pre.font.size = Pt(9.5)
        r_pre.font.bold = True
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(9.5)
    return p

# Title Block
title_p = doc_word.add_paragraph()
title_p.paragraph_format.space_before = Pt(0)
title_p.paragraph_format.space_after = Pt(2)
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_t1 = title_p.add_run("EAGLE ANALYTICAL SERVICES\n")
r_t1.font.name = 'Arial'
r_t1.font.size = Pt(13)
r_t1.font.bold = True
r_t1.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

r_t2 = title_p.add_run("CHANGE CONTROL REQUEST FORM (CORP-FORM-4 / 3.100.004.F01 Rev 10)\n")
r_t2.font.name = 'Arial'
r_t2.font.size = Pt(11)
r_t2.font.bold = True

r_t3 = title_p.add_run(f"CR Number: {cr_id} | Coupled with: {oos_id} & NCR {ncr_id}")
r_t3.font.name = 'Arial'
r_t3.font.size = Pt(10)
r_t3.font.italic = True

p_div = doc_word.add_paragraph()
p_div.paragraph_format.space_before = Pt(2)
p_div.paragraph_format.space_after = Pt(6)
r_div = p_div.add_run("―" * 65)
r_div.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

# Section 1
add_heading("1. Request Originator & Reason for Change")
table1 = doc_word.add_table(rows=4, cols=2)
table1.alignment = WD_TABLE_ALIGNMENT.CENTER
t1_data = [
    ("Name:", "Qiyue Chen", "Department:", "Microbiology"),
    ("Date Requested:", date_requested, "Change Control #:", cr_id),
    ("Type of Change:", "Other: EagleTrax Test Result Correction (Pass to Fail)", "Affected System:", "EagleTrax LIMS"),
    ("Reason for Change:", f"OOS #{oos_id} & NCR #{ncr_id} (Sample result transposition correction)", "Coupled With:", f"{oos_id} (Phase I OOS Report)"),
]
for row_idx, row in enumerate(table1.rows):
    d = t1_data[row_idx]
    c0 = row.cells[0]
    c1 = row.cells[1]
    c0.text = f"{d[0]} {d[1]}"
    c1.text = f"{d[2]} {d[3]}"
    for c in [c0, c1]:
        c.paragraphs[0].runs[0].font.name = 'Arial'
        c.paragraphs[0].runs[0].font.size = Pt(9)
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

# Section 2
add_heading("2. Change Description")
table2 = doc_word.add_table(rows=1, cols=2)
table2.alignment = WD_TABLE_ALIGNMENT.CENTER
c_ref = table2.rows[0].cells[0]
c_desc = table2.rows[0].cells[1]
c_ref.width = Inches(2.2)
c_desc.width = Inches(4.8)

c_ref.text = ref_text
c_desc.text = desc_text
for c in [c_ref, c_desc]:
    for p in c.paragraphs:
        for r in p.runs:
            r.font.name = 'Arial'
            r.font.size = Pt(8.5)

# Section 3
add_heading("3. Impact of Change")
table3 = doc_word.add_table(rows=2, cols=4)
table3.alignment = WD_TABLE_ALIGNMENT.CENTER
t3_headers = ["Frequency", "Severity", "Magnitude", "Impact Grading Score & Level"]
for col_i, h in enumerate(t3_headers):
    c = table3.rows[0].cells[col_i]
    c.text = h
    c.paragraphs[0].runs[0].font.name = 'Arial'
    c.paragraphs[0].runs[0].font.size = Pt(9)
    c.paragraphs[0].runs[0].font.bold = True
    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

t3_vals = [
    "<2 instances = 1",
    "Changes with retroactive implication = 3",
    "<2 Documents/Processes = 1",
    "Score: 1 × 3 × 1 = 3\nLevel: Major (Score 3 to 17)"
]
for col_i, v in enumerate(t3_vals):
    c = table3.rows[1].cells[col_i]
    c.text = v
    c.paragraphs[0].runs[0].font.name = 'Arial'
    c.paragraphs[0].runs[0].font.size = Pt(8.5)
    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

p_just = doc_word.add_paragraph()
p_just.paragraph_format.space_before = Pt(5)
p_just.paragraph_format.space_after = Pt(6)
r_j = p_just.add_run(just_text)
r_j.font.name = 'Arial'
r_j.font.size = Pt(9)
r_j.font.italic = True

# Section 4
add_heading("4. Reviewed & Approved")
table4 = doc_word.add_table(rows=3, cols=4)
table4.alignment = WD_TABLE_ALIGNMENT.CENTER
t4_headers = ["Department", "Name", "Title", "Date"]
for col_i, h in enumerate(t4_headers):
    c = table4.rows[0].cells[col_i]
    c.text = h
    c.paragraphs[0].runs[0].font.name = 'Arial'
    c.paragraphs[0].runs[0].font.size = Pt(9)
    c.paragraphs[0].runs[0].font.bold = True

t4_rows = [
    ("Microbiology", "Robin Seymour", "Microbiology Supervisor (Sterile Lab)", date_requested),
    ("Microbiology", "Kathan Parikh", "Associate Director of Microbiology", date_requested)
]
for r_i, r_data in enumerate(t4_rows):
    for c_i, val in enumerate(r_data):
        c = table4.rows[r_i+1].cells[c_i]
        c.text = val
        c.paragraphs[0].runs[0].font.name = 'Arial'
        c.paragraphs[0].runs[0].font.size = Pt(8.5)

# Quality Closure
add_heading("Quality Closure (Completed by Quality)")
add_para("Training Required: N/A (Documented Read & Understood retraining on MICRO-SOP-12 and CORP-SOP-15 was previously assigned and completed under Nonconformance Report NCR N26414 in ZenQMS).")
add_para("Comments: Result correction in EagleTrax coupled with closed OOS-261814 and closed NCR N26414. Ready for final Quality closure in accordance with CORP-SOP-4.", bold_prefix="Closure Comments: ")

out_docx_name = f"CR-26449 - Result Correction Pass to Fail - {sample_id} ({oos_id}).docx"
out_docx_scratch = os.path.join(OUTPUT_DIR, out_docx_name)
doc_word.save(out_docx_scratch)
print(f"Generated CR Word Document: {out_docx_scratch}")

# 3. Deploy / Sync files to Documents directory
files_to_sync = [
    out_pdf_scratch,
    out_docx_scratch
]
for fpath in files_to_sync:
    if fpath and os.path.exists(fpath):
        dest = os.path.join(DOCUMENTS_DIR, os.path.basename(fpath))
        try:
            shutil.copy2(fpath, dest)
            print(f"Synced to Documents: {dest}")
        except Exception as e:
            print(f"Sync error for {fpath}: {e}")

print("\n--- CR-26449 GENERATION AND DEPLOYMENT COMPLETED SUCCESSFULLY! ---")
