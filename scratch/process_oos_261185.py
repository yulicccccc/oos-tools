import os, sys, shutil
from datetime import datetime
import fitz
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client as win32

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docs_dir = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
history_dir = os.path.join(docs_dir, ".history")
scratch_dir = os.path.join(docs_dir, "scratch")
os.makedirs(history_dir, exist_ok=True)
os.makedirs(scratch_dir, exist_ok=True)

src_pdf_name = "OOS-261185.pdf"
src_docx_name = "EM table OOS-261185 08MAY2026.docx"

src_pdf_path = os.path.join(desktop_dir, src_pdf_name)
src_docx_path = os.path.join(desktop_dir, src_docx_name)

print("=== STEP 1: BACKUPS ===")
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
if os.path.exists(src_pdf_path):
    shutil.copy2(src_pdf_path, os.path.join(history_dir, f"OOS_261185_backup_{timestamp}.pdf"))
if os.path.exists(src_docx_path):
    shutil.copy2(src_docx_path, os.path.join(history_dir, f"EM_table_261185_backup_{timestamp}.docx"))
print("Backed up existing PDF and DOCX to .history")

# ==========================================
# STEP 2: AMEND WORD TABLE (STRICTLY 1 PAGE)
# ==========================================
print("\n=== STEP 2: AMEND WORD TABLE ===")
backup_table = os.path.join(history_dir, "EM_table_261185_backup.docx")
if os.path.exists(backup_table):
    doc_tbl = Document(backup_table)
elif os.path.exists(src_docx_path):
    doc_tbl = Document(src_docx_path)
else:
    raise FileNotFoundError("Could not find table DOCX source!")

# Update paragraph headers (do NOT modify section margins - keep original 1.0 in margins)
for p in doc_tbl.paragraphs:
    if "Table 1:" in p.text:
        p.text = ""
        r = p.add_run("Table 1: Read Dates & Incubation Observation")
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.bold = True
    elif "Table 2:" in p.text:
        p.text = ""
        r = p.add_run("Table 2: Environmental Monitoring Plates for Analyst and Cleanroom Bracketing")
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.bold = True

def update_cell_text(cell, text, font_size=7.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    lines = text.split('\n')
    for l_idx, line in enumerate(lines):
        p = cell.paragraphs[0] if l_idx == 0 else cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(line)
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size)
        r.bold = bold

# Table 1: Surgical edits
t1 = doc_tbl.tables[0]
update_cell_text(t1.rows[1].cells[0], "ETX-260518-0250", font_size=7.0)
update_cell_text(t1.rows[1].cells[1], "SMO\n08MAY 2026", font_size=7.0)
update_cell_text(t1.rows[1].cells[2], "Active air\nsampling plate\n(Suite 115B ISO 7)", font_size=7.0)
update_cell_text(t1.rows[1].cells[6], "11 CFUs on active\nair sample plate", font_size=7.0)
update_cell_text(t1.rows[1].cells[7], "Penicillium decumbens,\nCladosporium tenuissimum,\nCladosporium langeronii,\nCladosporium halotolerans", font_size=6.5)

# Table 2: Surgical edits (keeping native widths, borders, and column layouts)
t2 = doc_tbl.tables[1]
# Row 2 (29APR 2026 Active Air)
update_cell_text(t2.rows[2].cells[4], "Week Before\nTesting Date", font_size=6.5)

# Row 3 (08MAY 2026 Active Air)
update_cell_text(t2.rows[3].cells[4], "Week of\nTesting Date", font_size=6.5)
update_cell_text(t2.rows[3].cells[6], "ETX-260518-0249\nETX-260518-0250", font_size=6.5)
update_cell_text(t2.rows[3].cells[7], "Staphylococcus aureus\nPenicillium decumbens\nCladosporium tenuissimum\nCladosporium langeronii\nCladosporium halotolerans", font_size=6.0)

# Row 4 (14MAY 2026 Active Air)
update_cell_text(t2.rows[4].cells[4], "Week After\nTesting Date", font_size=6.5)
update_cell_text(t2.rows[4].cells[6], "ETX-260526-0391\nETX-260526-0395", font_size=6.5)
update_cell_text(t2.rows[4].cells[7], "Gram (+) cocci\nCandida orthopsilosis\nBudding yeast", font_size=6.5)

# Row 6 (29APR 2026 Surface)
update_cell_text(t2.rows[6].cells[0], "Surface Sampling of Cleanrooms", font_size=6.5)
update_cell_text(t2.rows[6].cells[4], "Week Before\nTesting Date", font_size=6.5)

# Row 7 (08MAY 2026 Surface)
update_cell_text(t2.rows[7].cells[4], "Week of\nTesting Date", font_size=6.5)

# Row 8 (14MAY 2026 Surface)
update_cell_text(t2.rows[8].cells[4], "Week After\nTesting Date", font_size=6.5)
update_cell_text(t2.rows[8].cells[5], "1 CFU on table in ISO 8 115 and\n1 CFU on the cart in 115 ISO 8", font_size=6.5)
update_cell_text(t2.rows[8].cells[6], "ETX-260526-0424\nETX-260526-0425", font_size=6.5)
update_cell_text(t2.rows[8].cells[7], "Gram (+) rods\nStaphylococcus capitis", font_size=6.5)

# Save Word Table
out_table_docx_1 = os.path.join(desktop_dir, "EM table OOS-261185 08MAY2026.docx")
out_table_docx_2 = os.path.join(desktop_dir, "Tables OOS-261185 EM SMO 115B Air 08MAY2026 - EM.docx")
doc_tbl.save(out_table_docx_1)
doc_tbl.save(out_table_docx_2)
print("Saved amended table DOCX to Desktop.")

# Convert to PDF via Word COM
out_table_pdf = os.path.join(desktop_dir, "EM table OOS-261185 08MAY2026.pdf")
out_table_pdf_2 = os.path.join(desktop_dir, "Tables OOS-261185 EM SMO 115B Air 08MAY2026 - EM.pdf")

word_app = win32.Dispatch("Word.Application")
word_app.Visible = False
try:
    doc_com = word_app.Documents.Open(os.path.abspath(out_table_docx_1))
    doc_com.SaveAs(os.path.abspath(out_table_pdf), FileFormat=17)
    doc_com.Close(False)
    shutil.copy2(out_table_pdf, out_table_pdf_2)
    print("Converted amended table to 1-page PDF via Word COM.")
finally:
    word_app.Quit()

# Check page count of table PDF
doc_check = fitz.open(out_table_pdf)
print(f"Table PDF page count: {len(doc_check)} (Must be 1)")
doc_check.close()

# ==========================================
# STEP 3: UPDATE PDF FORM FIELDS
# ==========================================
print("\n=== STEP 3: UPDATE PDF FORM FIELDS ===")
pristine_pdf = os.path.join(history_dir, "OOS_261185_backup_before_field_update.pdf")

doc_pdf = fitz.open(pristine_pdf)

field_updates = {
    # Page 1
    "Text Field0": "Simin Mohammad",
    "Date Field0": "08MAY2026",
    "Date Field1": "18MAY2026",
    "Date Field2": "18MAY2026",
    "Text Field1": "Environmental Monitoring",
    "Text Field2": "ETX-260518-0250",
    "Text Field3": (
        "Simin Mohammad (Weekly Active Air Sampling Plate Setup)\r"
        "Maraya Chukwumerije (Weekly Active Air Sampling Plate Reader)\r"
        "Sophia Santamaria (Weekly Active Air Sampling Plate Reader)"
    ),
    "Text Field4": "EM SMO 115B Air 08MAY2026",
    "Text Field5": "Plate",
    "Text Field6": "EM SMO 115B Air 08MAY2026",
    "Text Field7": "The CFU count for the environmental monitoring plate exceeded the action level.",
    "Text Field8": "MICRO-SOP-2",
    "Text Field9": "23-Jul-2026",
    "Text Field10": "16",
    "Text Field11": "Action level: >= 10 CFU/Plate",
    "Text Field12": "Kathan Parikh",
    "Date Field3": "18MAY2026",
    "Text Field13": "Yes, analysts Simin Mohammad, Maraya Chukwumerije, and Sophia Santamaria were comprehensively interviewed.",
    "Text Field14": "N/A",
    "Text Field15": "Yes, as per MICRO-SOP-2",
    "Text Field16": "Yes, as per MICRO-SOP-2",
    "Text Field17": "Yes, Information is available in EagleTrax under ETX-260518-0250",
    "Text Field18": "Yes, Information is available in EagleTrax under ETX-260323-0434",
    "Text Field19": "N/A",
    "Text Field20": "N/A",
    "Text Field21": "Yes, as per MICRO-SOP-2",
    
    # Page 2
    "Text Field32": "CR115 (Sensor E001737)",
    "Text Field43": "Incubator E001034 (Sensor E001501)\rIncubator E001031 (Sensor E001505)",
    "Text Field44": "Aug 2027 / Feb 2027\rAug 2027 / Feb 2027",
    
    # Page 3
    "Text Field49": (
        "The analyst involved in the Active Air Sampling plate setup, Simin Mohammad, and the analysts involved in reading the plate, "
        "Maraya Chukwumerije and Sophia Santamaria, were interviewed comprehensively. Their answers are recorded throughout this document.\n\n"
        "The EM plates were stored in compliance with the supplier's recommendations, and their integrity was visually inspected prior to use. "
        "Furthermore, the plates were confirmed to be within their valid expiration dates. All the supplies were thoroughly disinfected according "
        "to MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology). The functionality of both incubators was verified through a review "
        "of data obtained from our comprehensive in-house continuous monitoring system.\n\n"
        "Active Air Sampling was performed by analyst Simin Mohammad during weekly environmental monitoring processing in Suite 115 (ISO 8), "
        "Suite 115A (ISO 7), and Suite 115B (ISO 7) on 08 May 2026 as per MICRO-SOP-2 (Environmental Monitoring of the Cleanroom Facility). "
        "The plates were initially incubated at a temperature of 30–35°C in incubator E001031 for a minimum duration of 48 hours, commencing on "
        "08 May 2026. Following completion of a minimum of 48 hours of incubation on 11 May 2026, the plates were further incubated for a minimum of "
        "5 days at 20–25°C in incubator E001034, with the incubation concluding on 18 May 2026. Please see Table 1 for detailed information on the "
        "observations during respective incubations.\n\n"
        "Based on the observations in Table 1, since the CFU count exceeded the action level for one of the Active Air Sampling plates (115B), "
        "the plate was submitted for Microbial Identification under ETX-260518-0250. The colonies were identified as Penicillium decumbens, "
        "Cladosporium fasting / tenuissimum, Cladosporium langeronii, and Cladosporium halotolerans. To observe if the organisms identified were transient "
        "in nature or recurring, weekly environmental monitoring plates for the clean room were bracketed to include the week before testing, "
        "the week of testing, and the week after testing as detailed in Table 2 (please see attached)."
    ).replace("Cladosporium fasting / tenuissimum", "Cladosporium tenuissimum"),
    
    # Page 4
    "Text Field50": (
        "Environmental Monitoring Summary: Weekly environmental sampling for the previous week showed 1 CFU for 115 ISO 8 identified as Micrococcus luteus. "
        "Week of testing showed 1 CFU in 115A ISO 7 and 11 CFU in 115B ISO 7 identified as Staphylococcus aureus, Penicillium decumbens, "
        "Cladosporium tenuissimum, Cladosporium langeronii, and Cladosporium halotolerans. Lastly, the following week of testing showed 1 CFU in 115 ISO 8, "
        "11 CFU in 115B ISO 7, 1 CFU on table in ISO 8 115, and 1 CFU on the cart in 115 ISO 8 identified as Gram (+) cocci, Candida orthopsilosis, "
        "Budding yeast, Gram (+) rods, and Staphylococcus capitis.\n\n"
        "During the interview with the analyst, they indicated that no obvious abnormalities or deviations in the testing procedure were observed. "
        "All the samples were thoroughly disinfected prior to testing. Moreover, the CR115 was thoroughly cleaned and prepared before initiating "
        "the testing as per MICRO-SOP-2 (Environmental Monitoring of the Cleanroom Facility) and MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology).\n\n"
        "Prior to the event, monthly cleaning and disinfection of the outermost ISO 8 Anteroom (Suite 115), the middle ISO 7 Buffer room (Suite 115A), "
        "the innermost ISO 7 clean room (Suite 115B), and its containing ISO 5 Biosafety Cabinets were performed on 26-Apr-2026 by analysts Rey Estrada "
        "and Tamiru Kotisso as per MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology). It was documented that all H2O2 indicators passed, "
        "confirming the efficient pre-event cleaning and established state of control of Suite 115. Furthermore, subsequent monthly cleaning was "
        "performed on 31-May-2026 by analysts Tamiru Kotisso and Cuong Du with all H2O2 indicators passing, confirming ongoing facility control. "
        "Additionally, cleaning and disinfecting was performed both prior to and after the testing process as per MICRO-SOP-9."
    ),
    
    # Page 5
    "Text Field51": (
        "It is also important to note that no samples processed in Suite 115 for the week of testing "
        "(03-May-2026 to 09-May-2026) failed testing.\r\n\r\n"
        "Based on the available investigation findings, no specific analyst-related, procedural, "
        "equipment-related, or other laboratory-related cause was identified. The excursion appears "
        "to have been transient and non-recurring, and no assignable root cause was established. "
        "It is to be noted that the growth observed on the weekly active air and weekly surface sampling "
        "plates for the day of testing did not follow a trend, indicating that the contamination was "
        "transient in nature and that routine daily disinfection procedures were effective in eliminating "
        "the contamination. Furthermore, no trend was observed in the clean room's previous weekly EM "
        "data, therefore, no preventive and corrective actions are deemed necessary at this time."
    ),
    
    # OOS Number across all pages
    "Text Field57": "261185"
}

for page in doc_pdf:
    for w in page.widgets():
        # Update named fields
        if w.field_name in field_updates:
            new_val = field_updates[w.field_name]
            clean_val = str(new_val).replace("₂", "2")
            w.field_value = clean_val
            if w.field_name == "Text Field11":
                w.text_fontsize = 6.5
            w.update()
        # Update OOS Number widget if unnamed
        elif w.field_name == "" and abs(w.rect.y0 - 122) < 15:
            w.field_value = "261185"
            w.update()

# ==========================================
# STEP 4: ASSEMBLE 8-PAGE FULL PDF PACKAGE
# ==========================================
print("\n=== STEP 4: ASSEMBLE 8-PAGE FULL PDF PACKAGE ===")
# Page 7: Standard Version History page (take from 261186 or 261242)
vh_src = os.path.join(desktop_dir, "OOS-261186.pdf")
doc_vh = fitz.open(vh_src)
# Page index 6 is Page 7 (Version History)
doc_pdf.insert_pdf(doc_vh, from_page=6, to_page=6)
doc_vh.close()
print("Inserted Page 7: Standard Version History.")

# Page 8: Amended 1-page EM Table
doc_table = fitz.open(out_table_pdf)
doc_pdf.insert_pdf(doc_table)
doc_table.close()
print("Inserted Page 8: Amended 1-Page EM Table.")

final_full_pdf = os.path.join(desktop_dir, "OOS-261185 EM SMO 115B Air 08MAY2026 - EM.pdf")
doc_pdf.save(final_full_pdf)
doc_pdf.close()

# Synchronize original Desktop PDFs
shutil.copy2(final_full_pdf, src_pdf_path)
print(f"Assembled complete 8-page package to:\n  {final_full_pdf}\n  {src_pdf_path}")

# ==========================================
# STEP 5: GENERATE FULL WORD REPORT
# ==========================================
print("\n=== STEP 5: GENERATE FULL WORD REPORT ===")
report_docx_path = os.path.join(desktop_dir, "OOS-261185 EM SMO 115B Air 08MAY2026 - EM.docx")

doc_rep = Document()
for s in doc_rep.sections:
    s.top_margin = Inches(0.8)
    s.bottom_margin = Inches(0.8)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)

p_title = doc_rep.add_paragraph()
r_t = p_title.add_run("LABORATORY OUT-OF-SPECIFICATION (OOS) INVESTIGATION REPORT")
r_t.bold = True
r_t.font.name = "Calibri"
r_t.font.size = Pt(14)
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

p_sub = doc_rep.add_paragraph()
r_sub = p_sub.add_run("OOS-261185 | EM SMO 115B Air 08MAY2026 | Active Air Sampling")
r_sub.font.name = "Calibri"
r_sub.font.size = Pt(11)
r_sub.italic = True
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Summary Info Table
def format_report_cell(cell, text, font_size=9.5, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, bg_hex=None):
    if bg_hex:
        tcPr = cell._element.get_or_add_tcPr()
        for child in list(tcPr):
            if child.tag.endswith('shd'):
                tcPr.remove(child)
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_hex}"/>'))
    cell.text = ""
    lines = text.split('\n')
    for l_idx, line in enumerate(lines):
        p = cell.paragraphs[0] if l_idx == 0 else cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(line)
        r.font.name = "Calibri"
        r.font.size = Pt(font_size)
        r.bold = bold

info_table = doc_rep.add_table(rows=6, cols=2)
info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
info_data = [
    ("OOS Investigation Number:", "OOS-261185"),
    ("Test & Sample Description:", "Environmental Monitoring - EM SMO 115B Air 08MAY2026 (Active Air Plate)"),
    ("Testing Site & Location:", "Suite 115 (ISO 8 Anteroom), Suite 115A (ISO 7 Buffer), Suite 115B (ISO 7 Cleanroom)"),
    ("Initiator / Setup Analyst:", "Simin Mohammad (Weekly Active Air Sampling Plate Setup)"),
    ("Reading Analysts:", "Maraya Chukwumerije (≥ 48H Reader) & Sophia Santamaria (5-Day Reader)"),
    ("Action Level & Result:", "Action level: >= 10 CFU/Plate | Result: 11 CFUs (Penicillium decumbens & Cladosporium spp.)")
]
for idx, (label, val) in enumerate(info_data):
    r = info_table.rows[idx]
    r.cells[0].width = Inches(2.5)
    r.cells[1].width = Inches(4.5)
    format_report_cell(r.cells[0], label, font_size=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, bg_hex="F2F2F2")
    format_report_cell(r.cells[1], val, font_size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT)

doc_rep.add_paragraph().paragraph_format.space_after = Pt(8)

# Add Narrative Sections
sections_text = [
    ("1. Phase I Summary – Setup, Incubation & Microbial Identification", field_updates["Text Field49"]),
    ("2. Phase I Summary – Cleanroom Bracketing, Analyst Interview & Monthly Cleaning", field_updates["Text Field50"]),
    ("3. Phase I Summary – Cleanroom Control Assessment & Root Cause Statement", field_updates["Text Field51"])
]

for title, body in sections_text:
    p_sec = doc_rep.add_paragraph()
    r_s = p_sec.add_run(title)
    r_s.bold = True
    r_s.font.name = "Calibri"
    r_s.font.size = Pt(11.5)
    p_sec.paragraph_format.space_before = Pt(8)
    p_sec.paragraph_format.space_after = Pt(4)
    
    for para in body.split('\n\n'):
        p_b = doc_rep.add_paragraph()
        r_b = p_b.add_run(para.replace('\n', ' '))
        r_b.font.name = "Calibri"
        r_b.font.size = Pt(10)
        p_b.paragraph_format.space_after = Pt(6)
        p_b.paragraph_format.line_spacing = 1.15

doc_rep.save(report_docx_path)
print(f"Saved full Word report to: {report_docx_path}")

# ==========================================
# STEP 6: GENERATE ROBIN REVIEW EMAIL & CLIPBOARD
# ==========================================
print("\n=== STEP 6: GENERATE ROBIN REVIEW EMAIL & CLIPBOARD ===")
email_docx_out = os.path.join(desktop_dir, "OOS-261185 Review Email.docx")
email_html_out = os.path.join(desktop_dir, "OOS-261185_Review_Email_Robin_Format.html")

doc_email = Document()
for s in doc_email.sections:
    s.top_margin = Inches(0.8)
    s.bottom_margin = Inches(0.8)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)

p_g = doc_email.add_paragraph()
r_g = p_g.add_run("Good morning @Simin Mohammad,")
r_g.font.name = "Calibri"
r_g.font.size = Pt(11)
p_g.paragraph_format.space_after = Pt(6)

p_i = doc_email.add_paragraph()
r_i = p_i.add_run(
    "I have reviewed OOS-261185 and made my edits to this and summary of the major ones as below. "
    "Please also note that the highlighted sections in the table have been amended:"
)
r_i.font.name = "Calibri"
r_i.font.size = Pt(11)
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

email_table_data = [
    ["Section", "Unreviewed Version", "Reviewed Version"],
    [
        "Root Cause Conclusion",
        "The conclusion stated that the excursion \"may be attributed to a potential analyst error.\"",
        "Revised the root cause conclusion to adhere strictly to cGMP standards, removing speculative analyst error statements and confirming that based on comprehensive investigation findings, no specific analyst-related, procedural, equipment-related, or other laboratory-related cause was identified. The excursion was transient and non-recurring, with no assignable root cause established and effective routine disinfection demonstrated."
    ],
    [
        "Monthly Cleaning Documentation",
        "Only referenced single monthly cleaning date without clarifying pre-event status or subsequent ongoing facility control.",
        "Documented both the pre-event monthly cleaning performed on 26-Apr-2026 by analysts Rey Estrada and Tamiru Kotisso (confirming established state of control prior to testing with passing H2O2 indicators) and the subsequent monthly cleaning performed on 31-May-2026 by analysts Tamiru Kotisso and Cuong Du (confirming ongoing facility control)."
    ],
    [
        "Equipment Calibration Dates",
        "Listed incubator calibration due dates running together as \"Aug 2026 / Feb 2027\".",
        "Updated Incubator E001034 (Sensor E001501) and Incubator E001031 (Sensor E001505) calibration due dates to Aug 2027 / Feb 2027 per current calibration status."
    ],
    [
        "SOP Reference & Form Field Updates",
        "Referenced legacy revision of MICRO-SOP-2 Rev 15 (05-AUG-2025). Initiator was listed as \"Simin Mohammad (written by Simin Mohammad)\". Reading analysts ran together with setup analyst on a single line.",
        "Updated procedure to current MICRO-SOP-2 Rev 16 (Effective Date: 23-Jul-2026). Corrected initiator to \"Simin Mohammad\". Delineated setup analyst Simin Mohammad from plate readers Maraya Chukwumerije (≥ 48H Reader) and Sophia Santamaria (5-Day Reader) with dedicated roles across form fields."
    ],
    [
        "EM Table Formatting & Verification",
        "The draft Word table listed incorrect submission ID ETX-260526-0461 and unreviewed organism Talaromyces purpurogenus in Table 1, and active air bracketing in Table 2 was mislabeled as surface sampling.",
        "Amended Table 1 to reflect ETX-260518-0250 and accurate fungal identification (Penicillium decumbens, Cladosporium tenuissimum, Cladosporium langeronii, and Cladosporium halotolerans). Standardized Table 2 bracketing headers, corrected timing descriptions, and attached the amended 1-page table as Page 8 to form the complete 8-page package."
    ]
]

col_widths = [1.8, 2.6, 2.8]
e_table = doc_email.add_table(rows=len(email_table_data), cols=3)
e_table.alignment = WD_TABLE_ALIGNMENT.LEFT
e_table.autofit = False

tblPr = e_table._element.xpath('w:tblPr')
if tblPr:
    for jc in tblPr[0].xpath('w:jc'):
        tblPr[0].remove(jc)
    tblPr[0].append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
    
    for ind in tblPr[0].xpath('w:tblInd'):
        tblPr[0].remove(ind)
    tblPr[0].append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="0" w:type="dxa"/>'))

def add_formatted_runs(paragraph, text, align=WD_ALIGN_PARAGRAPH.LEFT):
    paragraph.alignment = align
    clean_text = text.replace('**', '')
    r = paragraph.add_run(clean_text)
    r.font.name = "Calibri"
    r.font.size = Pt(10)
    r.bold = False

for r_idx, row_content in enumerate(email_table_data):
    row = e_table.rows[r_idx]
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

# Closing
p_c = doc_email.add_paragraph()
p_c.paragraph_format.space_before = Pt(14)
p_c.paragraph_format.space_after = Pt(0)
r_c = p_c.add_run("please review them, sign and please route for signatures.")
r_c.font.name = "Calibri"
r_c.font.size = Pt(11)

doc_email.save(email_docx_out)
print(f"Saved Word review email to: {email_docx_out}")

# Clean HTML (NO BOLD inside cells)
def html_clean(text):
    lines = text.split('\n')
    clean_lines = [line.replace('**', '') for line in lines]
    return '<br>'.join(clean_lines)

html_rows = ""
for r_idx, row in enumerate(email_table_data[1:]):
    sec, unrev, rev = row
    sec_html = html_clean(sec)
    unrev_html = html_clean(unrev)
    rev_html = html_clean(rev)
    
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
<p style="margin: 0 0 12px 0; font-family: Calibri, sans-serif; font-size: 11pt;">I have reviewed OOS-261185 and made my edits to this and summary of the major ones as below. Please also note that the highlighted sections in the table have been amended:</p>
<table border="1" bordercolor="#000000" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Calibri, sans-serif; font-size: 10pt; margin: 12px 0 16px 0;">
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
<p style="margin: 14px 0 0 0; font-family: Calibri, sans-serif; font-size: 11pt;">please review them, sign and please route for signatures.</p>
<!--EndFragment-->
</body>
</html>
"""

with open(email_html_out, "w", encoding="utf-8") as f:
    f.write(full_html)
print(f"Saved clean HTML to: {email_html_out}")

# COPY DIRECTLY TO WINDOWS CLIPBOARD VIA WORD COM
try:
    word_copy = win32.Dispatch("Word.Application")
    word_copy.Visible = False
    doc_to_copy = word_copy.Documents.Open(os.path.abspath(email_docx_out))
    doc_to_copy.Content.Copy()
    doc_to_copy.Close(False)
    word_copy.Quit()
    print("SUCCESS: Copied native Word formatting directly to Windows Clipboard via Word COM!")
except Exception as e:
    print(f"Word COM copy failed: {e}")

# Render preview of assembled 8-page PDF
doc_v = fitz.open(final_full_pdf)
print(f"Final assembled PDF page count: {len(doc_v)}")
for i, page in enumerate(doc_v):
    pix = page.get_pixmap(dpi=150)
    pix.save(os.path.join(scratch_dir, f"assembled_261185_p{i+1}.png"))
print("Saved all 8 page preview images to scratch/")
doc_v.close()
print("=== COMPLETE: OOS-261185 REVISION PIPELINE FINISHED SUCCESSFULLY ===")
