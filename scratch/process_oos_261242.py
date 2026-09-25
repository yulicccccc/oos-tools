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
from pypdf import PdfReader, PdfWriter
import win32com.client as win32
import win32clipboard

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docs_dir = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
history_dir = os.path.join(docs_dir, ".history")
scratch_dir = os.path.join(docs_dir, "scratch")
os.makedirs(history_dir, exist_ok=True)
os.makedirs(scratch_dir, exist_ok=True)

src_pdf_name = "OOS-261242 (1).pdf"
src_docx_name = "EM table OOS-261242 14MAY2026 (3) (2) (1).docx"

src_pdf_path = os.path.join(desktop_dir, src_pdf_name)
src_docx_path = os.path.join(desktop_dir, src_docx_name)

print("=== STEP 1: BACKUPS ===")
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
shutil.copy2(src_pdf_path, os.path.join(history_dir, f"OOS_261242_backup_{timestamp}.pdf"))
shutil.copy2(src_docx_path, os.path.join(history_dir, f"EM_table_261242_backup_{timestamp}.docx"))
print("Backed up original PDF and DOCX to .history")

# ==========================================
# STEP 2: AMEND WORD TABLE (STRICTLY 1 PAGE)
# ==========================================
print("\n=== STEP 2: AMEND WORD TABLE ===")
doc_tbl = Document(src_docx_path)

# Set margins to 0.4 in to guarantee 1 page
for section in doc_tbl.sections:
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)

# Update paragraph headers
for p in doc_tbl.paragraphs:
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
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

def format_cell(cell, text, font_size=7.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, bg_hex=None):
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

# Table 1: Amend
t1 = doc_tbl.tables[0]
t1.autofit = False
# Row 0: Headers
for c in t1.rows[0].cells:
    txt = c.text.strip().replace(' // ', '\n')
    format_cell(c, txt, font_size=7.5, bold=True, bg_hex="F2F2F2")

# Row 1: Data
t1_r1 = t1.rows[1]
format_cell(t1_r1.cells[0], "ETX-260526-0395", font_size=7.0)
format_cell(t1_r1.cells[1], "SMO\n14MAY 2026", font_size=7.0)
format_cell(t1_r1.cells[2], "Active air sampling plate\n(Suite 115B ISO 7)", font_size=7.0)
format_cell(t1_r1.cells[3], "SAS", font_size=7.0)
format_cell(t1_r1.cells[4], "11 CFUs on active air sample plate", font_size=7.0)
format_cell(t1_r1.cells[5], "SAS", font_size=7.0)
format_cell(t1_r1.cells[6], "11 CFUs on active air sample plate", font_size=7.0)
format_cell(t1_r1.cells[7], "Candida orthopsilosis\n(Budding yeast)", font_size=7.0, bold=True)

# Table 2: Amend
t2 = doc_tbl.tables[1]
t2.autofit = False

# Row 0: Headers
for c in t2.rows[0].cells:
    txt = c.text.strip().replace(' // ', '\n')
    format_cell(c, txt, font_size=7.0, bold=True, bg_hex="F2F2F2")

# Section header 1: Row 1
for c in t2.rows[1].cells:
    format_cell(c, "Weekly Active Air Sampling Bracketing 115", font_size=7.0, bold=True, bg_hex="E8EEF5")

# Row 2: 08MAY 2026
t2_r2 = t2.rows[2]
format_cell(t2_r2.cells[0], "Active Air Sampling of Cleanrooms", font_size=6.5)
format_cell(t2_r2.cells[1], "Weekly", font_size=6.5)
format_cell(t2_r2.cells[2], "08MAY 2026", font_size=6.5)
format_cell(t2_r2.cells[3], "SMO", font_size=6.5)
format_cell(t2_r2.cells[4], "Week Before Testing Date", font_size=6.5)
format_cell(t2_r2.cells[5], "1 CFU in ISO 7 115A and\n11 CFUs in ISO 7 115B", font_size=6.5)
format_cell(t2_r2.cells[6], "ETX-260518-0249\nETX-260518-0250", font_size=6.5)
format_cell(t2_r2.cells[7], "Staphylococcus aureus\nPenicillium decumbens\nCladosporium tenuissimum\nCladosporium langeronii\nCladosporium halotolerans", font_size=6.5)
format_cell(t2_r2.cells[8], "None", font_size=6.5)

# Row 3: 14MAY 2026 (Test Date)
t2_r3 = t2.rows[3]
format_cell(t2_r3.cells[0], "Active Air Sampling of Cleanrooms", font_size=6.5)
format_cell(t2_r3.cells[1], "Weekly", font_size=6.5)
format_cell(t2_r3.cells[2], "14MAY 2026", font_size=6.5)
format_cell(t2_r3.cells[3], "SMO", font_size=6.5)
format_cell(t2_r3.cells[4], "Week of Testing Date", font_size=6.5)
format_cell(t2_r3.cells[5], "1 CFU in ISO 8 115 and\n11 CFUs in ISO 7 115B", font_size=6.5)
format_cell(t2_r3.cells[6], "ETX-260526-0391\nETX-260526-0395", font_size=6.5)
format_cell(t2_r3.cells[7], "Gram (+) cocci\nCandida orthopsilosis", font_size=6.5, bold=True)
format_cell(t2_r3.cells[8], "None", font_size=6.5)

# Row 4: 22MAY 2026
t2_r4 = t2.rows[4]
format_cell(t2_r4.cells[0], "Active Air Sampling of Cleanrooms", font_size=6.5)
format_cell(t2_r4.cells[1], "Weekly", font_size=6.5)
format_cell(t2_r4.cells[2], "22MAY 2026", font_size=6.5)
format_cell(t2_r4.cells[3], "SMO", font_size=6.5)
format_cell(t2_r4.cells[4], "Week After Testing Date", font_size=6.5)
format_cell(t2_r4.cells[5], "No growth", font_size=6.5)
format_cell(t2_r4.cells[6], "N/A", font_size=6.5)
format_cell(t2_r4.cells[7], "N/A", font_size=6.5)
format_cell(t2_r4.cells[8], "None", font_size=6.5)

# Section header 2: Row 5
for c in t2.rows[5].cells:
    format_cell(c, "Surface Sampling of Anteroom and Cleanroom Bracketing 115", font_size=7.0, bold=True, bg_hex="E8EEF5")

# Row 6: 08MAY 2026 Surface
t2_r6 = t2.rows[6]
format_cell(t2_r6.cells[0], "Surface Sampling of Cleanrooms", font_size=6.5)
format_cell(t2_r6.cells[1], "Weekly", font_size=6.5)
format_cell(t2_r6.cells[2], "08MAY 2026", font_size=6.5)
format_cell(t2_r6.cells[3], "SMO", font_size=6.5)
format_cell(t2_r6.cells[4], "Week Before Testing Date", font_size=6.5)
format_cell(t2_r6.cells[5], "No growth", font_size=6.5)
format_cell(t2_r6.cells[6], "N/A", font_size=6.5)
format_cell(t2_r6.cells[7], "N/A", font_size=6.5)
format_cell(t2_r6.cells[8], "None", font_size=6.5)

# Row 7: 14MAY 2026 Surface
t2_r7 = t2.rows[7]
format_cell(t2_r7.cells[0], "Surface Sampling of Cleanrooms", font_size=6.5)
format_cell(t2_r7.cells[1], "Weekly", font_size=6.5)
format_cell(t2_r7.cells[2], "14MAY 2026", font_size=6.5)
format_cell(t2_r7.cells[3], "SMO", font_size=6.5)
format_cell(t2_r7.cells[4], "Week of Testing Date", font_size=6.5)
format_cell(t2_r7.cells[5], "1 CFU on table in 115 ISO 8 and\n1 CFU on cart in 115 ISO 8", font_size=6.5)
format_cell(t2_r7.cells[6], "ETX-260526-0424\nETX-260526-0425", font_size=6.5)
format_cell(t2_r7.cells[7], "Gram (+) rods\nStaphylococcus capitis", font_size=6.5)
format_cell(t2_r7.cells[8], "None", font_size=6.5)

# Row 8: 22MAY 2026 Surface
t2_r8 = t2.rows[8]
format_cell(t2_r8.cells[0], "Surface Sampling of Cleanrooms", font_size=6.5)
format_cell(t2_r8.cells[1], "Weekly", font_size=6.5)
format_cell(t2_r8.cells[2], "22MAY 2026", font_size=6.5)
format_cell(t2_r8.cells[3], "SMO", font_size=6.5)
format_cell(t2_r8.cells[4], "Week After Testing Date", font_size=6.5)
format_cell(t2_r8.cells[5], "No growth", font_size=6.5)
format_cell(t2_r8.cells[6], "N/A", font_size=6.5)
format_cell(t2_r8.cells[7], "N/A", font_size=6.5)
format_cell(t2_r8.cells[8], "None", font_size=6.5)

# Save Word Table
out_table_docx_1 = os.path.join(desktop_dir, "EM table OOS-261242 14MAY2026.docx")
out_table_docx_2 = os.path.join(desktop_dir, "Tables OOS-261242 EM SMO 115B Air 14MAY2026 - EM.docx")
doc_tbl.save(out_table_docx_1)
doc_tbl.save(out_table_docx_2)
print("Saved amended table DOCX to Desktop.")

# Convert to PDF via Word COM
out_table_pdf = os.path.join(desktop_dir, "EM table OOS-261242 14MAY2026.pdf")
out_table_pdf_2 = os.path.join(desktop_dir, "Tables OOS-261242 EM SMO 115B Air 14MAY2026 - EM.pdf")

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

# ==========================================
# STEP 3: UPDATE PDF FORM FIELDS
# ==========================================
print("\n=== STEP 3: UPDATE PDF FORM FIELDS ===")
pristine_pdf = os.path.join(history_dir, "OOS_261242_orig_backup.pdf")
if not os.path.exists(pristine_pdf):
    pristine_pdf = src_pdf_path

doc_pdf = fitz.open(pristine_pdf)

field_updates = {
    # Page 1
    "Text Field0": "Simin Mohammad",
    "Text Field3": (
        "Simin Mohammad (Weekly Active Air Sampling Plate Setup)\r"
        "Sophia Santamaria (Weekly Active Air Sampling Plate Reader)"
    ),
    "Text Field8": "MICRO-SOP-2",
    "Text Field9": "23-Jul-2026",
    "Text Field10": "16",
    "Text Field11": "Action level: >= 10 CFU/Plate",
    "Text Field13": "Yes, analysts Simin Mohammad and Sophia Santamaria were comprehensively interviewed.",
    "Text Field15": "Yes, as per MICRO-SOP-2",
    "Text Field16": "Yes, as per MICRO-SOP-2",
    "Text Field21": "Yes, as per MICRO-SOP-2",
    
    # Page 2
    "Text Field32": "CR115 (Sensor E001737)",
    "Text Field43": "Incubator E001034 (Sensor E001501)\rIncubator E001031 (Sensor E001505)",
    "Text Field44": "Aug 2027 / Feb 2027\rAug 2027 / Feb 2027",
    
    # Page 3
    "Text Field49": (
        "The analyst involved in the Active Air Sampling plate setup, Simin Mohammad, and the analyst involved in reading the plate, "
        "Sophia Santamaria, were interviewed comprehensively. Their answers are recorded throughout this document.\n\n"
        "The EM plates were stored in compliance with the supplier's recommendations, and their integrity was visually inspected prior to use. "
        "Furthermore, the plates were confirmed to be within their valid expiration dates. All the supplies were thoroughly disinfected according "
        "to MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology). The functionality of both incubators was verified through a review "
        "of data obtained from our comprehensive in-house continuous monitoring system.\n\n"
        "Active Air Sampling was performed by analyst Simin Mohammad during weekly environmental monitoring processing in Suite 115 (ISO 8), "
        "Suite 115A (ISO 7), and Suite 115B (ISO 7) on 14 May 2026 as per MICRO-SOP-2 (Environmental Monitoring of the Cleanroom Facility).\n\n"
        "The plates were initially incubated at a temperature of 30–35°C in incubator E001031 for a minimum duration of 48 hours, commencing on "
        "14 May 2026. Following completion of a minimum of 48 hours of incubation on 18 May 2026, the plates were further incubated for a minimum of "
        "5 days at 20–25°C in incubator E001034, with the incubation concluding on 26 May 2026. Please see Table 1 for detailed information on the "
        "observations during respective incubations.\n\n"
        "Based on the observations in Table 1, since the CFU count exceeded the action level for one of the Active Air Sampling plates (115B), "
        "the plate was submitted for Microbial Identification under ETX-260526-0395. The colonies were identified as Candida orthopsilosis (Budding yeast).\n\n"
        "To observe if the organisms identified were transient in nature or recurring, weekly environmental monitoring plates for the clean room were "
        "bracketed to include the week before testing, the week of testing, and the week after testing as detailed in Table 2 (please see attached)."
    ),
    
    # Page 4
    "Text Field50": (
        "Environmental Monitoring Bracketing & Trend Assessment:\n"
        "Weekly environmental monitoring plates for the clean room were bracketed to include the week before testing (08-May-2026), "
        "the week of testing (14-May-2026), and the week after testing (22-May-2026) as detailed in Table 2 (attached). The bracketing data "
        "demonstrate that ISO 7 Room 115B had 11 CFUs during the week before testing (08-May-2026, under OOS-261185) and again 11 CFUs "
        "during the week of testing (14-May-2026, under OOS-261242), both reaching/exceeding the established action level of >= 10 CFU/plate. "
        "Acknowledging these two consecutive weekly active-air excursions in Room 115B, the recovered microbial populations were evaluated: "
        "on 08-May-2026, the isolates were filamentous fungal molds identified as Penicillium decumbens, Cladosporium tenuissimum, "
        "Cladosporium langeronii, and Cladosporium halotolerans (along with Staphylococcus aureus), whereas on 14-May-2026, the active air plate "
        "exclusively yielded a budding yeast identified as Candida orthopsilosis (along with Gram (+) cocci in ISO 8 115, and Gram (+) rods and "
        "Staphylococcus capitis on Suite 115 surfaces). While both consecutive weeks exhibited excursions meeting/exceeding the action level in "
        "Room 115B, the recovered organisms belonged to distinctly different microbial classes (filamentous molds vs. budding yeast).\n\n"
        "Analyst Interview & Cleanroom Disinfection:\n"
        "During the comprehensive interview with the setup analyst (Simin Mohammad) and reader (Sophia Santamaria), no procedural deviations, "
        "aseptic breaches, or equipment anomalies were identified. All samples and supplies were disinfected prior to introduction, and Suite 115 "
        "was thoroughly cleaned and prepared before testing as per MICRO-SOP-2 (Environmental Monitoring of the Cleanroom Facility) and "
        "MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology).\n\n"
        "Cleanroom Monthly Disinfection & State of Control:\n"
        "Prior to the event, monthly cleaning and disinfection of the outermost ISO 8 Anteroom (Suite 115), the middle ISO 7 Buffer room "
        "(Suite 115A), the innermost ISO 7 clean room (Suite 115B), and its containing ISO 5 Biosafety Cabinets were performed on 26-Apr-2026 "
        "by analysts Rey Estrada and Tamiru Kotisso as per MICRO-SOP-9. All H2O2 chemical indicator strips passed, confirming established "
        "pre-event facility control. Furthermore, subsequent monthly cleaning and disinfection with H2O2 fogging was performed on 31-May-2026 "
        "by analysts Tamiru Kotisso and Cuong Du with all H2O2 indicators passing, verifying complete facility restoration and ongoing "
        "environmental control. Routine daily and between-session disinfection was actively maintained."
    ),
    
    # Page 5
    "Text Field51": (
        "Evaluation of Concurrent Suite 115 Sterility Samples & Cladosporium halotolerans Correlation:\n"
        "A collective review of cleanroom operations and concurrent sterility testing in Suite 115 was performed for the week of testing:\n"
        "1. On 11-May-2026, sample ETX-251218-0360 was processed for USP <71> sterility testing in BSC E001314 by Guanchen Li. On 26-May-2026, "
        "the FTM vial yielded Staphylococcus lugdunensis (Gram (+) cocci, 2 colonies under ETX-260526-0342), which is taxonomically unrelated to "
        "either the mold or yeast recoveries.\n"
        "2. On 12-May-2026, sample ETX-260508-0478 was processed for USP <71> sterility testing in BSC E001314 by Devanshi Shah. On 19-May-2026, "
        "the TSB vial yielded Cladosporium halotolerans (Hyphae, 2 colonies under ETX-260519-0388).\n\n"
        "Collective Assessment of Environmental and Sterility Data:\n"
        "Significantly, Cladosporium halotolerans was recovered from both the preceding active air monitoring event in Room 115B on 08-May-2026 "
        "and from sterility sample ETX-260508-0478 processed inside Suite 115 on 12-May-2026. Evaluating this relationship collectively indicates "
        "that Cladosporium halotolerans fungal spores had a localized presence within Suite 115 during the second week of May, demonstrating an "
        "adverse environmental trend for that specific mold species during that timeframe (addressed and investigated under its respective "
        "sterility failure investigation). However, in evaluating the 14-May-2026 active air excursion under current investigation, the recovered "
        "isolate was Candida orthopsilosis, an asexual budding yeast that is biologically and phylogenetically distinct from Cladosporium halotolerans. "
        "While the consecutive 11 CFU counts on 08-May and 14-May reflect an elevated bioburden in Room 115B during this period, the transition from "
        "mold to yeast indicates that the 14-May excursion did not arise from persistent colonization or spreading of the Cladosporium mold, but "
        "rather represented a separate, transient yeast recovery.\n\n"
        "Cleanroom Clearance & Phase I Root Cause Conclusion:\n"
        "Subsequent active air sampling performed in Room 115B on 22-May-2026 yielded No Growth (0 CFU), and surface contact plates across Suite 115 "
        "on 22-May-2026 likewise yielded No Growth (0 CFU). Furthermore, the facility-wide monthly disinfection and H2O2 fogging completed on "
        "31-May-2026 (all chemical indicators verified passing) successfully remediated any residual mold or yeast bioburden. Based on these "
        "collective findings, the 14-May active air excursion was non-recurring following routine sanitization, and the cleanroom environment has "
        "returned to a verified state of microbiological control. No additional corrective actions are deemed necessary beyond ongoing adherence "
        "to strict cleanroom disinfection protocols."
    ),
    
    # OOS Number across all pages
    "Text Field57": "261242"
}

for page in doc_pdf:
    for w in page.widgets():
        if w.field_name in field_updates:
            new_val = field_updates[w.field_name]
            # Replace any non-cp1252 character to guarantee appearance stream (/AP) renders in all PDF viewers
            clean_val = str(new_val).replace("₂", "2")
            w.field_value = clean_val
            if w.field_name == "Text Field11":
                w.text_fontsize = 6.5
            w.update()
        elif w.field_name == "Check Box68":
            w.field_value = "Yes"
            w.update()
        elif w.field_name == "Check Box69":
            w.field_value = ""
            w.update()

# ==========================================
# STEP 4: ASSEMBLE 8-PAGE FULL PDF PACKAGE (PyMuPDF preservation)
# ==========================================
print("\n=== STEP 4: ASSEMBLE 8-PAGE FULL PDF PACKAGE ===")
# Append Page 8: Amended 1-page EM Table directly via fitz to preserve AcroForm fields
doc_table = fitz.open(out_table_pdf)
doc_pdf.insert_pdf(doc_table)

final_full_pdf = os.path.join(desktop_dir, "OOS-261242 EM SMO 115B Air 14MAY2026 - EM.pdf")
doc_pdf.save(final_full_pdf)
doc_pdf.close()
doc_table.close()

# Synchronize original Desktop PDFs
shutil.copy2(final_full_pdf, src_pdf_path)
std_pdf_path = os.path.join(desktop_dir, "OOS-261242.pdf")
shutil.copy2(final_full_pdf, std_pdf_path)

print(f"Assembled complete 8-page package to: {final_full_pdf}")

# ==========================================
# STEP 5: GENERATE FULL WORD REPORT
# ==========================================
print("\n=== STEP 5: GENERATE FULL WORD REPORT ===")
report_docx_path = os.path.join(desktop_dir, "OOS-261242 EM SMO 115B Air 14MAY2026 - EM.docx")

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
r_sub = p_sub.add_run("OOS-261242 | EM SMO 115B Air 14May2026 | Active Air Sampling")
r_sub.font.name = "Calibri"
r_sub.font.size = Pt(11)
r_sub.italic = True
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Summary Info Table
info_table = doc_rep.add_table(rows=6, cols=2)
info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
info_data = [
    ("OOS Investigation Number:", "OOS-261242"),
    ("Test & Sample Description:", "Environmental Monitoring - EM SMO 115B Air 14May2026 (Active Air Plate)"),
    ("Testing Site & Location:", "Suite 115 (ISO 8 Anteroom), Suite 115A (ISO 7 Buffer), Suite 115B (ISO 7 Cleanroom)"),
    ("Initiator / Setup Analyst:", "Simin Mohammad (Weekly Active Air Sampling Plate Setup)"),
    ("Reading Analyst:", "Sophia Santamaria (Weekly Active Air Sampling Plate Reader)"),
    ("Action Level & Result:", "Action level: >= 10 CFU/Plate | Result: 11 CFUs (Candida orthopsilosis - Budding yeast)")
]
for idx, (label, val) in enumerate(info_data):
    r = info_table.rows[idx]
    r.cells[0].width = Inches(2.5)
    r.cells[1].width = Inches(4.5)
    format_cell(r.cells[0], label, font_size=9.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, bg_hex="F2F2F2")
    format_cell(r.cells[1], val, font_size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT)

doc_rep.add_paragraph().paragraph_format.space_after = Pt(8)

# Add Narrative Sections
sections_text = [
    ("1. Phase I Summary – Setup, Incubation & Microbial Identification", field_updates["Text Field49"]),
    ("2. Phase I Summary – Cleanroom Bracketing, Analyst Interview & Monthly Cleaning", field_updates["Text Field50"]),
    ("3. Phase I Summary – Cleanroom BSC Samples Assessment & Root Cause Statement", field_updates["Text Field51"])
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
email_docx_out = os.path.join(desktop_dir, "OOS-261242 Review Email.docx")
email_html_out = os.path.join(desktop_dir, "OOS-261242_Review_Email_Robin_Format.html")

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
    "I have reviewed OOS-261242 and made my edits to this and summary of the major ones as below. "
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
        "Environmental Monitoring Conformance (Page 2)",
        "For the question \"Do the environmental monitoring results conform to required standards?\", the selection was marked as N/A.",
        "Revised selection to \"No\" (Check Box68), as the investigation was initiated due to an active-air result of 11 CFU in ISO 7 Room 115B reaching/exceeding the established action level of >= 10 CFU/plate."
    ],
    [
        "Equipment/Facility Description (Page 2)",
        "Only lists CR115 (E001737) in equipment section. In Section D, incubator IDs and calibration dates ran together without proper line spacing: \"Incubator E001034(Sensor E001501)Incubator E001031 (Sensor E001505)\" and \"Aug-2026 Feb-2027 Aug-2026 Feb-2027\".",
        "Added full cleanroom facility identifiers for Suite 115 (ISO 8 Anteroom), Suite 115A (ISO 7 Buffer room), and Suite 115B (ISO 7 Cleanroom) with Sensor E001737. Formatted Incubator E001034 (Sensor E001501) and Incubator E001031 (Sensor E001505) with clean line breaks and calibration dates (Aug 2027 / Feb 2027)."
    ],
    [
        "Reading Analyst, Processing Analyst & SOP Description (Page 1)",
        "Combined analyst names on a single line: \"Simin Mohammad(Weekly Active air Sampling Plate Setup)Sophia Santamaria(Weekly Active air Sampling Plate Readers\" (missing closing parenthesis and plural typo). Cited legacy \"MICRO-SOP-2\" Rev 15 (05-AUG-2025). Stated \"analysts Simin Mohammad, & Sophia Santamaria\" with redundant ampersand. In Section D, misspelled reader as \"Sophia Sanatamaria\", contained grammar error \"incubators were verified\", and erroneously cited \"MICRO-SOP-2 (Cleaning and Disinfecting Procedure for Microbiology)\".",
        "Delineated setup analyst Simin Mohammad (Weekly Active Air Sampling Plate Setup) and reader Sophia Santamaria (Weekly Active Air Sampling Plate Reader). Updated procedure to current MICRO-SOP-2 Rev 16 (Effective Date: 23-Jul-2026). Revised interview comment to \"Simin Mohammad and Sophia Santamaria\". Corrected spelling to Sophia Santamaria, grammar to incubator was verified, and corrected cleaning procedure reference to MICRO-SOP-9."
    ],
    [
        "Limits / Specification (Page 1)",
        "Listed as \"Action level: >10CFU/plate\" without proper spacing or cGMP equality symbol.",
        "Standardized to \"Action level: >= 10 CFU/Plate\"."
    ],
    [
        "Bracketing EM & Trend Assessment (Page 4)",
        "Described the 14-May active air excursion as \"an isolated, transient event\" without acknowledging that the preceding week (08-May-2026) in Room 115B also had an active-air excursion of 11 CFU. Referenced only 31-May monthly cleaning without citing preceding April monthly cleaning.",
        "Acknowledged two consecutive weekly active-air excursions reaching/exceeding the action level (11 CFU on 08-May-2026 under OOS-261185 and 11 CFU on 14-May-2026 under OOS-261242 in Room 115B). Evaluated the taxonomic shift between filamentous fungal molds (Penicillium decumbens, Cladosporium spp.) on 08-May and asexual budding yeast (Candida orthopsilosis) on 14-May. Cited both preceding 26-Apr-2026 and subsequent 31-May-2026 monthly H2O2 decontamination logs."
    ],
    [
        "Concurrent Sterility Samples & Mold Correlation (Page 5)",
        "Separately listed concurrent sterility test hits without evaluating the shared recovery of Cladosporium halotolerans between the 08-May active air plate and 12-May sterility sample ETX-260508-0478.",
        "Evaluated the shared Cladosporium halotolerans recovery collectively as a localized adverse environmental presence for that mold species in mid-May (investigated under sterility failure), distinguished it from the 14-May Candida orthopsilosis yeast excursion (confirming no persistent mold colonization), and cited subsequent cleanroom clearance on 22-May-2026 (0 CFU / No Growth on active air and surfaces) and verified restoration of environmental control."
    ],
    [
        "EM Table Formatting (Page 8)",
        "In the draft Word document, Table 1 date had excessive spacing (\"14MAY   2026\"), CFU count lacked spacing (\"11CFUs\"), Table 2 Row 7 had plural typo (\"1 CFUs on table  in 115 ISO 8 and 1 CFUs on cart in 115 ISO 8\"), and multi-organism entries lacked line breaks.",
        "Amended both tables to correct spacing (14MAY 2026, 11 CFUs), corrected plural typos (1 CFU on table in 115 ISO 8 and 1 CFU on cart in 115 ISO 8), formatted multi-organism line breaks, and standardized bracketing section headers."
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
<p style="margin: 0 0 12px 0; font-family: Calibri, sans-serif; font-size: 11pt;">I have reviewed OOS-261242 and made my edits to this and summary of the major ones as below. Please also note that the highlighted sections in the table have been amended:</p>
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

# Also render preview of assembled 8-page PDF
doc_v = fitz.open(final_full_pdf)
print(f"Final assembled PDF page count: {len(doc_v)}")
for i, page in enumerate(doc_v):
    pix = page.get_pixmap(dpi=150)
    pix.save(os.path.join(scratch_dir, f"assembled_261242_p{i+1}.png"))
print("Saved all 8 page preview images to scratch/")
