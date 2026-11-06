"""
build_celsis_oos262333_master_delivery.py
Complete, definitive builder for OOS-262333 Celsis Deliverables:
1. Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx (Strictly 3 Pages)
2. CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf (Strictly 9 Pages)
"""

import docx
import os
import win32com.client
import fitz
from pypdf import PdfReader, PdfWriter
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BENCHMARK_DOCX = r"scratch/qa_celsis_emails/262316_MILLS EM TABLE.docx"
BLANK_ACROFORM_PDF = r"scratch/qa_celsis_emails/262333_CORP-FORM-21 Laboratory OOS Investigation Form (v11.1) (5).pdf"

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = r"scratch"

DOCX_DELIVERABLE_DESKTOP = os.path.join(DESKTOP_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx")
DOCX_DELIVERABLE_DOCS = os.path.join(DOCS_DIR, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx")

TABLES_TEMP_PDF = os.path.abspath(os.path.join(SCRATCH_DIR, "temp_tables_oos262333_3p_final.pdf"))

PDF_DELIVERABLE_DESKTOP = os.path.join(DESKTOP_DIR, "CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf")
PDF_DELIVERABLE_DOCS = os.path.join(DOCS_DIR, "CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf")

# ==============================================================================
# PART 1: BUILD 3-PAGE STANDALONE DOCX FROM BENCHMARK TEMPLATE
# ==============================================================================
print("\n=== [PART 1] BUILDING STANDALONE DOCX TABLES ===")

doc = docx.Document(BENCHMARK_DOCX)

# 1. Remove floating table layout (w:tblpPr) from Table 2 and Table 3 to make them strictly inline
for t in doc.tables:
    tblPr = t._tbl.tblPr
    for p in tblPr.xpath('w:tblpPr'):
        tblPr.remove(p)

# 2. Add left indent -551 dxa (-27.55pt) to Table 2 and Table 3 so left border sits at x=50.0pt and right at x=593pt (centered inside 612pt page)
for i in [1, 2]:
    t = doc.tables[i]
    tblPr = t._tbl.tblPr
    for ind in tblPr.xpath('w:tblInd'):
        tblPr.remove(ind)
    tblPr.append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="-551" w:type="dxa"/>'))

# 3. Update Title Paragraphs and configure pageBreakBefore
p_t1 = None
p_t2 = None
p_t3 = None

for p in doc.paragraphs:
    if "Table 1:" in p.text:
        p_t1 = p
        p.text = "Table 1: Reading analyst, Days of incubation, and Microbial ID for Sample under Investigation"
        p.runs[0].bold = True
        p.runs[0].font.name = "Times New Roman"
    elif "Table 2:" in p.text:
        p_t2 = p
        p.text = "Table 2: Environmental Monitoring from Celsis Processing Performed on 05Oct2026"
        p.runs[0].bold = True
        p.runs[0].font.name = "Times New Roman"
        pPr = p._p.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:pageBreakBefore {nsdecls("w")}/>'))
    elif "Table 3:" in p.text:
        p_t3 = p
        p.text = "Table 3: Environmental Monitoring from Celsis Aliquoting Performed on 12Oct2026"
        p.runs[0].bold = True
        p.runs[0].font.name = "Times New Roman"
        pPr = p._p.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:pageBreakBefore {nsdecls("w")}/>'))

# 4. Remove all empty spacing paragraphs that were used as manual page breaks
body = doc.element.body
for p in list(body.xpath('w:p')):
    text = ''.join(p.xpath('.//w:t/text()')).strip()
    if not text:
        body.remove(p)

# Helper functions for clean cell styling
def update_cell_text(cell, text, font_size_pt=7.0, bold=False, italic=False):
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    p = docx.text.paragraph.Paragraph(parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>'), cell)
    lines = text.split('\n')
    for idx, line in enumerate(lines):
        r = p.add_run(line)
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size_pt)
        r.bold = bold
        r.italic = italic
        if idx < len(lines) - 1:
            p.add_run('\n')
    tc.append(p._p)

def update_cell_hyperlink(cell, url, text, font_size_pt=7.0, bold=False):
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    p = docx.text.paragraph.Paragraph(parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>'), cell)
    part = cell.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{r_id}"/>')
    new_run = parse_xml(
        f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'{"<w:b/>" if bold else ""}'
        f'<w:color w:val="0563C1"/>'
        f'<w:u w:val="single"/>'
        f'<w:sz w:val="{int(font_size_pt*2)}"/>'
        f'<w:szCs w:val="{int(font_size_pt*2)}"/>'
        f'</w:rPr>'
        f'<w:t>{text}</w:t>'
        f'</w:r>'
    )
    hyperlink.append(new_run)
    p._p.append(hyperlink)
    tc.append(p._p)

def update_cell_multispecies(cell, species_list, font_size_pt=6.5):
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    p = docx.text.paragraph.Paragraph(parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="180" w:lineRule="auto"/></w:pPr></w:p>'), cell)
    for idx, sp in enumerate(species_list):
        r = p.add_run(sp)
        r.font.name = "Times New Roman"
        r.font.size = Pt(font_size_pt)
        r.italic = True
        if idx < len(species_list) - 1:
            p.add_run('\n\n')
    tc.append(p._p)

# 5. Populate Table 1 (Sample Under Investigation)
t1 = doc.tables[0]
update_cell_text(t1.rows[1].cells[0], "CCD", font_size_pt=9.0)
update_cell_text(t1.rows[1].cells[1], "Day 7", font_size_pt=9.0)
update_cell_hyperlink(t1.rows[1].cells[2], "https://etrax.eagleanalytical.com/Submission/Details/Q57eK3N9WcR0Qh%24x6N5N6Q__", "ETX-261002-0288", font_size_pt=9.0)
update_cell_text(t1.rows[1].cells[3], "ETX-261012-0654", font_size_pt=9.0)
update_cell_text(t1.rows[1].cells[4], "1 x 300mL FTM", font_size_pt=9.0)

tc5 = t1.rows[1].cells[5]._tc
for p in tc5.xpath('w:p'):
    tc5.remove(p)
p_sp = docx.text.paragraph.Paragraph(parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>'), t1.rows[1].cells[5])
r_sp = p_sp.add_run("Bacillus paralicheniformis\n")
r_sp.font.name = "Times New Roman"
r_sp.font.size = Pt(9.0)
r_sp.italic = True
r_gr = p_sp.add_run("(Gram-positive rod)")
r_gr.font.name = "Times New Roman"
r_gr.font.size = Pt(8.5)
tc5.append(p_sp._p)

# 6. Populate Table 2 (Processing EM Bracketing 05Oct2026)
t2 = doc.tables[1]
update_cell_text(t2.rows[1].cells[0], "Personnel EM Bracketing 05Oct2026", font_size_pt=7.0, bold=True)
update_cell_text(t2.rows[2].cells[2], "02Oct2026")
update_cell_text(t2.rows[2].cells[3], "ES")
update_cell_text(t2.rows[3].cells[2], "05Oct2026")
update_cell_text(t2.rows[3].cells[3], "ES")
update_cell_text(t2.rows[4].cells[2], "06Oct2026")
update_cell_text(t2.rows[4].cells[3], "ES")

update_cell_text(t2.rows[5].cells[0], "Biological Safety Cabinet EM Bracketing E001316 for 05Oct2026", font_size_pt=7.0, bold=True)
update_cell_text(t2.rows[6].cells[2], "02Oct2026")
update_cell_text(t2.rows[6].cells[3], "ALA")
update_cell_text(t2.rows[7].cells[2], "05Oct2026")
update_cell_text(t2.rows[7].cells[3], "ES")
update_cell_text(t2.rows[8].cells[2], "06Oct2026")
update_cell_text(t2.rows[8].cells[3], "ES")

update_cell_text(t2.rows[9].cells[2], "02Oct2026")
update_cell_text(t2.rows[9].cells[3], "ALA")
update_cell_text(t2.rows[10].cells[2], "05Oct2026")
update_cell_text(t2.rows[10].cells[3], "ES")
update_cell_text(t2.rows[11].cells[2], "06Oct2026")
update_cell_text(t2.rows[11].cells[3], "ES")

update_cell_text(t2.rows[12].cells[0], "Weekly Active Air and Surface Sampling of Cleanroom Suite 114 with Processing BSC For 05Oct2026", font_size_pt=7.0, bold=True)
update_cell_text(t2.rows[14].cells[2], "05Oct2026")
update_cell_text(t2.rows[14].cells[3], "SMO")
update_cell_text(t2.rows[14].cells[4], "No Growth")
update_cell_text(t2.rows[14].cells[5], "Not Applicable")
update_cell_text(t2.rows[14].cells[7], "Not Applicable")
update_cell_text(t2.rows[14].cells[9], "None")

update_cell_text(t2.rows[15].cells[2], "05Oct2026")
update_cell_text(t2.rows[15].cells[3], "SMO")
update_cell_text(t2.rows[15].cells[4], "No Growth")
update_cell_text(t2.rows[15].cells[5], "Not Applicable")
update_cell_text(t2.rows[15].cells[7], "Not Applicable")
update_cell_text(t2.rows[15].cells[9], "None")

# 7. Populate Table 3 (Aliquoting EM Bracketing 12Oct2026)
t3 = doc.tables[2]
update_cell_text(t3.rows[1].cells[0], "Personnel EM Bracketing 12Oct2026", font_size_pt=7.0, bold=True)
update_cell_text(t3.rows[2].cells[2], "09Oct2026")
update_cell_text(t3.rows[2].cells[4], "Date Before Aliquoting")
update_cell_text(t3.rows[3].cells[2], "12Oct2026")
update_cell_text(t3.rows[3].cells[4], "Date of Aliquoting")
update_cell_text(t3.rows[4].cells[2], "13Oct2026")
update_cell_text(t3.rows[4].cells[4], "Date After Aliquoting")

update_cell_text(t3.rows[5].cells[0], "Biological Safety Cabinet EM for 12Oct2026", font_size_pt=7.0, bold=True)
update_cell_text(t3.rows[6].cells[2], "09Oct2026")
update_cell_text(t3.rows[6].cells[4], "Date Before Aliquoting")
update_cell_text(t3.rows[7].cells[2], "12Oct2026")
update_cell_text(t3.rows[7].cells[4], "Date of Aliquoting")
update_cell_text(t3.rows[8].cells[2], "13Oct2026")
update_cell_text(t3.rows[8].cells[4], "Date After Aliquoting")

update_cell_text(t3.rows[9].cells[2], "09Oct2026")
update_cell_text(t3.rows[9].cells[4], "Date Before Aliquoting")
update_cell_text(t3.rows[10].cells[2], "12Oct2026")
update_cell_text(t3.rows[10].cells[4], "Date of Aliquoting")
update_cell_text(t3.rows[11].cells[2], "13Oct2026")
update_cell_text(t3.rows[11].cells[4], "Date After Aliquoting")

update_cell_text(t3.rows[12].cells[0], "Weekly Active Air and Surface Sampling of Cleanroom Suite 114 with Processing BSC For 12Oct2026", font_size_pt=7.0, bold=True)
update_cell_text(t3.rows[14].cells[2], "12Oct2026")
update_cell_text(t3.rows[14].cells[3], "SMO")
update_cell_text(t3.rows[14].cells[4], "6 CFUs\n(ISO 8 114)")
update_cell_hyperlink(t3.rows[14].cells[5], "https://etrax.eagleanalytical.com/SubmissionTest/Details/Cw%24QZhYAogGUwoutnOavpw__", "ETX-261021-0578")

species_list = [
    "Moraxella osloensis",
    "Sporisorium graminicola",
    "Staphylococcus haemolyticus",
    "Staphylococcus hominis",
    "Kocuria indica",
    "Ustilago maydis",
    "Kocuria rhizophila"
]
update_cell_multispecies(t3.rows[14].cells[7], species_list, font_size_pt=6.5)
update_cell_text(t3.rows[14].cells[9], "None")

update_cell_text(t3.rows[15].cells[2], "12Oct2026")
update_cell_text(t3.rows[15].cells[3], "SMO")
update_cell_text(t3.rows[15].cells[4], "No Growth")
update_cell_text(t3.rows[15].cells[5], "Not Applicable")
update_cell_text(t3.rows[15].cells[7], "Not Applicable")
update_cell_text(t3.rows[15].cells[9], "None")

# Save DOCX deliverables
doc.save(DOCX_DELIVERABLE_DESKTOP)
doc.save(DOCX_DELIVERABLE_DOCS)
print(f"Saved DOCX deliverable: {DOCX_DELIVERABLE_DESKTOP}")
print(f"Saved DOCX deliverable: {DOCX_DELIVERABLE_DOCS}")

# Export Standalone Tables PDF via Word
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
wdoc = word.Documents.Open(os.path.abspath(DOCX_DELIVERABLE_DESKTOP))
wdoc.SaveAs(TABLES_TEMP_PDF, FileFormat=17)
wdoc.Close()
word.Quit()
print(f"Exported Standalone Tables PDF: {TABLES_TEMP_PDF}")

# Verify 3-page layout
p_tables = fitz.open(TABLES_TEMP_PDF)
print(f"VERIFIED STANDALONE TABLES PDF PAGE COUNT: {len(p_tables)} pages.")
assert len(p_tables) == 3, f"FATAL ERROR: Standalone tables PDF has {len(p_tables)} pages instead of 3!"

for pno in range(3):
    page = p_tables[pno]
    print(f"=== Tables PDF Page {pno+1} ===")
    for b in page.get_text("blocks"):
        text = b[4].strip().replace('\n', ' ')
        if any(k in text for k in ['Table 1', 'Table 2', 'Table 3', 'Reading', 'Personnel', 'SMO', 'Bacillus']):
            print(f"  bbox=({b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}) text={text[:65]}")

p_tables.close()

# ==============================================================================
# PART 2: FILL 6-PAGE ACROFORM PDF WITH AUTO-SIZED / ZERO-OVERFLOW TYPOGRAPHY
# ==============================================================================
print("\n=== [PART 2] FILLING 6-PAGE ACROFORM PDF ===")

doc_acro = fitz.open(BLANK_ACROFORM_PDF)
assert len(doc_acro) == 6, f"Expected 6 pages blank AcroForm, got {len(doc_acro)}"

# Personnel block: 4 analysts, each with name and title
personnel_block = (
    "Mukyung Jang (MJ)\r"
    "(Celsis Sterility Prepping Analyst)\r \r"
    "Alex Saravia (ES)\r"
    "(Celsis Sterility Processing Analyst)\r \r"
    "Cuong Du (CCD)\r"
    "(Celsis Sterility Aliquoting Analyst)\r \r"
    "Cuong Du (CCD)\r"
    "(Celsis Reading Analyst)"
)

# Text Field values
tf22_materials = (
    "Acidified Bleach: Lot: 08132026-1\r"
    "70% Sterile IPA: Lot: 26F05001\r"
    "Media:\r"
    "FTM (300mL): Lot: 08132026-1\r"
    "TSB (300mL): Lot: 08132026-1\r"
    "Celsis Reagents:\r"
    "LuminEX: Lot: 0001099839\r"
    "LuminATE: Lot: 0001103038\r"
    "LuminALL: Lot: 0001100346\r"
    "Celsis ATP Positive Control: Lot: 1220511483\r"
    "Instrument Blank / Microplate:\r"
    "96-well Microplate: Lot: 10424579\r"
    "Supplies:\r"
    "Sterile Pipette Tips: Lot: 19124-424C4-01\r"
    "Sterile Syringes: Lot: 240411\r"
    "Sterile Needles: Lot: 240212\r"
    "Tween 80: Lot: 24E28001"
)

tf23_expirations = (
    "Acidified Bleach: Exp: 12Aug2027\r"
    "70% Sterile IPA: Exp: 30Jun2028\r"
    "Media:\r"
    "FTM: Exp: 11Nov2026\r"
    "TSB: Exp: 11Nov2026\r"
    "Celsis Reagents:\r"
    "LuminEX: Exp: 28Feb2027\r"
    "LuminATE: Exp: 30Nov2026\r"
    "LuminALL: Exp: 31Dec2026\r"
    "Celsis ATP Positive Control: Exp: Jan 2027\r"
    "Instrument Blank / Microplate:\r"
    "96-well Microplate: Exp: 31Jul2028\r"
    "Supplies:\r"
    "Sterile Pipette Tips: Exp: 31Aug2028\r"
    "Sterile Syringes: Exp: 31Mar2029\r"
    "Sterile Needles: Exp: 28Feb2029\r"
    "Tween 80: Exp: 31May2027"
)

narrative_p3 = (
    "On 05Oct26, the sample vials for ETX-261002-0288 were retrieved from controlled refrigerated storage and transferred into the sterile microbiology laboratory suites. In accordance with standard operating procedures, container surfaces were sprayed with acidified bleach disinfectant, placed into pre-disinfected transfer bins, and allowed a verified 10-minute contact time. Secondary surface disinfection with acidified bleach was conducted in the ISO 8 anteroom 114 prior to entering cleanroom 114B. Inside ISO 7 cleanroom 114B, processing analyst Alex Saravia (ES) executed a final disinfection step before introducing the containers into the ISO 5 BSC E001316. Direct inoculation was performed by adding 15 mL of sample (divided across ten 3 mL vials) along with 3 mL of sterile Tween 80 into 300 mL of Fluid Thioglycollate Medium (FTM) and 300 mL of Tryptic Soy Broth (TSB) in accordance with method suitability study ETX-250826-0390. A matching negative control jar with identical modifications was prepared concurrently and handled identically. Following inoculation, media bottles were transferred into calibrated incubators to commence incubation."
)

narrative_p4 = (
    "Following instrument analysis on 12Oct26, sample ETX-261002-0288 was found to yield positive luminescence readings in the 1 x 300 mL FTM media jar. The FTM sample yielded an average luminescence of 3,647,568 RLU (duplicate tube readings of 3,608,783 RLU and 3,686,352 RLU, CV% 1.5%), significantly exceeding the positive cut-off threshold of 2,596.5 RLU. In contrast, the 1 x 300 mL TSB media jar yielded negative readings well below its positive cut-off threshold of 2,326.5 RLU. All instrument system controls and background checks met defined acceptance criteria, including the instrument blank (29 RLU, cut-off 100.0 RLU), reagent blank (103 RLU, cut-off 1,000.0 RLU), and Celsis ATP Positive Control (99,973 RLU, cut-off 50,000.0 RLU, passing specification).\n\n"
    "Upon identification of the elevated luminescence reading, the positive FTM media container was immediately subcultured into fresh isolation media per MICRO-SOP-44 and incubated under aerobic conditions to isolate viable organisms. Turbidity was observed in the subculture vessel, and an isolated colony was submitted for microbial identification under tracking ID ETX-261012-0654. Confirmatory identification identified the isolate as Bacillus paralicheniformis (Gram-positive rod). All equipment used during processing and aliquoting—including primary ISO 5 BSCs E001316 and E001798, Celsis Advance II Luminometer E002222, and incubators E001356, E001357, E001034, and E001031—were fully calibrated, certified, and operating within validated parameters with zero alarms or excursions.\n\n"
    "All culture media utilized were within their approved manufacturer expiration dates. Negative control media jars were prepared, manipulated, and incubated alongside the test sample under identical conditions, and all negative controls exhibited no microbial recovery (FTM negative control 866 RLU, TSB negative control 776 RLU)."
)

narrative_p5 = (
    "A thorough review of testing records, training files, and equipment documentation confirmed that all analysts involved in sample preparation, processing, aliquoting, and instrument reading were fully trained and qualified in accordance with MICRO-SOP-44. Comprehensive interviews conducted with prepping analyst Mukyung Jang (MJ), processing analyst Alex Saravia (ES), and aliquoting/reading analyst Cuong Du (CCD) confirmed strict adherence to aseptic technique, proper gowning procedures, and standard operating protocols throughout all handling steps. No procedural anomalies, technique compromises, spills, or unusual occurrences were identified during any phase of testing.\n\n"
    "Analyzing a 6-month sample history for Optimal Balance Pharmacy, this specific analyte \"TESAMORELIN 8 MG/ML (3 ML) Injection\" has had no prior Out-of-Specification (OOS) occurrences under Celsis Sterility Testing.\n\n"
    "Upon review of the environmental monitoring data, absence of microbial recovery (No Growth) was demonstrated across all critical ISO 5 zones and personnel monitoring on the date of testing.\n\n"
    "Personnel environmental monitoring performed on 05Oct26 for processing analyst Alex Saravia (ES) and on 12Oct26 for aliquoting analyst Cuong Du (CCD) demonstrated No Growth across all touch plates. Daily surface and settling plate monitoring of ISO 5 BSC E001316 (used for processing on 05Oct26) and ISO 5 BSC E001798 (used for aliquoting on 12Oct26) yielded No Growth across all locations. Weekly active air and surface monitoring performed on 05Oct26 across Cleanroom Suite 114 also demonstrated No Growth.\n\n"
    "During weekly active air monitoring performed on 12Oct26 across Cleanroom Suite 114, recovery of 6 CFUs was observed on an active air plate located in ISO 8 Anteroom 114 (ETX-261021-0578). Confirmatory identification of these isolates identified Moraxella osloensis, Sporisorium graminicola, Staphylococcus haemolyticus, Staphylococcus hominis, Kocuria indica, Ustilago maydis, and Kocuria rhizophila. Importantly, none of the isolates recovered from the ISO 8 anteroom active air plate matched Bacillus paralicheniformis, the Gram-positive rod recovered from the failing FTM sample. Furthermore, weekly surface monitoring of Cleanroom Suite 114 performed concurrently on 12Oct26 demonstrated No Growth.\n\n"
    "Monthly cleaning and disinfection of cleanrooms (ISO 7/8) and biological safety cabinets (ISO 5) were performed per MICRO-SOP-9, with all chemical indicators passing, confirming that environmental sanitization remained validated and in control.\n\n"
    "In conclusion, comprehensive investigation into the testing process, equipment, media, personnel technique, environmental monitoring, and sample history identified no laboratory error or testing aberration contributing to the elevated luminescence result. Phase I laboratory investigation confirms that the positive Celsis Sterility Test result is valid. Per laboratory procedures, Phase II investigation is initiated to evaluate client manufacturing and formulation factors."
)

pdf_field_values = {
    'Date Field0': "05-Oct-2026",
    'Date Field1': "12-Oct-2026",
    'Date Field2': "12-Oct-2026",
    'Date Field3': "12-Oct-2026",
    'Text Field0': "Cuong Du (Written by Qiyue Chen)",
    'Text Field1': "Celsis Sterility Test",
    'Text Field2': "TESAMORELIN 8 MG/ML (3 ML) Injection",
    'Text Field3': personnel_block,
    'Text Field4': "Optimal Balance Pharmacy (E19193)",
    'Text Field5': "LG342011076",
    'Text Field6': "19193",
    'Text Field7': "On 12 Oct 2026, sample ETX-261002-0288 was found to be positive for growth after Celsis Sterility Testing\r",
    'Text Field8': "MICRO-SOP-44",
    'Text Field9': "03-Jan-2025",
    'Text Field10': "01",
    'Text Field11': "Pass/Fail",
    'Text Field12': "Kathan Parikh",
    'Text Field13': "Yes, analysts Mukyung Jang, Alex Saravia and Cuong Du were comprehensively interviewed",
    'Text Field14': "Yes, Sample ID: ETX-261002-0288",
    'Text Field15': "Yes, as per MICRO-SOP-44",
    'Text Field16': "Yes, as per MICRO-SOP-44",
    'Text Field17': "Yes, Information is available in Eagle Trax under ETX-261002-0288",
    'Text Field18': "Yes, all analysts are trained and qualified by the quality department to perform the test",
    'Text Field19': "Not Applicable",
    'Text Field20': "Not Applicable",
    'Text Field21': "Yes, the samples were stored refrigerated as per client's instructions",
    'Text Field22': tf22_materials,
    'Text Field23': tf23_expirations,
    'Text Field24': "Celsis ATP Positive Control\r",
    'Text Field25': "1220511483",
    'Text Field26': "Jan 2027",
    'Text Field27': "Not Applicable ",
    'Text Field28': "Not Applicable ",
    'Text Field29': "Not Applicable ",
    'Text Field30': "E002222",
    'Text Field31': "Dec 2026",
    'Text Field32': "E001737 (CR114), BSC E001316 \rE001736 (CR114), BSC E001798",
    'Text Field33': "Dec 2026, Jun 2027\rDec 2026, Jun 2027",
    'Text Field34': "E002222",
    'Text Field35': "Dec 2026",
    'Text Field36': "Not Applicable",
    'Text Field37': "Not Applicable",
    'Text Field38': "Not Applicable",
    'Text Field39': "Not Applicable",
    'Text Field40': "See Phase I Summary",
    'Text Field41': "See Phase I Summary",
    'Text Field42': "See Phase I Summary",
    'Text Field43': "Incubator E001356 \r(Monitored by\rSensor E001450)\r \rIncubator E001357\r(Monitored by \rSensor E001449)\r \rIncubator E001034\r(Monitored by \rSensor E001501)\r \r \rIncubator E001031\r(Monitored by\rSensor E001505)\r \rISO 5 BSC E001316\r \rISO 5 BSC E001798",
    'Text Field44': "Jan 2027\rFeb 2027\r \rJan 2027\rFeb 2027\r \rAug 2027\rFeb 2027\r \r \rAug 2027\rFeb 2027\r \rJun 2027\r \rJun 2027",
    'Text Field45': "See Phase I Summary",
    'Text Field46': "Not Applicable",
    'Text Field47': "Not Applicable",
    'Text Field48': "Not Applicable\r",
    'Text Field49': narrative_p3,
    'Text Field50': narrative_p4,
    'Text Field51': narrative_p5,
    'Text Field53': "Qiyue Chen",
    'Text Field54': "Robin Seymour",
    'Text Field57': "262333",
}

cb_whitelists = {
    'Check Box0': 'Yes', 'Check Box1': 'Yes', 'Check Box2': 'Yes',
    'Check Box3': 'Off', 'Check Box4': 'Yes', 'Check Box5': 'Off',
    'Check Box6': 'Off', 'Check Box7': 'Yes', 'Check Box8': 'Off',
    'Check Box9': 'Off', 'Check Box10': 'Yes', 'Check Box11': 'Off',
    'Check Box12': 'Off', 'Check Box13': 'Yes', 'Check Box14': 'Off',
    'Check Box15': 'Off', 'Check Box16': 'Yes', 'Check Box17': 'Off',
    'Check Box18': 'Off', 'Check Box19': 'Yes', 'Check Box20': 'Off',
    'Check Box21': 'Off', 'Check Box22': 'Off', 'Check Box23': 'Off',
    'Check Box24': 'Yes', 'Check Box25': 'Off', 'Check Box26': 'Off',
    'Check Box27': 'Yes', 'Check Box28': 'Yes', 'Check Box29': 'Off',
    'Check Box30': 'Off', 'Check Box31': 'Off', 'Check Box32': 'Yes',
    'Check Box33': 'Off', 'Check Box34': 'Yes', 'Check Box35': 'Off',
    'Check Box36': 'Yes', 'Check Box37': 'Off', 'Check Box38': 'Yes',
    'Check Box39': 'Off', 'Check Box40': 'Off', 'Check Box41': 'Off',
    'Check Box42': 'Yes', 'Check Box43': 'Yes', 'Check Box44': 'Off',
    'Check Box45': 'Off', 'Check Box46': 'Off', 'Check Box47': 'Off',
    'Check Box48': 'Yes', 'Check Box49': 'Off', 'Check Box50': 'Off',
    'Check Box51': 'Yes', 'Check Box52': 'Yes', 'Check Box53': 'Off',
    'Check Box54': 'Off', 'Check Box55': 'Yes', 'Check Box56': 'Off',
    'Check Box57': 'Off', 'Check Box58': 'Yes', 'Check Box59': 'Off',
    'Check Box60': 'Off', 'Check Box61': 'Off', 'Check Box62': 'Off',
    'Check Box63': 'Yes', 'Check Box64': 'Off', 'Check Box65': 'Off',
    'Check Box66': 'Yes', 'Check Box67': 'Yes', 'Check Box68': 'Off',
    'Check Box69': 'Off', 'Check Box70': 'Yes', 'Check Box71': 'Off',
    'Check Box72': 'Off', 'Check Box73': 'Yes', 'Check Box74': 'Off',
    'Check Box75': 'Off', 'Check Box76': 'Off', 'Check Box77': 'Off',
    'Check Box78': 'Yes', 'Check Box79': 'Yes', 'Check Box80': 'Off',
    'Check Box81': 'Off', 'Check Box82': 'Off', 'Check Box83': 'Off',
    'Check Box84': 'Off',  # STRICTLY OFF!
    'Check Box85': 'Off', 'Check Box86': 'Off', 'Check Box87': 'Off',
    'Check Box88': 'Yes', 'Check Box89': 'Off', 'Check Box90': 'Off',
    'Check Box91': 'Off', 'Check Box92': 'Off',
}

# Typography font sizes:
# Text Field3 is calibrated to 5.5pt (verified in crop_field3_5.5.png: 0 wrap, 0 overflow, 0 [+] badge!)
# All other narrative & multi-line fields use optimal calibrated sizes
font_sizes = {
    'Date Field0': 10.0,
    'Date Field1': 10.0,
    'Date Field2': 10.0,
    'Date Field3': 10.0,
    'Text Field0': 5.5,
    'Text Field1': 8.0,
    'Text Field2': 10.0,
    'Text Field3': 5.5,  # 5.5pt: perfectly fits 4 analysts & titles without line wrapping, eliminates [+] badge completely
    'Text Field4': 8.0,
    'Text Field5': 9.5,
    'Text Field6': 10.0,
    'Text Field7': 8.0,
    'Text Field8': 9.5,
    'Text Field9': 9.5,
    'Text Field10': 9.5,
    'Text Field11': 9.5,
    'Text Field12': 10.0,
    'Text Field13': 5.0,
    'Text Field14': 9.0,
    'Text Field15': 9.0,
    'Text Field16': 9.0,
    'Text Field17': 6.725,
    'Text Field18': 5.0,
    'Text Field19': 9.0,
    'Text Field20': 9.0,
    'Text Field21': 5.85,
    'Text Field22': 4.8,
    'Text Field23': 4.8,
    'Text Field24': 9.5,
    'Text Field25': 9.5,
    'Text Field26': 9.5,
    'Text Field27': 9.5,
    'Text Field28': 9.5,
    'Text Field29': 9.5,
    'Text Field30': 9.5,
    'Text Field31': 9.5,
    'Text Field32': 4.825,
    'Text Field33': 4.825,
    'Text Field34': 9.5,
    'Text Field35': 9.5,
    'Text Field36': 9.5,
    'Text Field37': 9.5,
    'Text Field38': 9.5,
    'Text Field39': 9.5,
    'Text Field40': 8.0,
    'Text Field41': 8.0,
    'Text Field42': 8.0,
    'Text Field43': 6.8,
    'Text Field44': 6.8,
    'Text Field45': 8.0,
    'Text Field46': 9.5,
    'Text Field47': 9.5,
    'Text Field48': 8.5,
    'Text Field49': 8.6,
    'Text Field50': 9.0,
    'Text Field51': 8.5,
    'Text Field53': 10.0,
    'Text Field54': 10.0,
    'Text Field57': 6.5,
}

for page_idx in range(6):
    page = doc_acro[page_idx]
    for widget in page.widgets():
        name = widget.field_name
        if name in pdf_field_values:
            widget.field_value = pdf_field_values[name]
            widget.text_font = "Helv"
            if name in font_sizes:
                widget.text_fontsize = font_sizes[name]
            else:
                widget.text_fontsize = 9.5
            widget.update()
        elif name in cb_whitelists:
            widget.field_value = cb_whitelists[name]
            widget.update()

TEMP_FILLED_ACRO_PDF = os.path.join(SCRATCH_DIR, "temp_filled_acro_6p.pdf")
doc_acro.save(TEMP_FILLED_ACRO_PDF)
doc_acro.close()
print(f"Saved filled 6-page AcroForm PDF: {TEMP_FILLED_ACRO_PDF}")

# ==============================================================================
# PART 3: COMBINE 6-PAGE ACROFORM + 3-PAGE TABLES INTO 9-PAGE MASTER PDF
# ==============================================================================
print("\n=== [PART 3] ASSEMBLING 9-PAGE MASTER DELIVERABLE PDF ===")

writer = PdfWriter()

reader_form = PdfReader(TEMP_FILLED_ACRO_PDF)
assert len(reader_form.pages) == 6, f"Expected 6 pages form, got {len(reader_form.pages)}"
for p in reader_form.pages:
    writer.add_page(p)

reader_tables = PdfReader(TABLES_TEMP_PDF)
assert len(reader_tables.pages) == 3, f"Expected 3 pages tables, got {len(reader_tables.pages)}"
for p in reader_tables.pages:
    writer.add_page(p)

TEMP_COMBINED_9P = os.path.join(SCRATCH_DIR, "temp_combined_9p_pre_link.pdf")
with open(TEMP_COMBINED_9P, "wb") as f_out:
    writer.write(f_out)

print(f"Combined into pre-link 9-page PDF: {TEMP_COMBINED_9P}")

# Add active hyperlinks on Page 7 and Page 9
doc_final = fitz.open(TEMP_COMBINED_9P)
assert len(doc_final) == 9, f"Expected strictly 9 pages, got {len(doc_final)}"

p7 = doc_final[6] # Page 7 (Table 1)
p9 = doc_final[8] # Page 9 (Table 3)

rects_p7 = p7.search_for("ETX-261002-0288")
if rects_p7:
    print(f"Adding hyperlink on Page 7 for ETX-261002-0288 at rect: {rects_p7[0]}")
    p7.insert_link({
        'kind': fitz.LINK_URI,
        'from': rects_p7[0],
        'uri': "https://etrax.eagleanalytical.com/Submission/Details/Q57eK3N9WcR0Qh%24x6N5N6Q__"
    })

rects_p9 = p9.search_for("ETX-261021-0578")
if rects_p9:
    print(f"Adding hyperlink on Page 9 for ETX-261021-0578 at rect: {rects_p9[0]}")
    p9.insert_link({
        'kind': fitz.LINK_URI,
        'from': rects_p9[0],
        'uri': "https://etrax.eagleanalytical.com/SubmissionTest/Details/Cw%24QZhYAogGUwoutnOavpw__"
    })

# Save Master PDF to Desktop and Documents
doc_final.save(PDF_DELIVERABLE_DESKTOP)
doc_final.save(PDF_DELIVERABLE_DOCS)
doc_final.close()

print(f"\n==================================================================")
print(f"MASTER PDF BUILD COMPLETE AND SYNCHRONIZED!")
print(f"Desktop:   {PDF_DELIVERABLE_DESKTOP}")
print(f"Documents: {PDF_DELIVERABLE_DOCS}")
print(f"Pages:     9 PAGES STRICTLY")
print(f"==================================================================")
