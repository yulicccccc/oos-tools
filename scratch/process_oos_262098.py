import os, sys, json, re, io, copy, shutil
from datetime import datetime

OOS_ROOT = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
sys.path.insert(0, OOS_ROOT)
os.chdir(OOS_ROOT)

import docx
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
from docxtpl import DocxTemplate
from pypdf import PdfWriter, PdfReader
import win32com.client
import fitz

DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
OUTPUT_DIR = os.path.join(OOS_ROOT, "scratch")

print("--- STARTING USP <71> OOS-262098 FULL REPORT GENERATION (DUAL SAMPLE COVERAGE) ---")

# 1. Master Dataset covering BOTH samples (ETX-260902-0505 and ETX-260914-0470)
data = {
    "oos_id": "262098",
    "client_name": "Solyn LLC (E75000)",
    "sample_id": "ETX-260902-0505, ETX-260914-0470",
    "sample_id_pure": "ETX-260902-0505 and ETX-260914-0470",
    "sample_url_1": "https://etrax.eagleanalytical.com/Submission/Details/%247Ydn%24JOiYgmsnaKjbxn-g__",
    "sample_url_2": "https://etrax.eagleanalytical.com/Submission/Details/18rFOYVWqrO0Afk1OwB2hg__",
    "sample_name": "GLP3R/Cagrilinitide",
    "lot_number": "2608-216",
    "dosage_form": "Liquid",
    "test_name": "USP <71> / EP 2.6.1 Sterility (Every Day Read) Test",
    "test_method": "MICRO-SOP-5",
    "sop_effective_date": "24Jul2026",
    "sop_rev": "19",
    "method_suitability": "N/A",
    "method_performed": "Membrane Filtration",
    "amount_filtered": "3 vials (04Sep26) / 4 vials (15Sep26)",
    "process_date": "04Sep26, 15Sep26",
    "process_date_full": "04 Sep 2026 and 15 Sep 2026",
    "before_test": "03Sep26",
    "after_test": "08Sep26",
    "test_date": "10Sep26, 21Sep26",
    "test_date_full": "10 Sep 2026 and 21 Sep 2026",
    "received_data": "02Sep26, 14Sep26",
    "incubation_time": "14",
    "positive_media": "1 x 100mL TSB bottle (each)",
    "positive_id_1": "ETX-260910-0290",
    "positive_id_2": "ETX-260921-0498",
    "positive_id_url_1": "https://etrax.eagleanalytical.com/Submission/Details?id=oZa2TD1UI-0eAJOFysBJ3A__",
    "positive_id_url_2": "https://etrax.eagleanalytical.com/Submission/Details/7K%24JgQe6bduNBFY3pBa03w__",
    "positive_org": "Microbacterium sp. PM5",
    "prepper_initial": "ES, AC",
    "prepper_name": "Alex Saravia, Andrew Carrillo",
    "analyst_initial": "ES, AOD",
    "analyst_name": "Alex Saravia, Abayomi Odugbesi",
    "reading_initial": "EN, AC",
    "reading_name": "Elysse Nioupin, Andrew Carrillo",
    "writer_name": "Qiyue Chen",
    "qa_manager": "Robin Seymour",
    "qa_notified": "Kathan Parikh",
    "bsc_id": "1316, 1314",
    "cr_suit": "114, 115",
    "cr_id": "1736",
    "smart_cr_id": "CR 114 (E001736) and CR 115",
    "suit": "B",
    "bsc_location": "innermost ISO 7 room (114B) and ISO 7 room (Suite 115)",
    "monthly_cleaning_date": "30Aug26",
    "monthly_cleaning_date_full": "30 Aug 2026",
    "ftm_lot": "2567320",
    "ftm_exp": "12/16/2026",
    "tsb_lot": "2548010, 2567280",
    "tsb_exp": "04/19/2027, 05/10/2027",
    "fluid_d_lot": "685561, 687774",
    "fluid_d_exp": "10/31/2026, 02/28/2027",
}

# 2. Build Comprehensive Narrative Paragraphs covering BOTH samples
# Split across Page 3 (Setup & Incubation) and Page 4 (EM, Cross-Contamination, Non-uniform Distribution Justification)

# --- PAGE 3 (Text Field 49) ---
p1 = (
    "All analysts involved in the prepping, processing, and reading of both samples – Alex Saravia, Elysse Nioupin, "
    "Andrew Carrillo, and Abayomi Odugbesi – were interviewed comprehensively. Their answers are recorded throughout this document."
)

p2 = (
    "Upon arrival, both sample submissions (ETX-260902-0505 and ETX-260914-0470) were stored in accordance with the Client’s "
    "instructions. Analysts verified the integrity of the sample containers throughout both preparation and processing stages. "
    "No leaks, cracks, or turbidity were observed prior to testing, verifying container integrity."
)

p3 = (
    "All reagents and supplies mentioned in the material section above were stored according to suppliers’ recommendations, "
    "and their integrity was visually verified before utilization. Moreover, all culture media, rinses, and supplies possessed "
    "valid expiration dates. The functionality of all equipment was confirmed by reviewing continuous monitoring data."
)

p4 = (
    "Initial testing was conducted in Cleanroom Suite 114 (comprising the innermost ISO 7 cleanroom 114B, middle ISO 7 buffer room 114A, "
    "and outermost ISO 8 anteroom 114). Retest processing was conducted in Cleanroom Suite 115. Positive air pressure cascades are maintained "
    "throughout both suites to ensure controlled, outward unidirectional airflow."
)

p5 = (
    "Initial sample ETX-260902-0505 was received on 02 Sep 2026 and processed on 04 Sep 2026 by Alex Saravia in certified ISO 5 "
    "BSC E001316 (Cleanroom 114B) as per MICRO-SOP-5 (USP <71> / EP 2.6.1 Sterility Test). Prior to entry, vials were disinfected "
    "with acidified bleach allowing a 10-minute contact time at each transfer boundary. Membrane filtration was performed (3 vials "
    "filtered). Media bottles were transferred into designated incubators E001356 and E001357. On 10 Sep 2026 (Day 6 of incubation), "
    "turbidity was observed in 1 x 100mL TSB bottle by reading analyst Elysse Nioupin and confirmed by supervisor Robin Seymour. "
    "The positive bottle was submitted for Microbial Identification under ETX-260910-0290, which definitively identified the isolate "
    "as Microbacterium sp. PM5 (Gram (+) rods). Concurrent negative controls remained sterile."
)

p6 = (
    "Secondary sample ETX-260914-0470 was prepped on 14 Sep 2026 by Andrew Carrillo and processed on 15 Sep 2026 by Abayomi Odugbesi "
    "in certified ISO 5 BSC E001314 in Cleanroom Suite 115 as per MICRO-SOP-5 via membrane filtration (4 vials / 60 gm filtered). "
    "Standard multi-barrier disinfection and strict aseptic techniques were maintained. Media bottles were incubated in incubators "
    "E001356 and E001357. On 21 Sep 2026 (Day 6 of incubation), turbidity was observed in 1 x 100mL TSB bottle by reading analyst "
    "Andrew Carrillo and confirmed by supervisor Robin Seymour. The positive bottle was submitted for Microbial Identification under "
    "ETX-260921-0498, which definitively identified the identical isolate, Microbacterium sp. PM5 (Gram (+) short rods). Concurrent "
    "negative controls remained sterile."
)

text_field_49 = "\r \r".join([p1, p2, p3, p4, p5, p6])

# --- PAGE 4 (Text Field 50) ---
p10 = (
    "ISO 5 BSC Environmental Monitoring Results Evaluation:\r"
    "After reviewing the Environmental Monitoring (EM) records for both testing dates (04 Sep 2026 and 15 Sep 2026), no microbial growth "
    "was detected on operator personnel fingertip touch plates, settling plates, or ISO 5 BSC surface contact plates for either session, "
    "nor on preceding or subsequent testing days (including 03Sep26, 04Sep26, 08Sep26 for ES in BSC E001316, and 14Sep26, 15Sep26, 16Sep26 for AOD in BSC E001314)."
)

p11 = (
    "Cleanroom Environmental Monitoring Results Evaluation:\r"
    "For the initial processing on 04 Sep 2026 in Suite 114, weekly active surface sampling showed no growth; weekly active air sampling in ISO 8 cleanroom 114 "
    "recovered 2 CFUs (ETX-260914-0487), identified as Corynebacterium ureicelerivorans and Mycobacterium grossiae; on 10 Sep 2026, "
    "1 CFU was recovered (ETX-260921-0520), identified as Micrococcus luteus. These environmental isolates differed distinctly in genus and morphology "
    "from the sample isolate (Microbacterium sp. PM5). For the retest processing on 15 Sep 2026 in Cleanroom Suite 115, weekly active surface sampling showed no growth; "
    "weekly active air sampling in ISO 8 cleanroom 115 recovered 9 CFUs (ETX-260923-0402), identified as Kocuria indica, Brevibacterium sp. CS2, Corynebacterium sp, "
    "Micrococcus luteus, Paracoccus yeei, Staphylococcus hominis, and Kocuria rhizophila. 100% of these environmental isolates differed distinctly in genus and morphology "
    "from the product contamination isolate (Microbacterium sp. PM5). Both air recoveries were strictly confined to the outermost ISO 8 anterooms (ISO 8 114 and ISO 8 115). "
    "All actual testing occurred within certified ISO 5 BSCs (E001316 in Suite 114 and E001314 in Suite 115) located in ISO 7 cleanrooms separated by positive pressure "
    "cascades. Test containers were transferred in disinfected, lidded bins, and critical zone settling, surface contact, and analyst fingertip plates remained 100% sterile "
    "(0 CFU). Therefore, cleanroom air recoveries were unrelated to the sample contamination. (Refer to Table 1 for sample details, Table 2 for environmental monitoring "
    "from processing performed on 04 Sep 2026, and Table 3 for environmental monitoring from processing performed on 15 Sep 2026)."
)

p12 = (
    "The complete absence of contamination across all analyst glove touch plates, settling plates, and ISO 5 work surfaces across both processing "
    "sessions confirms that the ISO 5 critical zones remained in optimal control. Full compliance with MICRO-SOP-9 and MICRO-SOP-5 was verified. "
    "Monthly cleaning and disinfection of Cleanroom Suite 114 and Suite 115 was performed on 30 Aug 2026 as per MICRO-SOP-9, with all chemical H2O2 indicators passing."
)

p13 = (
    "A 6-month historical review for client Solyn LLC (E75000) confirmed that analyte GLP3R/Cagrilinitide had no prior sterility failures. "
    "Batch cross-contamination reviews were conducted for all samples processed on 04 Sep 2026 (where ETX-260902-0505 was 3rd in the batch) "
    "and 15 Sep 2026. All other client samples processed during both testing sessions tested negative, ruling out cross-contamination."
)

p14 = (
    "Importantly, the investigation evaluated why the prior rapid ScanRDI sterility test on this lot (ETX-260807-0602) yielded a passing result. "
    "In compounded parenteral products, low-level bioburden is characteristically non-uniformly distributed across individual vials. "
    "Under low-concentration non-uniform contamination, random unit sampling naturally results in some vials containing zero viable cells—as "
    "occurred with the vials sampled for ScanRDI under ETX-260807-0602—while other vials harbor low-level viable bioburden. Furthermore, no "
    "method suitability was on file for ScanRDI with this formulation. When low-level viable organisms were introduced into the 14-day USP <71> "
    "membrane filtration enrichment process, the microorganisms were successfully enriched and recovered in TSB on Day 6 in both independent "
    "testing events (ETX-260902-0505 and ETX-260914-0470), yielding the identical strain (Microbacterium sp. PM5). The successful replication "
    "across different processing analysts (Alex Saravia vs. Abayomi Odugbesi), different BSCs (E001316 vs. E001314), and different cleanroom suites "
    "(Suite 114 vs. Suite 115) conclusively refutes the preliminary hypothesis of analyst handling error during reconstitution. The passing ScanRDI result and both failing USP <71> results "
    "are scientifically consistent with non-uniform microbial distribution. Consequently, inherent product contamination in Lot 2608-216 is "
    "confirmed, and Lot 2608-216 fails USP <71> sterility requirements."
)

text_field_50 = "\r \r".join([p10, p11, p12, p13, p14])

# --- PAGE 5 (Text Field 51) ---
writer_initial = "QYC"
text_field_51 = f"N/A {writer_initial} 08Oct26"

# Assemble Master Narrative for Word Template
smart_phase1_full = f"{text_field_49}\n\n{text_field_50}"

data["smart_phase1_summary"] = smart_phase1_full
data["equipment_summary"] = p5
data["narrative_summary"] = p11
data["sample_history_paragraph"] = p13
data["cross_contamination_summary"] = p13
data["smart_comment_interview"] = "Yes. Analysts Alex Saravia, Elysse Nioupin, Andrew Carrillo, and Abayomi Odugbesi were interviewed comprehensively. Handling error was refuted by identical retest recovery."
data["smart_comment_samples"] = f"Yes, sample IDs: {data['sample_id_pure']}"
data["smart_comment_records"] = f"Yes, information is available on EagleTrax under {data['sample_id_pure']}"
data["smart_comment_storage"] = f"Yes, both samples were stored as per client's instructions. Information is available in EagleTrax Sample Location History under {data['sample_id_pure']}"
data["smart_personnel_block"] = (
    "Prepping Analyst:\r"
    "Alex Saravia (ES), Andrew Carrillo (AC)\r \r"
    "Processing Analyst:\r"
    "Alex Saravia (ES), Abayomi Odugbesi (AOD)\r \r"
    "Reading Analyst:\r"
    "Elysse Nioupin (EN), Andrew Carrillo (AC)"
)
data["analyst_signature"] = f"Alex Saravia (Written by: {data['writer_name']})"
data["smart_incident_opening"] = (
    f"On 10 Sep 2026, sample ETX-260902-0505 was confirmed positive and on 21 Sep 2026, sample ETX-260914-0470 was confirmed "
    f"positive for viable microorganisms (1 x 100mL TSB each) after USP <71> sterility testing. Both samples belong to lot {data['lot_number']} "
    f"under investigation in OOS-{data['oos_id']}."
)
data["usp71_id"] = "E001316, E001314"
data["reader_name"] = "Elysse Nioupin, Andrew Carrillo"

# 3. Master Word Report Generation
out_doc_report = os.path.join(OUTPUT_DIR, f"OOS-{data['oos_id']} {data['client_name']} - USP71.docx")
primary_tpl = "USP71 OOS P1 template.docx" if os.path.exists("USP71 OOS P1 template.docx") else "USP71 OOS P1 template 0.docx"
tpl_report = DocxTemplate(primary_tpl)
tpl_report.render(data)

# Read the standalone table document to insert clean tables into Master Word report
standalone_docx = os.path.join(OUTPUT_DIR, f"Tables OOS-{data['oos_id']} {data['client_name']} - USP71.docx")
if not os.path.exists(standalone_docx):
    standalone_docx = os.path.join(DESKTOP_DIR, f"Tables OOS-{data['oos_id']} {data['client_name']} - USP71.docx")

if os.path.exists(standalone_docx):
    doc_tables_src = docx.Document(standalone_docx)
    if len(tpl_report.docx.tables) >= 3 and len(doc_tables_src.tables) >= 2:
        t2_elem = tpl_report.docx.tables[1]._element
        p2_parent = t2_elem.getparent()
        idx2 = p2_parent.index(t2_elem)
        p2_parent.remove(t2_elem)
        new_t1_elem = copy.deepcopy(doc_tables_src.tables[0]._element)
        p2_parent.insert(idx2, new_t1_elem)

        t3_elem = tpl_report.docx.tables[2]._element
        p3_parent = t3_elem.getparent()
        idx3 = p3_parent.index(t3_elem)
        p3_parent.remove(t3_elem)
        new_t2_elem = copy.deepcopy(doc_tables_src.tables[1]._element)
        p3_parent.insert(idx3, new_t2_elem)

        if len(doc_tables_src.tables) >= 3:
            p_t3_heading = parse_xml(
                '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:pPr><w:spacing w:before="240" w:after="160"/>'
                '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:u w:val="single"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:u w:val="single"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
                '<w:t>Table 3: Environmental Monitoring from Processing Performed on 15Sep26</w:t></w:r></w:p>'
            )
            new_t3_elem = copy.deepcopy(doc_tables_src.tables[2]._element)
            p3_parent.insert(idx3 + 1, p_t3_heading)
            p3_parent.insert(idx3 + 2, new_t3_elem)

for p in tpl_report.docx.paragraphs:
    if "Table 1:" in p.text:
        p.text = f"Table 1: Information for {data['sample_id']} under investigation"
        p.runs[0].bold = True
    elif "Table 2:" in p.text:
        p.text = "Table 2: Environmental Monitoring from Processing Performed on 04Sep26"
        p.runs[0].bold = True

tpl_report.save(out_doc_report)
print("Saved Master Word Report to:", out_doc_report)

# 4. Official PDF Form (6 Pages base + Page 7 Standalone Table)
base_pdf = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 - P1 04 SEP 2026.pdf"
if not os.path.exists(base_pdf):
    base_pdf = "USP71 OOS P1 template.pdf"
print("Using base PDF form:", base_pdf)

smart_consumables_text = (
    "Process Consumables:\r\n"
    "ETX-260902-0505: FTM Lot 2567320, TSB Lot 2548010, Fluid D Lot 685561\r\n"
    "ETX-260914-0470: FTM Lot 2567320, TSB Lot 2567280, Fluid D Lot 687774, Bacteriostatic Water NM7681\r\n \r\n"
    "Environmental Plates:\r\n"
    "TSA and Surface Plate: see attached environmental logs\r\n \r\n"
    "Monthly Cleaning:\r\n"
    "H2O2 strips and IPA: See attached monthly cleaning logs"
)

pdf_map = {
    'Text Field57': data['oos_id'],
    'Date Field0': "04-Sep-2026",
    'Date Field1': "10-Sep-2026",
    'Date Field2': "10-Sep-2026",
    'Date Field3': "10-Sep-2026",
    'Text Field1': data['test_name'],
    'Text Field2': data['sample_id'],
    'Text Field3': data['smart_personnel_block'],
    'Text Field4': data['sample_name'],
    'Text Field5': data['dosage_form'],
    'Text Field6': data['lot_number'],
    'Text Field7': data['smart_incident_opening'],
    'Text Field8': data['test_method'],
    'Text Field9': data['sop_effective_date'],
    'Text Field10': data['sop_rev'],
    'Text Field11': "Pass/Fail",
    'Text Field12': data['qa_notified'],
    'Text Field0': data['analyst_signature'],
    'Text Field13': data['smart_comment_interview'],
    'Text Field14': data['smart_comment_samples'],
    'Text Field15': f"Yes, as per {data['test_method']}",
    'Text Field16': f"Yes, as per {data['test_method']}",
    'Text Field17': data['smart_comment_records'],
    'Text Field18': "Yes, the analysts are trained and qualified by quality to perform the test",
    'Text Field19': "Not Applicable",
    'Text Field20': "Not Applicable",
    'Text Field21': data['smart_comment_storage'],
    'Text Field22': smart_consumables_text,
    'Text Field23': smart_consumables_text,
    'Text Field24': "Not applicable",
    'Text Field25': "Not applicable",
    'Text Field26': "Not applicable",
    'Text Field27': "Not applicable",
    'Text Field28': "Not applicable",
    'Text Field29': "Not applicable",
    'Text Field30': "ISO 5 BSC E001316, ISO 5 BSC E001314",
    'Text Field31': "Jun 2027, Jun 2027",
    'Text Field32': data['smart_cr_id'],
    'Text Field33': "Dec 2026",
    'Text Field34': "Not applicable",
    'Text Field35': "Not applicable",
    'Text Field36': "Not applicable",
    'Text Field37': "Not applicable",
    'Text Field38': "Not applicable",
    'Text Field39': "Not applicable",
    'Text Field40': "See Phase I Summary",
    'Text Field41': "See Phase I Summary",
    'Text Field42': "See Phase I Summary",
    'Text Field43': (
        "Incubator E001356 (Sensor E001450)\r \r"
        "Incubator E001357 (Sensor E001449)\r \r"
        "Incubator E001034 (Sensor E001501)\r \r"
        "Incubator E001031 (Sensor E001505)"
    ),
    'Text Field44': (
        "Jan 2027 / Feb 2027\r \r"
        "Jan 2027 / Feb 2027\r \r"
        "Aug 2027 / Feb 2027\r \r"
        "Aug 2027 / Feb 2027"
    ),
    'Text Field45': "See Phase I Summary",
    'Text Field46': "Not applicable",
    'Text Field47': "Not applicable",
    'Text Field48': f"N/A {writer_initial} 08Oct26",
    'Text Field49': text_field_49,
    'Text Field50': text_field_50,
    'Text Field51': text_field_51,
    'Text Field52': (
        f"Inherent product contamination in Lot {data['lot_number']}, confirmed by independent dual-sample recovery "
        f"of identical strain Microbacterium sp. PM5 on Day 6 in TSB under ETX-260902-0505 and ETX-260914-0470."
    ),
    'Text Field53': data['writer_name'],
    'Text Field54': ""
}

# Checkboxes matching approved gold-standard
checkbox_fields = {
    'Check Box0': '/Yes',
    'Check Box1': '/Yes',
    'Check Box2': '/Yes',
    'Check Box4': '/Yes',
    'Check Box7': '/Yes',
    'Check Box10': '/Yes',
    'Check Box13': '/Yes',
    'Check Box16': '/Yes',
    'Check Box19': '/Yes',
    'Check Box24': '/Yes',
    'Check Box27': '/Yes',
    'Check Box28': '/Yes',
    'Check Box32': '/Yes',
    'Check Box34': '/Yes',
    'Check Box38': '/Yes',
    'Check Box42': '/Yes',
    'Check Box43': '/Yes',
    'Check Box48': '/Yes',
    'Check Box51': '/Yes',
    'Check Box52': '/Yes',
    'Check Box55': '/Yes',
    'Check Box60': '/Yes',
    'Check Box63': '/Yes',
    'Check Box66': '/Yes',
    'Check Box67': '/Yes',  # Environmental monitoring conforms
    'Check Box70': '/Yes',  # Controls conform
    'Check Box73': '/Yes',
    'Check Box78': '/Yes',
    'Check Box79': '/Yes',
    'Check Box87': '/Yes',  # Lab error identified (preliminary)
    'Check Box89': '/Yes',  # Initiate Phase II Form 3.100.019.F02
    'Check Box91': '/Yes',  # Cannot close - initiate Phase II
}
pdf_map.update(checkbox_fields)

writer = PdfWriter(clone_from=base_pdf)
while len(writer.pages) > 6:
    del writer.pages[-1]

for p in writer.pages:
    writer.update_page_form_field_values(p, pdf_map)

# Append Standalone Table PDF as Page 7
standalone_pdf = os.path.join(OUTPUT_DIR, f"Tables OOS-{data['oos_id']} {data['client_name']} - USP71.pdf")
if not os.path.exists(standalone_pdf):
    standalone_pdf = os.path.join(DESKTOP_DIR, f"Tables OOS-{data['oos_id']} {data['client_name']} - USP71.pdf")

if os.path.exists(standalone_pdf):
    table_reader = PdfReader(standalone_pdf)
    for table_page in table_reader.pages:
        writer.add_page(table_page)
    print(f"Appended Tables PDF to form. Total pages: {len(writer.pages)} (Must be 8)")

out_pdf_report = os.path.join(OUTPUT_DIR, f"OOS-{data['oos_id']} {data['client_name']} - USP71.pdf")
with open(out_pdf_report, "wb") as f:
    writer.write(f)
print("Saved complete PDF Report to:", out_pdf_report)

# Calibrate font sizes in PDF fields per Rule 11 (Natural Human Typography & Acrobat '+' Elimination)
try:
    doc_fitz = fitz.open(out_pdf_report)
    font_sizes = {
        # Page 1
        'Text Field0': 8.5,
        'Date Field0': 8.5,
        'Date Field1': 8.5,
        'Date Field2': 8.5,
        'Date Field3': 8.5,
        'Text Field1': 8.5,
        'Text Field2': 8.0,
        'Text Field3': 5.8,
        'Text Field4': 8.5,
        'Text Field5': 8.5,
        'Text Field6': 8.5,
        'Text Field7': 6.2,
        'Text Field8': 8.5,
        'Text Field9': 8.5,
        'Text Field10': 8.5,
        'Text Field11': 8.5,
        'Text Field12': 8.5,
        'Text Field13': 4.825,
        'Text Field14': 6.0,
        'Text Field15': 9.0,
        'Text Field16': 9.0,
        'Text Field17': 6.5,
        'Text Field18': 4.825,
        'Text Field19': 9.0,
        'Text Field20': 9.0,
        'Text Field21': 4.8,
        
        # Page 2
        'Text Field22': 4.2,
        'Text Field23': 4.2,
        'Text Field24': 8.5,
        'Text Field25': 8.5,
        'Text Field26': 8.5,
        'Text Field27': 8.5,
        'Text Field28': 8.5,
        'Text Field29': 8.5,
        'Text Field30': 7.0,
        'Text Field31': 7.0,
        'Text Field32': 4.825,
        'Text Field33': 4.825,
        'Text Field34': 8.5,
        'Text Field35': 8.5,
        'Text Field36': 8.5,
        'Text Field37': 8.5,
        'Text Field38': 8.5,
        'Text Field39': 8.5,
        'Text Field40': 8.5,
        'Text Field41': 8.5,
        'Text Field42': 8.5,
        'Text Field43': 7.05,
        'Text Field44': 7.05,
        'Text Field45': 8.5,
        
        # Page 3
        'Text Field46': 8.5,
        'Text Field47': 8.5,
        'Text Field48': 8.5,
        'Text Field49': 8.0,
        
        # Page 4
        'Text Field50': 8.0,
        
        # Page 5
        'Text Field51': 7.8,
        
        # Page 6
        'Text Field52': 7.5,
        'Text Field53': 10.0,
        'Text Field57': 6.5,
    }
    for page in doc_fitz:
        for w in page.widgets():
            if w.field_name in font_sizes:
                w.text_fontsize = font_sizes[w.field_name]
                w.update()
    temp_pdf_out = out_pdf_report + ".tmp.pdf"
    doc_fitz.save(temp_pdf_out)
    doc_fitz.close()
    shutil.move(temp_pdf_out, out_pdf_report)
    print("Fine-tuned all field font sizes with PyMuPDF.")
except Exception as e:
    print(f"PyMuPDF font tuning notice: {e}")

# Save State Data (.txt / .json)
out_json_path = os.path.join(OUTPUT_DIR, f"SAVE_OOS-{data['oos_id']} {data['client_name']} - USP71.txt")
with open(out_json_path, "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2)
print("Saved session state JSON to:", out_json_path)

# Copy QYC named PDF
qyc_pdf_report = os.path.join(OUTPUT_DIR, f"OOS-{data['oos_id']} {data['client_name']} - USP71 - QYC.pdf")
shutil.copy2(out_pdf_report, qyc_pdf_report)

# Sync all files to Desktop and Documents safely
files_to_sync = [
    out_doc_report,
    out_pdf_report,
    qyc_pdf_report,
    out_json_path
]
for fpath in files_to_sync:
    if fpath and os.path.exists(fpath):
        for target_dir in [DOCUMENTS_DIR, DESKTOP_DIR]:
            dest = os.path.join(target_dir, os.path.basename(fpath))
            try:
                shutil.copy2(fpath, dest)
                print(f"Synced to: {dest}")
            except Exception as e:
                print(f"Notice: {dest} locked or skipped ({e})")
                if target_dir == DESKTOP_DIR and fpath.endswith(".pdf"):
                    revised_dest = os.path.join(target_dir, os.path.splitext(os.path.basename(fpath))[0] + " - Revised.pdf")
                    try:
                        shutil.copy2(fpath, revised_dest)
                        print(f"Saved Revised PDF fallback to: {revised_dest}")
                    except Exception:
                        pass
                elif target_dir == DESKTOP_DIR and fpath.endswith(".docx"):
                    revised_dest = os.path.join(target_dir, os.path.splitext(os.path.basename(fpath))[0] + " - Revised.docx")
                    try:
                        shutil.copy2(fpath, revised_dest)
                        print(f"Saved Revised DOCX fallback to: {revised_dest}")
                    except Exception:
                        pass

print("\n--- ALL MASTER OOS-262098 DELIVERABLES GENERATED SUCCESSFULLY! ---")
