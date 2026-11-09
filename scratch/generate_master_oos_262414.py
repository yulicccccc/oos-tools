"""
Script: generate_master_oos_262414.py
Purpose: Complete generation of Master Word Document and 7-Page AcroForm PDF
         for Scan RDI OOS-262414 (Reshape Rx Houston (E74105), Tesamorelin 20mg/2ml).
         Adheres strictly to RULE 00 and Rule 92 (Universal 6-Step Golden Template 1:1 Cloning SOP).
"""

import os
import sys
import copy
import shutil
import fitz  # PyMuPDF
import docx
from docx.shared import Pt, Inches
from docxtpl import DocxTemplate
from pypdf import PdfWriter, PdfReader
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"

OOS_ID = "262414"
CLIENT_NAME = "Reshape Rx Houston (E74105)"
SAMPLE_ID = "ETX-261021-0286"
SAMPLE_URL = "https://etrax.eagleanalytical.com/Submission/Details/LMMxVFbm9M7KsDBu5-mJXg__"
SAMPLE_NAME = "Tesamorelin 20mg/2ml"
LOT_NUMBER = "QC264093"
DOSAGE_FORM = "Injectable"
TEST_DATE = "22Oct26"
TEST_DATE_FULL = "22-Oct-2026"
TEST_RECORD = "102226-2132-2"
ANALYSIS_ID = "102226-2132-2021"

PROCESSOR_NAME = "Karla Silva"
PROCESSOR_INIT = "KSM"
CHANGEOVER_NAME = "Sonal Uprety"
CHANGEOVER_INIT = "SU"
READER_NAME = "Sonal Uprety"
READER_INIT = "SU"

BSC_PROC = "1938"
BSC_CHG = "1937"
SCAN_ID = "2132"
MONTHLY_CLEANING_DATE = "27Sep26"

EVENTS_COUNT = "145"
CONFIRMED_COUNT = "84"
MORPHOLOGY = "rod-shaped morphology"

# Positive Control from run 102226-2132-2 (Control + KSM, Analysis ID 102226-2132-2016)
CONTROL_POSITIVE = "A. brasiliensis"
CONTROL_LOT = "05222026-16404-AB"
CONTROL_EXP = "22May28"

# -------------------------------------------------------------
# 1. Narrative Text Blocks (Aligned 1:1 with Golden Template OOS-262289)
# -------------------------------------------------------------
p1 = (
    f"All analysts involved in the preparation, processing, changeover, and reading of the sample – "
    f"{PROCESSOR_NAME} and {READER_NAME} – were interviewed comprehensively. "
    "Their responses are documented throughout this report. This investigation was performed in "
    "accordance with MICRO-SOP-53, Sterility Test Out-of-Specification Investigation Procedure."
)

p2 = (
    "Upon arrival, the sample was stored refrigerated in accordance with the Client's instructions. "
    f"Processing analyst - {PROCESSOR_NAME} - verified the physical integrity and packaging of the sample "
    "throughout both the preparation and processing stages. No leaks, cracks, or visible turbidity were "
    "observed prior to testing, verifying container integrity."
)

p3 = (
    "All reagents and supplies mentioned in the material section above were stored according to suppliers’ "
    "recommendations, and their integrity was visually verified before utilization. Moreover, each reagent, "
    "media, and consumable possessed valid expiration dates. The functionality of all equipment was confirmed "
    "by reviewing continuous environmental and operational monitoring data."
)

p4 = (
    f"During the preparation phase, {PROCESSOR_NAME} disinfected the sample containers using acidified bleach "
    "and placed them into a pre-disinfected storage bin. On 22Oct26, prior to sample processing, "
    f"{PROCESSOR_NAME} performed a second disinfection with acidified bleach, allowing a minimum contact "
    "time of 10 minutes before transferring the samples into the cleanroom suites. A final disinfection step "
    "was completed immediately before introducing the sample into the ISO 5 Biological Safety Cabinet "
    f"(BSC E00{BSC_PROC}) located in the ISO 7 cleanroom (145). All activities were performed in "
    "accordance with MICRO-SOP-12, Rapid Scan RDI® Test using FIFU Method."
)

p5 = (
    "Cleanroom L-Suite, used for both sample processing and changeover, consists of ISO 7 cleanroom 145, "
    "ISO 7 cleanroom 144, ISO 8 anteroom 143, and outermost ISO 8 room 142. Outward-cascading positive "
    "air pressure differentials are continuously maintained across the suite to ensure controlled "
    "unidirectional airflow."
)

p6 = (
    f"The ISO 5 BSC E00{BSC_PROC} and ISO 5 BSC E00{BSC_CHG} located in Cleanroom 145 and Cleanroom 144 "
    "(L-Suite) were thoroughly cleaned and disinfected prior to their respective procedures in accordance with "
    "MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. Both BSCs were certified and approved "
    "by the Engineering and Quality Assurance teams prior to testing."
)

p7 = (
    f"Pre-labeling and filtration activities were performed by {PROCESSOR_NAME} on 22Oct26 within the ISO 5 "
    f"biosafety cabinet (BSC E00{BSC_PROC}) located in the ISO 7 cleanroom (145). The changeover procedure "
    f"was subsequently performed by analyst {CHANGEOVER_NAME} on the same date within the ISO 5 biosafety cabinet "
    f"(BSC E00{BSC_CHG}) located in Cleanroom 144 of L-Suite. The reading analyst, {READER_NAME}, confirmed that the "
    f"ScanRDI instrument E00{SCAN_ID} was set up as per ENG-SOP-4 (Scan RDI® System – Operations (Standard C3 Quality "
    f"Check and Microscope Setup and Maintenance)), and the negative control and positive control for analyst {PROCESSOR_NAME} "
    "yielded expected results."
)

p8 = (
    f"On 22Oct26, a rapid sterility test was performed on the sample using the ScanRDI method under test record {TEST_RECORD} "
    f"(Analysis ID {ANALYSIS_ID}). The sample was processed by {PROCESSOR_NAME} (12th sample processed, ksm12), "
    f"changeover performed by {CHANGEOVER_NAME}, and subsequently read by {READER_NAME}. The test revealed "
    f"{EVENTS_COUNT} total events and {CONFIRMED_COUNT} confirmed microbial events exhibiting {MORPHOLOGY}, see Table 1."
)

p9 = (
    f"Table 2 presents the environmental monitoring results associated with testing performed on {TEST_DATE}. "
    "The environmental monitoring (EM) plates were incubated for no less than 48 hours at 30–35°C and for no less than "
    "an additional five days at 20–25°C, as per MICRO-SOP-2, Environmental Monitoring of the Clean-room Facility."
)

p10 = (
    "Upon review of the environmental monitoring data, absence of microbial recovery (No Growth) was demonstrated "
    "across all critical ISO 5 zones and personnel monitoring on the date of testing. Daily personnel monitoring for "
    f"processing analyst {PROCESSOR_NAME} and changeover analyst {CHANGEOVER_NAME} yielded no microbial growth on left "
    f"and right fingertip touch plates. Daily surface sampling (4 locations) and settling plates within processing BSC E00{BSC_PROC} "
    f"and changeover BSC E00{BSC_CHG} yielded no microbial growth."
)

p11 = (
    "Weekly surface monitoring of Cleanroom L-Suite conducted on 20Oct26 demonstrated no microbial recovery across all "
    "sampled cleanroom surfaces, confirming effective sanitization and surface control."
)

p12 = (
    "Weekly active air monitoring of Cleanroom L-Suite conducted on 20Oct26 demonstrated no microbial recovery in ISO 7 "
    "cleanroom 145; however, 2 CFUs under test sample ETX-261027-0484 were recovered from ISO 7 cleanroom 144 Section I, "
    "21 CFUs under test sample ETX-261027-0487 were recovered from ISO 8 room 143 Section I, 17 CFUs under test sample "
    "ETX-261027-0492 were recovered from ISO 8 room 143 Section II, and 16 CFUs under test sample ETX-261027-0494 were "
    "recovered from outermost ISO 8 room 142. Microbial identification testing for all recovered environmental isolates "
    "remains pending as of 09Nov2026 and continues to be monitored under extended review."
)

p13 = (
    f"It is important to emphasize that all sample processing activities were performed strictly within the validated ISO 5 "
    f"BSC E00{BSC_PROC} (located in ISO 7 cleanroom 145) and changeover activities within ISO 5 BSC E00{BSC_CHG} (located in "
    f"ISO 7 cleanroom 144). There is minimal exposure of the test sample to ambient ISO 8 air, as samples and supplies are "
    "transferred in disinfected, closed containers on carts through the layered cleanroom suites. Furthermore, the active air "
    "recoveries occurred in the general cleanroom and lower-classified outer room environments, which are physically segregated "
    "from the ISO 5 processing zone by closed doors and an outward-cascading positive air pressure gradient. In conjunction with "
    "the complete absence of microbial growth on all analyst glove touch plates, settling plates, and ISO 5 BSC work surfaces "
    "throughout testing, there is no plausible contamination transfer pathway from the ambient environment to the test sample."
)

p14 = (
    f"Monthly cleaning and disinfection of the cleanrooms (ISO 7) and their containing biosafety cabinets (ISO 5) were performed "
    f"on {MONTHLY_CLEANING_DATE} for L-Suite, as per MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. It was documented "
    "that all chemical indicators passed, which confirmed the efficient monthly cleaning of all cleanrooms in L-suite."
)

p15 = (
    f"Analyzing a 6-month sample history for {CLIENT_NAME}, this specific analyte \"{SAMPLE_NAME}\" has had no prior Out-of-Specification (OOS) occurrences under ScanRDI sterility testing."
)

p16 = (
    "All samples processed within the same batch and across the facility on the date of testing were also evaluated "
    "for evidence of sample-to-sample cross-contamination or shared laboratory vectors. The test sample "
    f"{SAMPLE_ID} was processed as the 12th sample (ksm12) by {PROCESSOR_NAME} on 22Oct26 within BSC E00{BSC_PROC} "
    "(Cleanroom 145, Parallel Suite). The sample was released for testing with \"No Method Suitability Test on File.\" "
    f"All other samples processed within the same testing batch by {PROCESSOR_NAME} yielded negative results for microbial growth. "
    "One other sterility failure occurred on the same testing date (ETX-261019-0832, documented under OOS-262410 for "
    "BelieveRX Pharmacy (E74120), Lot 101626-7B). However, that sample was processed by a completely different analyst "
    "(Halaina Smith, 5th sample processed) within a physically separate cleanroom (ISO 7 Cleanroom 144 of L-Suite) and "
    "BSC E001937, with changeover and reading performed by Maryam Naeem. Furthermore, the two samples exhibited different "
    "cellular morphologies: cocci and diplococci shaped morphology in ETX-261019-0832 under OOS-262410 versus rod-shaped "
    f"morphology in {SAMPLE_ID} under OOS-262414. Analyst {PROCESSOR_NAME} verified that gloves were sanitized between samples "
    "and standard aseptic techniques were strictly maintained throughout testing. Consequently, there is no evidence of "
    "sample-to-sample cross-contamination or a shared laboratory vector."
)

p17 = (
    "Based on the laboratory investigation, no assignable laboratory cause was identified involving the analyst, instrument, "
    "reagents, supplies, or monitored laboratory environment. Therefore, the likelihood of laboratory error contributing to this "
    "failure is considered minimal, and the original result is deemed valid."
)

# Text Field partitions for PDF
text_field_49 = "\r \r".join([p1, p2, p3, p4, p5, p6, p7, p8, p9])

p11_header = "Environmental Monitoring (EM) Recoveries Review:"
p11_sub1 = f"ISO 5 BSC and Personnel Evaluation:\r{p10}"
p12_sub2 = f"Cleanroom EM Evaluation:\r{p11}"

text_field_50 = "\r \r".join([p11_header, p11_sub1, p12_sub2, p12, p13, p14, p15, p16, p17])
text_field_51 = "N/A QYC 22Oct26"

tf43_equipment = (
    "Incubator E001034\r \r"
    "(Sensor E001501)\r \r"
    "Incubator E001031\r \r"
    "(Sensor E001505)\r \r"
    f"BSC E00{BSC_PROC}\r \r"
    f"BSC E00{BSC_CHG}"
)
tf44_cal_due = (
    "Aug 2027\r \r"
    "Feb 2027\r \r"
    "Aug 2027\r \r"
    "Feb 2027\r \r"
    "Jul 2027\r \r"
    "Jul 2027"
)

tf22_materials = (
    "Scan Consumables:\rSee the attached data packet\r \r"
    "Environmental Plates:\rTSA and Surface Plate: see attached environmental logs\r \r"
    "Monthly Cleaning:\rH2O2 strips and IPA: See attached monthly cleaning logs"
)

# -------------------------------------------------------------
# 2. Render Master Word Document (Internal Scratch Only)
# -------------------------------------------------------------
def generate_master_word_doc():
    template_path = os.path.join(ROOT_DIR, "ScanRDI OOS P1 template.docx")
    if not os.path.exists(template_path):
        template_path = os.path.join(ROOT_DIR, "ScanRDI OOS template.docx")
    
    tpl = DocxTemplate(template_path)
    
    personnel_block = (
        f"Processing Analyst:\n{PROCESSOR_NAME} ({PROCESSOR_INIT})\n\n"
        f"Changeover Analyst:\n{CHANGEOVER_NAME} ({CHANGEOVER_INIT})\n\n"
        f"Reading Analyst:\n{READER_NAME} ({READER_INIT})"
    )
    
    context = {
        "oos_id": OOS_ID,
        "client_name": CLIENT_NAME,
        "sample_id": SAMPLE_ID,
        "sample_url": SAMPLE_URL,
        "sample_name": SAMPLE_NAME,
        "dosage_form": DOSAGE_FORM,
        "lot_number": LOT_NUMBER,
        "test_date": TEST_DATE,
        "test_date_full": TEST_DATE_FULL,
        "test_record": TEST_RECORD,
        "processor_name": PROCESSOR_NAME,
        "changeover_name": CHANGEOVER_NAME,
        "reader_name": READER_NAME,
        "bsc_processing": BSC_PROC,
        "bsc_changeover": BSC_CHG,
        "smart_cr_id": "E001979 (CR145)\rE001978 (CR144)",
        "smart_bsc_bracketing_header": f"Biological Safety Cabinet EM for E00{BSC_PROC} and E00{BSC_CHG}",
        "control_positive": CONTROL_POSITIVE,
        "control_lot": CONTROL_LOT,
        "control_data": CONTROL_EXP,
        "event_number": EVENTS_COUNT,
        "confirm_number": CONFIRMED_COUNT,
        "organism_morphology": MORPHOLOGY,
        "smart_personnel_block": personnel_block,
        "smart_incident_opening": f"On {TEST_DATE}, sample {SAMPLE_ID} was found positive for viable microorganisms after Scan RDI sterility testing.",
        "smart_comment_interview": f"Yes, analysts {PROCESSOR_NAME} and {READER_NAME} were interviewed comprehensively.",
        "smart_comment_samples": f"Yes, {SAMPLE_ID}",
        "smart_comment_records": f"Yes, See {TEST_RECORD} for more information.",
        "smart_comment_storage": f"Yes, Information is available in Eagle Trax Sample Location History under {SAMPLE_ID}",
        "smart_phase1_summary": "\n\n".join([p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17]),
        "smart_phase1_continued": "\n\n".join([p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17]),
    }
    
    tpl.render(context)
    
    # Replace Table 1 & Table 2 with verified standalone tables from scratch
    tables_source_docx = os.path.join(SCRIPT_DIR, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.docx")
    if os.path.exists(tables_source_docx):
        src_doc = docx.Document(tables_source_docx)
        if len(tpl.docx.tables) >= 3 and len(src_doc.tables) >= 2:
            # Replace Table 1
            t1_elem = tpl.docx.tables[1]._element
            t1_parent = t1_elem.getparent()
            idx1 = t1_parent.index(t1_elem)
            t1_parent.remove(t1_elem)
            t1_new = copy.deepcopy(src_doc.tables[0]._element)
            t1_parent.insert(idx1, t1_new)
            
            # Replace Table 2
            t2_elem = tpl.docx.tables[2]._element
            t2_parent = t2_elem.getparent()
            idx2 = t2_parent.index(t2_elem)
            t2_parent.remove(t2_elem)
            t2_new = copy.deepcopy(src_doc.tables[1]._element)
            t2_parent.insert(idx2, t2_new)
            
            # Remove trailing table 3 if present
            if len(tpl.docx.tables) >= 4:
                t3_elem = tpl.docx.tables[3]._element
                t3_elem.getparent().remove(t3_elem)
                
    out_docx_scratch = os.path.join(SCRIPT_DIR, f"OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.docx")
    tpl.save(out_docx_scratch)
    print("Master Word Docx saved to scratch:", out_docx_scratch)
    return out_docx_scratch

# -------------------------------------------------------------
# 3. Generate Complete 7-Page PDF Report (1:1 Golden Template Clone)
# -------------------------------------------------------------
def generate_master_pdf_report():
    golden_pdf = os.path.join(ROOT_DIR, "templates", "golden", "ScanRDI_Golden_OOS-262289.pdf")
    doc = fitz.open(golden_pdf)
    
    personnel_block = (
        f"Processing Analyst:\n{PROCESSOR_NAME} ({PROCESSOR_INIT})\n\n"
        f"Changeover Analyst:\n{CHANGEOVER_NAME} ({CHANGEOVER_INIT})\n\n"
        f"Reading Analyst:\n{READER_NAME} ({READER_INIT})"
    )
    
    pdf_field_values = {
        'Text Field57': OOS_ID,
        'Date Field0': TEST_DATE_FULL,
        'Date Field1': TEST_DATE_FULL,
        'Date Field2': TEST_DATE_FULL,
        'Date Field3': "23-Oct-2026",
        'Text Field2': SAMPLE_ID,
        'Text Field6': LOT_NUMBER,
        'Text Field4': "Tesamorelin\r20mg/2ml",
        'Text Field5': DOSAGE_FORM,
        'Text Field8': "MICRO-SOP-12 (17)\rENG-SOP-4 (05)",
        'Text Field9': "28Sep26\r24Jul26",
        'Text Field10': "Rev: 17\rRev: 05",
        'Text Field15': "Yes, as per MICRO-SOP-12, ENG-SOP-4",
        'Text Field16': "Yes, as per MICRO-SOP-12, ENG-SOP-4",
        'Text Field0': f"{READER_NAME} (Written by: Qiyue Chen)",  # Initiator = Reader
        'Text Field1': "Scan RDI Sterility Test",
        'Text Field3': personnel_block,
        'Text Field7': f"On {TEST_DATE}, sample {SAMPLE_ID} was found positive for viable microorganisms after Scan RDI sterility testing.",
        'Text Field11': "<1 event",
        'Text Field12': "Kathan Parikh",
        'Text Field13': f"Yes, analysts {PROCESSOR_NAME} and {READER_NAME} were interviewed comprehensively.",
        'Text Field14': f"Yes, {SAMPLE_ID}",
        'Text Field17': f"Yes, See {TEST_RECORD} for more information.",
        'Text Field18': "Yes, all analysts are trained and qualified by quality to perform the test.",
        'Text Field19': "Not Applicable",
        'Text Field20': "Not Applicable",
        'Text Field21': f"Yes, Information is available in Eagle Trax Sample Location History under {SAMPLE_ID}",
        'Text Field22': tf22_materials,
        'Text Field23': tf22_materials,
        'Text Field24': CONTROL_POSITIVE,
        'Text Field25': CONTROL_LOT,
        'Text Field26': CONTROL_EXP,
        'Text Field27': "Not Applicable",
        'Text Field28': "Not Applicable",
        'Text Field29': "Not Applicable",
        'Text Field30': f"E00{SCAN_ID}",
        'Text Field31': "Oct 2026",
        'Text Field32': "E001979 (CR145)\rE001978 (CR144)",
        'Text Field33': "Nov 2026\rNov 2026",
        'Text Field34': f"E00{SCAN_ID}",
        'Text Field35': "Oct 2026",
        'Text Field36': "Not Applicable",
        'Text Field37': "Not Applicable",
        'Text Field38': "Not Applicable",
        'Text Field39': "Not Applicable",
        'Text Field40': "See Phase I Summary",
        'Text Field41': "See Phase I Summary",
        'Text Field42': "See Phase I Summary",
        'Text Field43': tf43_equipment,
        'Text Field44': tf44_cal_due,
        'Text Field45': "See Phase I Summary",
        'Text Field46': "Not Applicable",
        'Text Field47': "Not Applicable",
        'Text Field48': "N/A QYC 22Oct26",
        'Text Field49': text_field_49,
        'Text Field50': text_field_50,
        'Text Field51': text_field_51,
        'Text Field53': "Qiyue Chen",
        'Text Field54': "Robin Seymour",
        'Text Field55': "",
        'Text Field52': "",
    }
    
    cb_whitelists = {
        # Page 1 Header
        'Check Box0': 'Yes',
        'Check Box1': 'Yes',
        'Check Box2': 'Yes',
        'Check Box3': 'Off',
        # Page 1 Section B (Personnel Discrepancy)
        'Check Box4': 'Yes', 'Check Box5': 'Off', 'Check Box6': 'Off',
        'Check Box7': 'Yes', 'Check Box8': 'Off', 'Check Box9': 'Off',
        'Check Box10': 'Yes', 'Check Box11': 'Off', 'Check Box12': 'Off',
        'Check Box13': 'Yes', 'Check Box14': 'Off', 'Check Box15': 'Off',
        'Check Box16': 'Yes', 'Check Box17': 'Off', 'Check Box18': 'Off',
        'Check Box19': 'Yes', 'Check Box20': 'Off', 'Check Box21': 'Off',
        'Check Box22': 'Off', 'Check Box23': 'Off', 'Check Box24': 'Yes',
        'Check Box25': 'Off', 'Check Box26': 'Off', 'Check Box27': 'Yes',
        'Check Box28': 'Yes', 'Check Box29': 'Off', 'Check Box30': 'Off',
        # Page 1 Section B (Material Discrepancy)
        'Check Box31': 'Off', 'Check Box32': 'Yes', 'Check Box33': 'Off',
        'Check Box34': 'Yes', 'Check Box35': 'Off', 'Check Box36': 'Off',
        'Check Box37': 'Off', 'Check Box38': 'Yes', 'Check Box39': 'Off',
        # Page 2 Checkboxes
        'Check Box40': 'Off', 'Check Box41': 'Off', 'Check Box42': 'Yes',
        'Check Box43': 'Yes', 'Check Box44': 'Off', 'Check Box45': 'Off',
        'Check Box46': 'Off', 'Check Box47': 'Off', 'Check Box48': 'Yes',
        'Check Box49': 'Off', 'Check Box50': 'Off', 'Check Box51': 'Yes',
        'Check Box52': 'Yes', 'Check Box53': 'Off', 'Check Box54': 'Off',
        'Check Box55': 'Yes', 'Check Box56': 'Off', 'Check Box57': 'Off',
        'Check Box58': 'Yes', 'Check Box59': 'Off', 'Check Box60': 'Off',
        'Check Box61': 'Off', 'Check Box62': 'Off', 'Check Box63': 'Yes',
        'Check Box64': 'Off', 'Check Box65': 'Off', 'Check Box66': 'Yes',
        'Check Box67': 'Yes', 'Check Box68': 'Off', 'Check Box69': 'Off',
        'Check Box70': 'Yes', 'Check Box71': 'Off', 'Check Box72': 'Off',
        'Check Box73': 'Yes', 'Check Box74': 'Off', 'Check Box75': 'Off',
        # Page 3 Checkboxes
        'Check Box78': 'Yes',
        'Check Box79': 'Yes',
        # Page 6 Section D & E
        'Check Box87': 'Off',
        'Check Box88': 'Yes',
        'Check Box89': 'Off',
        'Check Box90': 'Off',
        'Check Box91': 'Off',
        'Check Box92': 'Off',
    }

    # Font size calibration per Rule 74
    calibrated_font_sizes = {
        'Text Field32': 4.4,
        'Text Field33': 4.4,
        'Text Field49': 7.0,
        'Text Field50': 6.5,
    }

    # Populate first 6 pages of AcroForm
    for pno in range(6):
        page = doc[pno]
        for w in page.widgets():
            fn = w.field_name
            if fn in pdf_field_values:
                w.field_value = pdf_field_values[fn]
                if fn in calibrated_font_sizes:
                    w.text_fontsize = calibrated_font_sizes[fn]
                w.update()
            elif fn in cb_whitelists:
                w.field_value = cb_whitelists[fn]
                w.update()

    # Save 6-page filled form to temp file
    temp_6page_pdf = os.path.join(SCRIPT_DIR, f"temp_6page_{OOS_ID}.pdf")
    doc.save(temp_6page_pdf)
    doc.close()

    # Append Standalone Table 1 & Table 2 as Page 7
    tables_source_pdf = os.path.join(SCRIPT_DIR, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
    final_master_pdf = os.path.join(SCRIPT_DIR, f"CORP-FORM-21- 22 Oct 2026 OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
    
    writer = PdfWriter()
    r6 = PdfReader(temp_6page_pdf)
    for p in r6.pages[:6]:
        writer.add_page(p)
        
    rt = PdfReader(tables_source_pdf)
    writer.add_page(rt.pages[0])
        
    with open(final_master_pdf, "wb") as f:
        writer.write(f)
        
    print(f"Master 7-Page PDF generated: {final_master_pdf}")
    
    # Verify Page Count
    chk_doc = fitz.open(final_master_pdf)
    print(f"Verified Final Master PDF Page Count: {len(chk_doc)} (Strict requirement: 7)")
    assert len(chk_doc) == 7, f"ERROR: Master PDF must be 7 pages, got {len(chk_doc)}"
    
    # Render preview images for verification
    preview_dir = os.path.join(SCRIPT_DIR, f"preview_oos_{OOS_ID}")
    os.makedirs(preview_dir, exist_ok=True)
    for i, p in enumerate(chk_doc):
        pix = p.get_pixmap(dpi=150)
        pix.save(os.path.join(preview_dir, f"page_{i+1}.png"))
    print(f"Rendered all 7 pages to: {preview_dir}")
    chk_doc.close()

    # Sync to Desktop and Root
    dst_desktop = os.path.join(DESKTOP_DIR, os.path.basename(final_master_pdf))
    dst_root = os.path.join(ROOT_DIR, os.path.basename(final_master_pdf))
    shutil.copy2(final_master_pdf, dst_root)
    try:
        shutil.copy2(final_master_pdf, dst_desktop)
        print("Master PDF synced to Desktop:", dst_desktop)
    except Exception as e:
        print("Desktop PDF lock note:", e)

    return final_master_pdf

if __name__ == "__main__":
    print(f"=== Generating Master OOS Package for OOS-{OOS_ID} ({CLIENT_NAME}) ===")
    docx_path = generate_master_word_doc()
    pdf_path = generate_master_pdf_report()
    print("\n=== Generation and Synchronization Complete! ===")
