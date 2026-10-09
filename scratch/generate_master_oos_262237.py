"""
Script: generate_master_oos_262237.py
Purpose: Complete generation of Master Word Document and 7-Page AcroForm PDF
         for Scan RDI OOS-262237 (GoGoMeds Select, Semaglutide Lot 20066).
"""

import os
import sys
import json
import re
import copy
import shutil
from datetime import datetime
import fitz  # PyMuPDF
import docx
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
from docxtpl import DocxTemplate
from pypdf import PdfWriter, PdfReader
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"

OOS_ID = "262237"
CLIENT_NAME = "GoGoMeds Select (E10747)"
SAMPLE_ID = "ETX-260921-0148"
SAMPLE_URL = "https://etrax.eagleanalytical.com/Submission/Details/08VuIoFsBaXGehnOxJ3PMQ__"
SAMPLE_NAME = "Semaglutide R 2.5mg/ml Benzyl Alcohol 0.9%"
LOT_NUMBER = "20066"
DOSAGE_FORM = "Injectable"
TEST_DATE = "24Sep26"
TEST_DATE_FULL = "24-Sep-2026"
TEST_RECORD = "092426-2132-3"

PREPPER_NAME = "Eniola Nioupin"
PREPPER_INIT = "ENI"
PROCESSOR_NAME = "Guanchen Li"
PROCESSOR_INIT = "GL"
CHANGEOVER_NAME = "Muralidhar Bythatagari"
CHANGEOVER_INIT = "MRB"
READER_NAME = "Muralidhar Bythatagari"
READER_INIT = "MRB"

BSC_PROC = "1312"
BSC_CHG = "1937"
SCAN_ID = "2132"
MONTHLY_CLEANING_DATE = "13Sep26"
MONTHLY_CLEANING_FULL = "13 Sep 2026"

EVENTS_COUNT = "362"
CONFIRMED_COUNT = "90"
MORPHOLOGY = "curved rod-shaped morphology"

# -------------------------------------------------------------
# 1. Narrative Text Blocks
# -------------------------------------------------------------
# Page 3 Paragraphs (p1 to p7)
p1 = (
    "All analysts involved in the preparation, processing, changeover, and reading of the sample – "
    f"{PREPPER_NAME}, {PROCESSOR_NAME}, and {READER_NAME} – were interviewed comprehensively. "
    "Their responses are documented throughout this report. This investigation was performed in "
    "accordance with MICRO-SOP-53, Sterility Test Out-of-Specification Investigation Procedure."
)

p2 = (
    "Upon arrival, the sample was stored refrigerated in accordance with the Client's instructions. "
    f"Prepping analyst {PREPPER_NAME} and processing analyst {PROCESSOR_NAME} verified the integrity "
    "of the containers throughout both the preparation and processing stages. No leaks, cracks, or "
    "visible turbidity were observed prior to testing, verifying container integrity."
)

p3 = (
    "All reagents and supplies mentioned in the material section above were stored according to suppliers’ "
    "recommendations, and their integrity was visually verified before utilization. Moreover, each reagent, "
    "media, and consumable possessed valid expiration dates. The functionality of all equipment was confirmed "
    "by reviewing continuous environmental and operational monitoring data."
)

p4 = (
    f"During the preparation phase, {PREPPER_NAME} disinfected the sample containers using acidified bleach "
    "and placed them into a pre-disinfected storage bin. On 24Sep26, prior to sample processing, "
    f"{PROCESSOR_NAME} performed a second disinfection with acidified bleach, allowing a minimum contact "
    "time of 10 minutes before transferring the samples into the cleanroom suites. A final disinfection step "
    f"was completed immediately before introducing the sample into the ISO 5 Biological Safety Cabinet "
    f"(BSC E00{BSC_PROC}) located in the innermost ISO 7 cleanroom (116B). All activities were performed in "
    "accordance with MICRO-SOP-12, Rapid Scan RDI® Test using FIFU Method."
)

p5 = (
    "Cleanroom Suite 116, used for sample processing, comprises the innermost ISO 7 cleanroom 116B, middle ISO 7 "
    "buffer room 116A, and outermost ISO 8 anteroom 116. Cleanroom L-Suite, used for changeover, consists of ISO 7 "
    "cleanroom 145, ISO 7 buffer room 144, ISO 8 anteroom 143, and outermost ISO 8 room 142. Outward-cascading "
    "positive air pressure differentials are continuously maintained across both suites to ensure controlled "
    "unidirectional airflow."
)

p6 = (
    f"The ISO 5 BSC E00{BSC_PROC} located in Cleanroom 116B (Suite 116) and ISO 5 BSC E00{BSC_CHG} located in "
    "Cleanroom 144 (L-Suite) were thoroughly cleaned and disinfected prior to their respective procedures in "
    "accordance with MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. Both BSCs were certified "
    "and approved by the Engineering and Quality Assurance teams prior to testing."
)

p7 = (
    f"Pre-labeling and filtration activities were performed by {PROCESSOR_NAME} on 24Sep26 within the ISO 5 "
    f"biosafety cabinet (BSC E00{BSC_PROC}) located in the innermost ISO 7 cleanroom (116B). The changeover procedure "
    f"was subsequently performed by analyst {CHANGEOVER_NAME} on the same date within the ISO 5 biosafety cabinet "
    f"(BSC E00{BSC_CHG}) located in Cleanroom L-Suite. The reading analyst, {READER_NAME}, confirmed that the "
    f"ScanRDI instrument E00{SCAN_ID} was set up as per ENG-SOP-4 (Scan RDI® System – Operations (Standard C3 Quality "
    f"Check and Microscope Setup and Maintenance)), and the negative control and positive control for analyst {PROCESSOR_NAME} "
    "yielded expected results."
)

# Page 4 / 5 Paragraphs (p8 to p17)
p8 = (
    f"On 24Sep26, a rapid sterility test was performed on the sample using the ScanRDI method under test record {TEST_RECORD}. "
    f"The sample was initially prepared by analyst {PREPPER_NAME}, processed by {PROCESSOR_NAME} (27th sample processed), "
    f"changeover performed by {CHANGEOVER_NAME}, and subsequently read by {READER_NAME}. The test revealed "
    f"{EVENTS_COUNT} total events and {CONFIRMED_COUNT} confirmed microbial events exhibiting {MORPHOLOGY}, see Table 1."
)

p9 = (
    f"Table 2 presents the environmental monitoring results associated with testing performed on {TEST_DATE}. "
    "The environmental monitoring (EM) plates were incubated for no less than 48 hours at 30–35°C and for no less than "
    "an additional five days at 20–25°C, as per SOP 2.600.002, Environmental Monitoring of the Clean-room Facility."
)

p10 = (
    "Upon review of the environmental monitoring data, 100% absence of microbial recovery (No Growth) was demonstrated "
    "across all critical ISO 5 zones and personnel monitoring on the date of testing. Daily personnel monitoring for "
    f"processing analyst {PROCESSOR_NAME} and changeover analyst {CHANGEOVER_NAME} yielded no microbial growth on left "
    f"and right fingertip touch plates. Daily surface sampling (4 locations) and settling plates within processing BSC E00{BSC_PROC} "
    f"and changeover BSC E00{BSC_CHG} yielded no microbial growth."
)

p11 = (
    "Weekly surface monitoring of Cleanroom Suite 116 conducted on 24Sep26 and Cleanroom L-Suite conducted on 22Sep26 "
    "demonstrated no microbial recovery across all sampled cleanroom surfaces, confirming effective sanitization and surface control."
)

p12 = (
    "Weekly active air monitoring of Cleanroom Suite 116 conducted on 24Sep26 demonstrated no microbial recovery in the "
    "ISO 7 cleanrooms; however, 3 CFUs (ETX-261005-0773, [Pending Differential Staining: Gram (+/-) ... / Pending Microbial ID]) "
    "were recovered from the ISO 8 anteroom (116). Weekly active air monitoring of Cleanroom L-Suite conducted on 22Sep26 demonstrated no "
    "microbial recovery in the ISO 7 cleanrooms (145 or 144); however, 6 CFUs (ETX-260929-0335, [Pending Differential Staining: "
    "Gram (+/-) ... / Pending Microbial ID]) and 8 CFUs (ETX-260929-0341, [Pending Differential Staining: Gram (+/-) ... / "
    "Pending Microbial ID]) were recovered from the ISO 8 anteroom (143 Section I and Section II, respectively), and 2 CFUs "
    "(ETX-260929-0344, [Pending Differential Staining: Gram (+/-) ... / Pending Microbial ID]) were recovered from the outermost "
    "ISO 8 room (142)."
)

p13 = (
    f"It is important to emphasize that all sample processing activities were performed strictly within the validated ISO 5 "
    f"BSC E00{BSC_PROC} located in the innermost ISO 7 cleanroom (116B), and changeover activities were performed within ISO 5 "
    f"BSC E00{BSC_CHG} located in the ISO 7 cleanroom of L-Suite. The test sample does not come into contact with ambient ISO 8 "
    "air, as samples and supplies are transferred in disinfected, closed containers on carts through the layered cleanroom "
    "suites. Furthermore, the recoveries occurred in the lower-classified ISO 8 anteroom and outer room environments, which "
    "are physically segregated from the ISO 5 processing zone by closed doors and an outward-cascading positive air pressure "
    "gradient. In conjunction with the complete absence of microbial growth on all analyst glove touch plates, settling plates, "
    "and ISO 5 BSC work surfaces throughout testing, there is no plausible contamination transfer pathway from the ambient "
    "ISO 8 environment to the test sample."
)

p14 = (
    f"Monthly cleaning and disinfection of the cleanrooms (ISO 7) and their containing biosafety cabinets (ISO 5) were performed "
    f"on {MONTHLY_CLEANING_DATE}, as per MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. It was documented that "
    "all chemical indicators passed, verifying that environmental sanitization remained validated and in control prior to testing."
)

p15 = (
    f"Analyzing a 6-month sample history for {CLIENT_NAME}, this specific analyte \"{SAMPLE_NAME}\" has had no prior failures "
    "using the Scan RDI method during this period."
)

p16 = (
    "The microbiological findings were also evaluated for evidence of sample-to-sample cross-contamination. The test sample "
    f"{SAMPLE_ID} was the 27th sample processed by {PROCESSOR_NAME} on 24Sep26. All other samples processed by the same analyst, "
    "as well as samples processed by other laboratory personnel on the date of testing, yielded negative results. Analyst "
    f"{PROCESSOR_NAME} verified that gloves were disinfected between samples and standard aseptic techniques were strictly "
    "maintained. The absence of additional positive samples or clustering patterns provides further evidence against cross-contamination."
)

p17 = (
    "Based on the comprehensive laboratory investigation, all media, reagents, consumables, and equipment met validated acceptance "
    "criteria. Aseptic protocols were strictly adhered to with no documented deviations, and environmental controls in the critical "
    "ISO 5 zones remained fully effective. No assignable laboratory error was identified. Consequently, the microbial recovery "
    f"observed in sample {SAMPLE_ID} is considered an isolated occurrence within the sample itself, and the original failing test "
    "result is deemed valid."
)

# Text Field partitions for PDF
text_field_49 = "\r \r".join([p1, p2, p3, p4, p5, p6, p7])
text_field_50 = "\r \r".join([p8, p9, p10, p11, p12, p13])
text_field_51 = "\r \r".join([p14, p15, p16, p17])

smart_phase1_full = "\n\n".join([p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17])

print("Narrative text blocks prepared successfully.")

# -------------------------------------------------------------
# 2. Render Master Word Document (ScanRDI OOS P1 template.docx)
# -------------------------------------------------------------
def generate_master_word_doc():
    template_path = os.path.join(ROOT_DIR, "ScanRDI OOS P1 template.docx")
    if not os.path.exists(template_path):
        template_path = os.path.join(ROOT_DIR, "ScanRDI OOS template.docx")
    
    tpl = DocxTemplate(template_path)
    
    personnel_block = (
        f"Prepping Analyst:\n{PREPPER_NAME} ({PREPPER_INIT})\n\n"
        f"Processing Analyst:\n{PROCESSOR_NAME} ({PROCESSOR_INIT})\n\n"
        f"Changeover Analyst:\n{CHANGEOVER_NAME} ({CHANGEOVER_INIT})\n\n"
        f"Reading Analyst:\n{READER_NAME} ({READER_INIT})"
    )
    
    analyst_sig_text = f"{PROCESSOR_NAME} (Written by: Qiyue Chen)"
    
    context = {
        "oos_id": OOS_ID,
        "client_name": CLIENT_NAME,
        "sample_id": SAMPLE_ID,
        "sample_url": SAMPLE_URL,
        "sample_name": SAMPLE_NAME,
        "lot_number": LOT_NUMBER,
        "dosage_form": DOSAGE_FORM,
        "test_date": TEST_DATE,
        "test_record": TEST_RECORD,
        "report_header": f"{SAMPLE_ID}\n\n{CLIENT_NAME}",
        "analyst_signature": analyst_sig_text,
        "analyst_initial": PROCESSOR_INIT,
        "analyst_name": PROCESSOR_NAME,
        "changeover_initial": CHANGEOVER_INIT,
        "changeover_name": CHANGEOVER_NAME,
        "reader_name": READER_NAME,
        "prepper_initial": PREPPER_INIT,
        "prepper_name": PREPPER_NAME,
        "bsc_id": BSC_PROC,
        "chgbsc_id": BSC_CHG,
        "scan_id": SCAN_ID,
        "smart_scan_id": f"E00{SCAN_ID}",
        "cr_id": "116B",
        "cr_suit": "116",
        "smart_cr_id": "CR116 (E001738) & L-Suite (CR1978)",
        "smart_bsc_bracketing_header": f"Biological Safety Cabinet EM Bracketing Biological Safety Cabinet (BSC) E00{BSC_PROC} and E00{BSC_CHG}",
        "control_positive": "A. brasiliensis",
        "control_lot": "24MAY28-01",
        "control_data": "24May28",
        "event_number": EVENTS_COUNT,
        "confirm_number": CONFIRMED_COUNT,
        "organism_morphology": "Curved rod",
        "smart_personnel_block": personnel_block,
        "smart_incident_opening": f"On {TEST_DATE}, sample {SAMPLE_ID} was found positive for viable microorganisms after Scan RDI sterility testing.",
        "smart_comment_interview": f"Yes, analysts {PREPPER_NAME}, {PROCESSOR_NAME}, and {READER_NAME} were interviewed comprehensively.",
        "smart_comment_samples": f"Yes, {SAMPLE_ID}",
        "smart_comment_records": f"Yes, See {TEST_RECORD} for more information.",
        "smart_comment_storage": f"Yes, Information is available in Eagle Trax Sample Location History under {SAMPLE_ID}",
        "smart_phase1_summary": smart_phase1_full,
        "smart_phase1_continued": smart_phase1_full,
    }
    
    tpl.render(context)
    
    # Replace Table 1 & Table 2 with verified standalone tables
    tables_source_docx = os.path.join(DESKTOP_DIR, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.docx")
    if os.path.exists(tables_source_docx):
        src_doc = docx.Document(tables_source_docx)
        
        # In tpl.docx: Table 0 is header info, Table 1 is sample info, Table 2 is EM table
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
            
            # Remove any trailing trend table if present
            if len(tpl.docx.tables) >= 4:
                t3_elem = tpl.docx.tables[3]._element
                t3_elem.getparent().remove(t3_elem)
                
    out_docx_desktop = os.path.join(DESKTOP_DIR, f"OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.docx")
    out_docx_docs = os.path.join(DOCUMENTS_DIR, f"OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.docx")
    
    tpl.save(out_docx_docs)
    try:
        tpl.save(out_docx_desktop)
        print("Master Word Docx saved to Desktop:", out_docx_desktop)
    except Exception as e:
        print("Desktop Docx lock note:", e)

    return out_docx_docs

# -------------------------------------------------------------
# 3. Generate Complete 7-Page PDF Report (AcroForm + Tables)
# -------------------------------------------------------------
def generate_master_pdf_report():
    source_pdf = os.path.join(ROOT_DIR, "ScanRDI OOS template.pdf")
    if not os.path.exists(source_pdf):
        source_pdf = os.path.join(ROOT_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")
        
    doc = fitz.open(source_pdf)
    
    analyst_sig_text = f"{PROCESSOR_NAME} (Written by: Qiyue Chen)"
    personnel_block = (
        f"Prepping Analyst:\n{PREPPER_NAME} ({PREPPER_INIT})\n\n"
        f"Processing Analyst:\n{PROCESSOR_NAME} ({PROCESSOR_INIT})\n\n"
        f"Changeover Analyst:\n{CHANGEOVER_NAME} ({CHANGEOVER_INIT})\n\n"
        f"Reading Analyst:\n{READER_NAME} ({READER_INIT})"
    )
    
    pdf_field_values = {
        'Text Field57': OOS_ID,
        'Date Field0': TEST_DATE_FULL,
        'Date Field1': TEST_DATE_FULL,
        'Date Field2': TEST_DATE_FULL,
        'Date Field3': TEST_DATE_FULL,
        'Text Field2': SAMPLE_ID,
        'Text Field6': LOT_NUMBER,
        'Text Field4': SAMPLE_NAME,
        'Text Field5': DOSAGE_FORM,
        'Text Field8': "MICRO-SOP-12 (16)\rENG-SOP-4 (05)",
        'Text Field9': "24Jul26\r24Jul26",
        'Text Field10': "Rev: 16\rRev: 05",
        'Text Field15': "Yes, as per MICRO-SOP-12, ENG-SOP-4",
        'Text Field16': "Yes, as per MICRO-SOP-12, ENG-SOP-4",
        'Text Field0': READER_NAME,  # Initiator
        'Text Field3': personnel_block,
        'Text Field7': f"On {TEST_DATE}, sample {SAMPLE_ID} was found positive for viable microorganisms after Scan RDI sterility testing.",
        'Text Field13': f"Yes, analysts {PREPPER_NAME}, {PROCESSOR_NAME}, and {READER_NAME} were interviewed comprehensively.",
        'Text Field14': f"Yes, {SAMPLE_ID}",
        'Text Field17': f"Yes, See {TEST_RECORD} for more information.",
        'Text Field21': f"Yes, Information is available in Eagle Trax Sample Location History under {SAMPLE_ID}",
        'Text Field30': f"E00{SCAN_ID}",
        'Text Field32': f"E001738 (Suite 116) & E001978 (L-Suite)",
        'Text Field34': f"E00{SCAN_ID}",
        'Text Field24': "A. brasiliensis",
        'Text Field25': "24MAY28-01",
        'Text Field26': "24May28",
        'Text Field48': f"N/A QYC {datetime.now().strftime('%d%b%y')}",
        'Text Field49': text_field_49,
        'Text Field50': text_field_50,
        'Text Field51': text_field_51,
        'Text Field53': analyst_sig_text,
    }
    
    # Strict Whitelist Checkbox configurations (Rule 10)
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
    }
    
    # Apply field values and font sizes using PyMuPDF for natural appearance
    for page in doc:
        for widget in page.widgets():
            fn = widget.field_name
            if fn in pdf_field_values:
                widget.field_value = pdf_field_values[fn]
                # Calibrated font sizing to prevent Acrobat + overflow
                if fn in ['Text Field49', 'Text Field50', 'Text Field51']:
                    widget.text_fontsize = 9.2
                    widget.text_font = "Helv"
                elif fn == 'Text Field3':
                    widget.text_fontsize = 6.5
                elif fn == 'Text Field7':
                    widget.text_fontsize = 8.0
                elif fn in ['Text Field13', 'Text Field14', 'Text Field15', 'Text Field16', 'Text Field17', 'Text Field21']:
                    widget.text_fontsize = 7.5
                widget.update()
            elif fn in cb_whitelists:
                widget.field_value = cb_whitelists[fn]
                widget.update()
                
    temp_filled_pdf = os.path.join(SCRIPT_DIR, "temp_filled_master_262237.pdf")
    doc.save(temp_filled_pdf)
    doc.close()
    
    # Now append Standalone Tables PDF as Page 7
    tables_pdf_desktop = os.path.join(DESKTOP_DIR, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
    writer = PdfWriter(clone_from=temp_filled_pdf)
    
    if os.path.exists(tables_pdf_desktop):
        t_reader = PdfReader(tables_pdf_desktop)
        for t_page in t_reader.pages:
            writer.add_page(t_page)
        print("Appended Table 1 & Table 2 as Page 7. Total pages:", len(writer.pages))
        
    out_pdf_desktop = os.path.join(DESKTOP_DIR, f"OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
    out_pdf_docs = os.path.join(DOCUMENTS_DIR, f"OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
    
    with open(out_pdf_docs, "wb") as f:
        writer.write(f)
    print("Master PDF saved to Documents:", out_pdf_docs)
    
    try:
        with open(out_pdf_desktop, "wb") as f:
            writer.write(f)
        print("Master PDF saved to Desktop:", out_pdf_desktop)
    except Exception as e:
        print("Desktop PDF lock note:", e)

if __name__ == "__main__":
    print("=== Generating Master Word Report and Master PDF Report for OOS-262237 ===")
    generate_master_word_doc()
    generate_master_pdf_report()
    print("Master reports generation completed successfully!")
