r"""
Script: transcribe_to_corp_form_2150.py
Purpose: Transcribe OOS-262150 data into the official ZenQMS CORP-FORM-21 template
         (C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 - P1 - 16 Sep 2026.pdf),
         leaving Page 6 Reviewer comments and Closure checkboxes blank (since no one has reviewed it yet),
         and appending Table 1 & Table 2 as Page 7.
"""

import os
import sys
import copy
import shutil
from datetime import datetime
import fitz  # PyMuPDF
from pypdf import PdfWriter, PdfReader

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"

OOS_ID = "262150"
CLIENT_NAME = "BelieveRX Pharmacy (E74120)"
SAMPLE_ID = "ETX-260914-0306"
SAMPLE_NAME = "Tirzepatide 5mg, Glycine 0.5mg/mL"
LOT_NUMBER = "091126-66A"
DOSAGE_FORM = "Injectable"
TEST_DATE = "16Sep26"
TEST_DATE_FULL = "16-Sep-2026"
TEST_RECORD = "091626-2132-2"

PREPPER_NAME = "Eniola Nioupin"
PREPPER_INIT = "ENI"
PROCESSOR_NAME = "Elizabeth Bennett"
PROCESSOR_INIT = "ELB"
CHANGEOVER_NAME = "Elizabeth Bennett"
CHANGEOVER_INIT = "ELB"
READER_NAME = "Sonal Uprety"
READER_INIT = "SU"

BSC_PROC = "1938"
BSC_CHG = "1937"
SCAN_ID = "2132"
MONTHLY_CLEANING_DATE = "30Aug26"
MONTHLY_CLEANING_FULL = "30 Aug 2026"

EVENTS_COUNT = "56"
CONFIRMED_COUNT = "5"
MORPHOLOGY = "cocci shaped morphology"

# -------------------------------------------------------------
# 1. Narrative Text Blocks
# -------------------------------------------------------------
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
    "and placed them into a pre-disinfected storage bin. On 16Sep26, prior to sample processing, "
    f"{PROCESSOR_NAME} performed a second disinfection with acidified bleach, allowing a minimum contact "
    "time of 10 minutes before transferring the samples into the cleanroom suite. A final disinfection step "
    f"was completed immediately before introducing the sample into the ISO 5 Biological Safety Cabinet "
    f"(BSC E00{BSC_PROC}) located in the innermost ISO 7 cleanroom (145). All activities were performed in "
    "accordance with MICRO-SOP-12, Rapid Scan RDI® Test using FIFU Method."
)

p5 = (
    "Cleanroom L-Suite consists of ISO 7 cleanroom 145, ISO 7 buffer room 144, ISO 8 anteroom 143, and "
    "outermost ISO 8 room 142. Outward-cascading positive air pressure differentials are continuously "
    "maintained across the suite to ensure controlled unidirectional airflow."
)

p6 = (
    f"The ISO 5 BSC E00{BSC_PROC} located in Cleanroom 145 (L-Suite) and ISO 5 BSC E00{BSC_CHG} located in "
    "Buffer Room 144 (L-Suite) were thoroughly cleaned and disinfected prior to their respective procedures in "
    "accordance with MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. Both BSCs were certified "
    "and approved by the Engineering and Quality Assurance teams prior to testing."
)

p7 = (
    f"Pre-labeling and filtration activities were performed by {PROCESSOR_NAME} on 16Sep26 within the ISO 5 "
    f"biosafety cabinet (BSC E00{BSC_PROC}) located in ISO 7 cleanroom (145). The changeover procedure "
    f"was subsequently performed by analyst {CHANGEOVER_NAME} on the same date within the ISO 5 biosafety cabinet "
    f"(BSC E00{BSC_CHG}) located in Buffer Room 144 of L-Suite. The reading analyst, {READER_NAME}, confirmed that the "
    f"ScanRDI instrument E00{SCAN_ID} was set up as per ENG-SOP-4 (Scan RDI® System – Operations (Standard C3 Quality "
    f"Check and Microscope Setup and Maintenance)), and the negative control and positive control for analyst {PROCESSOR_NAME} "
    "yielded expected results."
)

# Page 4 Paragraphs (p8 to p16)
p8 = (
    f"On 16Sep26, a rapid sterility test was performed on the sample using the ScanRDI method under test record {TEST_RECORD}. "
    f"The sample was initially prepared by analyst {PREPPER_NAME}, processed by {PROCESSOR_NAME} (47th sample processed), "
    f"changeover performed by {CHANGEOVER_NAME}, and subsequently read by {READER_NAME}. The test revealed "
    f"{EVENTS_COUNT} total events and {CONFIRMED_COUNT} confirmed microbial events exhibiting {MORPHOLOGY}, see Table 1."
)

p9 = (
    f"Table 2 (see attached tables) presents the environmental monitoring results for {SAMPLE_ID}. "
    "The environmental monitoring (EM) plates were incubated for no less than 48 hours at 30–35°C and for no less than "
    "an additional five days at 20–25°C, as per SOP 2.600.002, Environmental Monitoring of the Clean-room Facility."
)

p10 = (
    "Upon review of the environmental monitoring data associated with the sterility test, no microbial growth was recovered from "
    f"any of the four ISO 5 work surface monitoring locations within processing BSC E00{BSC_PROC} located in room CR145 or within "
    f"changeover BSC E00{BSC_CHG} located in CR144 on the date of testing. In addition, no microbial growth was observed on the "
    f"processing or changeover personnel monitoring (left- and right-hand touch plates for analyst {PROCESSOR_NAME}) or settling plates "
    f"within BSC E00{BSC_PROC} and BSC E00{BSC_CHG}."
)

p11 = (
    "Environmental Monitoring (EM) Recoveries Review:\n\n"
    "ISO 5 BSC and Personnel Evaluation:\n"
    f"Personnel monitoring was performed on 16Sep26 for analyst {PROCESSOR_NAME}. No microbial growth was recovered from the "
    f"left- or right-hand touch plates for processing or changeover activities. Therefore, personnel monitoring did not identify "
    f"a potential operator-borne source of the {MORPHOLOGY} observed in test sample {SAMPLE_ID}. The absence of microbial "
    "growth from personnel touch plates, ISO 5 BSC work surfaces, and settling plates indicates that no microbiological recovery "
    "was detected from the monitored critical processing environments on the date of testing. Therefore, the contemporaneous "
    "personnel and ISO 5 environmental monitoring results do not support processing personnel, the ISO 5 BSC used for sample "
    f"processing, or the changeover BSC as a source of the microorganisms observed in {SAMPLE_ID}."
)

p12 = (
    "Cleanroom EM Evaluation:\n"
    "Routine weekly active-air monitoring of the L-Suite was performed on 15Sep26, during the week of testing. Microbial recoveries "
    "were detected only in lower-classified ISO 8 background areas: CR143, section I (6 CFUs); CR143, section II (21 CFUs); and "
    "CR142 (9 CFUs). Weekly surface sampling of Cleanroom L-Suite, performed on 15Sep26, demonstrated no microbial growth across all "
    f"locations. Importantly, ISO 5 BSC E00{BSC_PROC} and ISO 5 BSC E00{BSC_CHG}, where the sample was handled during processing "
    "and changeover, demonstrated no growth from all monitored work-surface locations and settling plates. Although some active-air "
    "recoveries from ISO 8 CR142 and CR143 included cocci-form organisms, including Micrococcus, Kocuria, and Staphylococcus species, "
    "these recoveries were confined to lower-classified background locations rather than the ISO 5 critical work zones. Because these "
    "recoveries were physically segregated from the ISO 5 processing environments and no corresponding microbial recovery occurred in the "
    f"ISO 5 BSCs or personnel monitoring, the available EM data do not establish a direct microbiological link between the ISO 8 "
    f"background-area recoveries and the {MORPHOLOGY} observed in {SAMPLE_ID}. Additionally, there is minimal contact with the "
    "ambient ISO 8 air, as samples and supplies are transferred in disinfected, closed containers on carts through the layered cleanroom suite."
)

p13 = (
    f"Monthly cleaning and disinfection, using H2O2, of the cleanrooms (ISO 7s) containing Biosafety Cabinets (BSCs, ISO 5) were performed "
    f"on {MONTHLY_CLEANING_DATE}, as per MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. It was documented that all "
    "H2O2 indicators passed."
)

p14 = (
    f"A review of the six-month testing history for {CLIENT_NAME} found no prior ScanRDI method failures for the \"{SAMPLE_NAME}\" product."
)

p15 = (
    "To evaluate the potential for sample-to-sample contamination, all samples processed on the same day were reviewed. The test sample "
    f"{SAMPLE_ID} was the 47th sample processed by {PROCESSOR_NAME} on 16Sep26. All other samples processed by the same analyst and by other "
    "analysts on that day yielded negative results, indicating that cross-contamination is unlikely. Analyst "
    f"{PROCESSOR_NAME} verified that gloves were disinfected between samples and standard aseptic techniques were strictly maintained."
)

p16 = (
    "Based on the cumulative evidence, it is highly unlikely that the failing results were due to reagents, supplies, the cleanroom "
    "environment, the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is "
    "minimal. Therefore, the original test result is deemed valid."
)

text_field_49 = "\r \r".join([p1, p2, p3, p4, p5, p6, p7])
text_field_50 = "\r \r".join([p8, p9, p10, p11, p12, p13, p14, p15, p16])
text_field_51 = f"N/A QYC {datetime.now().strftime('%d%b%Y')}"

personnel_block = (
    f"Prepping Analyst:\n{PREPPER_NAME} ({PREPPER_INIT})\n\n"
    f"Processing Analyst:\n{PROCESSOR_NAME} ({PROCESSOR_INIT})\n\n"
    f"Changeover Analyst:\n{CHANGEOVER_NAME} ({CHANGEOVER_INIT})\n\n"
    f"Reading Analyst:\n{READER_NAME} ({READER_INIT})"
)

# -------------------------------------------------------------
# 2. PDF Field Values and Checkboxes
# -------------------------------------------------------------
pdf_field_values = {
    'Text Field57': OOS_ID,
    'Date Field0': TEST_DATE_FULL,
    'Date Field1': TEST_DATE_FULL,
    'Date Field2': TEST_DATE_FULL,
    'Date Field3': TEST_DATE_FULL,
    'Text Field1': "Scan RDI Sterility Test",
    'Text Field2': SAMPLE_ID,
    'Text Field6': LOT_NUMBER,
    'Text Field4': SAMPLE_NAME,
    'Text Field5': DOSAGE_FORM,
    'Text Field8': "MICRO-SOP-12 (16)\rENG-SOP-4 (05)",
    'Text Field9': "24Jul26\r24Jul26",
    'Text Field10': "Rev: 16\rRev: 05",
    'Text Field11': "<1 event",
    'Text Field12': "Kathan Parikh",
    'Date Field4': TEST_DATE_FULL,
    'Text Field0': f"{PROCESSOR_NAME} (Written by: Qiyue Chen)",  # Initiator
    'Text Field3': personnel_block,
    'Text Field7': f"On {TEST_DATE}, sample {SAMPLE_ID} was found positive for viable microorganisms after Scan RDI sterility testing.",
    'Text Field13': f"Yes, analysts {PREPPER_NAME}, {PROCESSOR_NAME}, and {READER_NAME} were interviewed comprehensively.",
    'Text Field14': f"Yes, {SAMPLE_ID}",
    'Text Field15': "Yes, as per MICRO-SOP-12, ENG-SOP-4",
    'Text Field16': "Yes, as per MICRO-SOP-12, ENG-SOP-4",
    'Text Field17': f"Yes, See {TEST_RECORD} for more information.",
    'Text Field18': "Yes, all analysts are trained and qualified by quality to perform the test.",
    'Text Field19': "Not Applicable",
    'Text Field20': "Not Applicable",
    'Text Field21': f"Yes, Information is available in Eagle Trax Sample Location History under {SAMPLE_ID}",
    'Text Field22': "Scan Consumables:\nSee the attached data packet\n\nEnvironmental Plates:\nTSA and Surface Plate: see attached environmental logs\n\nMonthly Cleaning:\nH2O2 strips and IPA: See attached monthly cleaning logs",
    'Text Field23': "Scan Consumables:\nSee the attached data packet\n\nEnvironmental Plates:\nTSA and Surface Plate: see attached environmental logs\n\nMonthly Cleaning:\nH2O2 strips and IPA: See attached monthly cleaning logs",
    'Text Field24': "S.aureus",
    'Text Field25': "04232026-6538-SA",
    'Text Field26': "23 Apr 2028",
    'Text Field27': "Not Applicable",
    'Text Field28': "Not Applicable",
    'Text Field29': "Not Applicable",
    'Text Field30': f"E00{SCAN_ID}",
    'Text Field31': "July 2027",
    'Text Field32': "E001979 (CR145),  E001978 (CR144)\nE001977 (CR143),  E001976 (CR142)",
    'Text Field33': "Nov 2026,  Nov 2026\nNov 2026,  Nov 2026",
    'Text Field34': f"E00{SCAN_ID}",
    'Text Field35': "July 2027",
    'Text Field36': "Not Applicable",
    'Text Field37': "Not Applicable",
    'Text Field38': "Not Applicable",
    'Text Field39': "Not Applicable",
    'Text Field40': "See Phase I Summary",
    'Text Field41': "See Phase I Summary",
    'Text Field42': "See Phase I Summary",
    'Text Field43': "Incubator E001034\r \r(Sensor E001501)\r \rIncubator E001031\r \r(Sensor E001505)",
    'Text Field44': "Aug 2027\r \rFeb 2027\r \rAug 2027\r \rFeb 2027",
    'Text Field45': "See Phase I Summary",
    'Text Field46': "Not Applicable",
    'Text Field47': "Not Applicable",
    'Text Field48': f"N/A QYC {datetime.now().strftime('%d%b%Y')}",
    'Text Field49': text_field_49,
    'Text Field50': text_field_50,
    'Text Field51': text_field_51,
    # CRITICAL: We are the WRITER. No one has reviewed it yet.
    # Therefore, Text Field52 (Reviewer comments) is BLANK!
    'Text Field52': "",
    'Text Field53': "Qiyue Chen",  # Prepared by (Print)
    'Text Field54': "",            # Lab Manager/Designee (Print) - Blank
    'Text Field55': "",            # QA Representative (Print) - Blank
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
    # Page 3 Top & Section C Checkboxes
    'Check Box76': 'Off', 'Check Box77': 'Off', 'Check Box78': 'Yes',
    'Check Box79': 'Yes',
    'Check Box80': 'Off', 'Check Box81': 'Off', 'Check Box82': 'Off',
    'Check Box83': 'Off', 'Check Box84': 'Off', 'Check Box85': 'Off', 'Check Box86': 'Off',
    # Page 6 Checkboxes:
    # Has root cause been identified as possible laboratory error? No, OOS is valid.
    'Check Box87': 'Off', 'Check Box88': 'Yes', 'Check Box89': 'Off',
    # Section E: Quality Disposition - NOT REVIEWED/CLOSED YET -> ALL OFF!
    'Check Box90': 'Off', 'Check Box91': 'Off', 'Check Box92': 'Off',
}

calibrated_fonts = {
    # Page 1 Header
    'Text Field0': 5.12,
    'Text Field1': 7.5,
    'Text Field2': 9.65,
    'Text Field3': 6.0,
    'Text Field4': 8.5,
    'Text Field5': 9.5,
    'Text Field6': 9.5,
    'Text Field7': 7.5,
    'Text Field8': 5.18,
    'Text Field9': 5.18,
    'Text Field10': 5.18,
    'Text Field11': 9.5,
    'Text Field12': 10.0,
    'Text Field57': 6.0,
    'Date Field0': 10.0,
    'Date Field1': 10.0,
    'Date Field2': 10.0,
    'Date Field3': 10.0,
    'Date Field4': 10.0,
    # Page 1 Section B
    'Text Field13': 4.825,
    'Text Field14': 8.5,
    'Text Field15': 9.0,
    'Text Field16': 9.0,
    'Text Field17': 6.725,
    'Text Field18': 4.825,
    'Text Field19': 9.0,
    'Text Field20': 9.0,
    'Text Field21': 4.825,
    # Page 2 Materials & Equipment
    'Text Field22': 4.35,
    'Text Field23': 4.75,
    'Text Field24': 10.5,
    'Text Field25': 8.0,
    'Text Field26': 9.0,
    'Text Field27': 9.5,
    'Text Field28': 9.5,
    'Text Field29': 9.5,
    'Text Field30': 9.5,
    'Text Field31': 9.5,
    'Text Field32': 4.4,
    'Text Field33': 4.4,
    'Text Field34': 9.5,
    'Text Field35': 9.5,
    'Text Field36': 9.5,
    'Text Field37': 9.5,
    'Text Field38': 9.5,
    'Text Field39': 9.5,
    'Text Field40': 7.5,
    'Text Field41': 7.5,
    'Text Field42': 7.5,
    'Text Field43': 8.5,
    'Text Field44': 8.5,
    'Text Field45': 7.5,
    # Page 3 Top
    'Text Field46': 9.0,
    'Text Field47': 9.0,
    'Text Field48': 8.5,
    # Pages 3 - 6 Narratives & Conclusions
    'Text Field49': 9.2,
    'Text Field50': 8.23,
    'Text Field51': 11.0,
    'Text Field52': 7.5,
    'Text Field53': 7.5,
    'Text Field54': 7.5,
    'Text Field55': 7.5,
}

def transcribe_to_pdf(template_path, output_path, append_tables_path=None):
    print(f"Transcribing into template: {template_path}")
    doc = fitz.open(template_path)
    
    for page in doc:
        for widget in page.widgets():
            fn = widget.field_name
            if fn in pdf_field_values:
                widget.field_value = pdf_field_values[fn]
                widget.text_font = "Helv"
                if fn in calibrated_fonts:
                    widget.text_fontsize = calibrated_fonts[fn]
                widget.update()
            elif fn in cb_whitelists:
                widget.field_value = cb_whitelists[fn]
                widget.update()
                
    temp_filled = os.path.join(SCRIPT_DIR, "temp_transcribed_filled.pdf")
    doc.save(temp_filled)
    doc.close()
    
    writer = PdfWriter(clone_from=temp_filled)
    if append_tables_path and os.path.exists(append_tables_path):
        t_reader = PdfReader(append_tables_path)
        for t_page in t_reader.pages:
            writer.add_page(t_page)
        print(f"Appended Table pages. Total pages: {len(writer.pages)}")
        
    with open(output_path, "wb") as f:
        writer.write(f)
    print(f"Successfully saved to: {output_path}")

if __name__ == "__main__":
    corp_form_desktop = os.path.join(DESKTOP_DIR, "CORP-FORM-21 - P1 - 16 Sep 2026.pdf")
    tables_pdf_desktop = os.path.join(DESKTOP_DIR, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
    master_pdf_desktop = os.path.join(DESKTOP_DIR, f"OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
    master_pdf_docs = os.path.join(DOCUMENTS_DIR, f"OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
    
    # 1. Fill CORP-FORM-21 - P1 - 16 Sep 2026.pdf directly on Desktop
    # We load from the pristine backup to ensure clean transcription
    backup_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\.history\CORP-FORM-21 - P1 - 16 Sep 2026_backup_20261009_114913.pdf"
    
    print("\n=== Transcribing into Desktop CORP-FORM-21 - P1 - 16 Sep 2026.pdf ===")
    transcribe_to_pdf(backup_path, corp_form_desktop, tables_pdf_desktop)
    
    print("\n=== Also updating Desktop & Documents Master OOS PDFs ===")
    transcribe_to_pdf(backup_path, master_pdf_desktop, tables_pdf_desktop)
    transcribe_to_pdf(backup_path, master_pdf_docs, tables_pdf_desktop)
    
    print("\nAll transcriptions completed successfully!")
