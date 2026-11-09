import os
import sys
import copy
import shutil
import re
import fitz
from pypdf import PdfReader, PdfWriter
import docx

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"

FORM9_PATH = os.path.join(DESKTOP_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1) (9).pdf")

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

# 4-Analyst roster for OOS-262414 (Nicholas Trinh prepped on 21Oct26 at 20:51 CST)
PREPPER_NAME = "Nicholas Trinh"
PREPPER_INIT = "NT"
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

CONTROL_POSITIVE = "A. brasiliensis"
CONTROL_LOT = "05222026-16404-AB"
CONTROL_EXP = "22May28"

# Narrative text
p1 = (
    f"All analysts involved in the preparation, processing, changeover, and reading of the sample — "
    f"{PREPPER_NAME}, {PROCESSOR_NAME}, and {READER_NAME} — were interviewed comprehensively. "
    "Their responses are documented throughout this report. This investigation was performed in "
    "accordance with MICRO-SOP-53, Sterility Test Out-of-Specification Investigation Procedure."
)

p2 = (
    "Upon arrival, the sample was stored refrigerated in accordance with the Client's instructions. "
    f"Prepping analyst {PREPPER_NAME} and processing analyst {PROCESSOR_NAME} verified the physical integrity "
    "and packaging of the sample throughout both the preparation and processing stages. No leaks, cracks, "
    "or visible turbidity were observed prior to testing, verifying container integrity."
)

p3 = (
    "All reagents and supplies mentioned in the material section above were stored according to suppliers' "
    "recommendations, and their integrity was visually verified before utilization. Moreover, each reagent, "
    "media, and consumable possessed valid expiration dates. The functionality of all equipment was confirmed "
    "by reviewing continuous environmental and operational monitoring data."
)

p4 = (
    f"During the preparation phase, {PREPPER_NAME} disinfected the sample containers using acidified bleach "
    "and placed them into a pre-disinfected storage bin. On 22Oct26, prior to sample processing, "
    f"{PROCESSOR_NAME} performed a second disinfection with acidified bleach, allowing a minimum contact "
    "time of 10 minutes before transferring the samples into the cleanroom suites. A final disinfection step "
    "was completed immediately before introducing the sample into the ISO 5 Biological Safety Cabinet "
    f"(BSC E00{BSC_PROC}) located in the ISO 7 cleanroom (145). All activities were performed in "
    "accordance with MICRO-SOP-12, Rapid Scan RDI Test using FIFO Method."
)

p5 = (
    "Cleanroom Parallel Suite, used for sample processing and changeover, consists of ISO 7 cleanroom 145, "
    "ISO 7 cleanroom 144, ISO 8 anteroom 143, and outermost ISO 8 room 142. Outward-cascading positive "
    "air pressure differentials are continuously maintained across the suite to ensure controlled "
    "unidirectional airflow."
)

p6 = (
    f"The ISO 5 BSC E00{BSC_PROC} and ISO 5 BSC E00{BSC_CHG} located in Cleanroom 145 and Cleanroom 144 (Parallel Suite) "
    "were thoroughly cleaned and disinfected prior to their respective procedures in accordance with "
    "MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. Both BSCs were certified and approved "
    "by the Engineering and Quality Assurance teams prior to testing."
)

p7 = (
    f"Pre-labeling and filtration activities were performed by {PROCESSOR_NAME} on 22Oct26 within the ISO 5 "
    f"biosafety cabinet (BSC E00{BSC_PROC}) located in the ISO 7 cleanroom (145). The changeover procedure "
    f"was subsequently performed by analyst {CHANGEOVER_NAME} on the same date within the ISO 5 biosafety cabinet "
    f"(BSC E00{BSC_CHG}) located in Cleanroom 144 of Parallel Suite. The reading analyst, {READER_NAME}, confirmed that the "
    f"ScanRDI instrument E00{SCAN_ID} was set up as per ENG-SOP-4 (Scan RDI System — Operations (Standard C3 Quality "
    f"Check and Microscope Setup and Maintenance)), and the negative control and positive control for analyst {PROCESSOR_NAME} "
    "yielded expected results."
)

# Paragraph 8: Clean narrative without redundant processing sequence per user rule
p8 = (
    f"On 22Oct26, a rapid sterility test was performed on the sample using the ScanRDI method under test record {TEST_RECORD} "
    f"(Analysis ID {ANALYSIS_ID}). The sample was prepped by {PREPPER_NAME}, processed by {PROCESSOR_NAME}, "
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
    "Weekly surface monitoring of Cleanroom Parallel Suite conducted on 20Oct26 demonstrated no microbial recovery across all "
    "sampled cleanroom surfaces, confirming effective sanitization and surface control."
)

p12 = (
    "Weekly active air monitoring of Cleanroom Suite conducted on 20Oct26 demonstrated no microbial recovery in ISO 7 "
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
    f"on {MONTHLY_CLEANING_DATE} for the suite, as per MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. It was documented "
    "that all chemical indicators passed, which confirmed the efficient monthly cleaning of all cleanrooms."
)

p15 = (
    f"Analyzing a 6-month sample history for {CLIENT_NAME}, this specific analyte \"{SAMPLE_NAME}\" has had no prior Out-of-Specification (OOS) occurrences under ScanRDI sterility testing."
)

# Paragraph 16: Confidentiality sanitized (no third-party client names) per user rule
p16 = (
    "All samples processed within the same batch and across the facility on the date of testing were also evaluated "
    "for evidence of sample-to-sample cross-contamination or shared laboratory vectors. The test sample "
    f"{SAMPLE_ID} was processed as the 12th sample by {PROCESSOR_NAME} on 22Oct26 within BSC E00{BSC_PROC} "
    "(Cleanroom 145, Parallel Suite). The sample was released for testing with \"No Method Suitability Test on File.\" "
    f"All other samples processed within the same testing batch by {PROCESSOR_NAME} yielded negative results for microbial growth. "
    "One other sterility failure occurred on the same testing date (ETX-261019-0832 under OOS-262410, Lot 101626-7B). "
    "However, that sample was processed by a completely different analyst (Halaina Smith, 5th sample processed) "
    "within a physically separate cleanroom (ISO 7 Cleanroom 144 of L-Suite) and BSC E001937, with changeover and reading "
    f"performed by Maryam Naeem. Furthermore, the two samples exhibited different cellular morphologies: cocci and diplococci "
    f"shaped morphology in ETX-261019-0832 under OOS-262410 versus rod-shaped morphology in {SAMPLE_ID} under OOS-262414. "
    f"Analyst {PROCESSOR_NAME} verified that gloves were sanitized between samples and standard aseptic techniques were "
    "strictly maintained throughout testing. Consequently, there is no evidence of sample-to-sample cross-contamination "
    "or a shared laboratory vector."
)

p17 = (
    "Based on the laboratory investigation, no assignable laboratory cause was identified involving the analyst, instrument, "
    "reagents, supplies, or monitored laboratory environment. Therefore, the likelihood of laboratory error contributing to this "
    "failure is considered minimal, and the original result is deemed valid."
)

text_field_49 = "\r \r".join([p1, p2, p3, p4, p5, p6, p7, p8, p9])

p11_header = "Environmental Monitoring (EM) Recoveries Review:"
p11_sub1 = f"ISO 5 BSC and Personnel Evaluation:\r{p10}"
p12_sub2 = f"Cleanroom EM Evaluation:\r{p11}"

text_field_50 = "\r \r".join([p11_header, p11_sub1, p12_sub2, p12, p13, p14, p15, p16, p17])
text_field_51 = "N/A QYC 09Nov26"

tf43_equipment = (
    "Incubator E001034\r \r"
    "(Sensor E001501)\r \r"
    "Incubator E001031 \r \r"
    "(Sensor E001505)\r \r"
    "Incubator E001933\r \r"
    "Incubator E001932\r \r"
    "Incubator E001033 \r \r"
    "(Sensor E001491)\r \r"
    "BSC E001312\r \r"
    "BSC E001937"
)
tf44_cal_due = (
    "Aug 2027\r \r"
    "Feb 2027 \r \r"
    "Aug 2027\r \r"
    "Feb 2027\r \r"
    "Mar 2027\r \r"
    "Mar 2027 \r \r"
    "Jul 2027 \r \r"
    "Feb 2027\r \r"
    "Jun 2027\r \r"
    "Jul 2027"
)

tf22_materials = (
    "Scan Consumables:\rSee the attached data packet\r \r"
    "Environmental Plates:\rTSA and Surface Plate: see attached environmental logs\r \r"
    "Monthly Cleaning:\rH2O2 strips and IPA: See attached monthly cleaning logs"
)

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
    'Text Field0': f"{READER_NAME} (Written by: Qiyue Chen)",
    'Text Field1': "Scan RDI Sterility Test",
    'Text Field3': personnel_block,
    'Text Field7': f"On {TEST_DATE}, sample {SAMPLE_ID} was found positive for viable microorganisms after Scan RDI sterility testing.",
    'Text Field11': "<1 event",
    'Text Field12': "Kathan Parikh",
    'Text Field13': f"Yes, analysts {PREPPER_NAME}, {PROCESSOR_NAME}, and {READER_NAME} were interviewed comprehensively.",
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
    'Text Field32': "CR145 (E001979)\rCR144 (E001978)",
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
    'Text Field46': "No assignable cause",
    'Text Field47': "No assignable cause",
    'Text Field48': "N/A QYC 09Nov26",
    'Text Field49': text_field_49,
    'Text Field50': text_field_50,
    'Text Field51': text_field_51,
    'Text Field53': "Qiyue Chen",
    'Text Field54': "",
}

cb_whitelists = {
    'Check Box0': 'Yes', 'Check Box1': 'Yes', 'Check Box2': 'Yes',
    'Check Box4': 'Yes', 'Check Box7': 'Yes', 'Check Box10': 'Yes',
    'Check Box13': 'Yes', 'Check Box16': 'Yes', 'Check Box19': 'Yes',
    'Check Box24': 'Yes', 'Check Box27': 'Yes', 'Check Box28': 'Yes',
    'Check Box31': 'Off', 'Check Box32': 'Yes', 'Check Box33': 'Off',
    'Check Box34': 'Yes', 'Check Box35': 'Off', 'Check Box36': 'Off',
    'Check Box37': 'Yes', 'Check Box38': 'Off', 'Check Box39': 'Off',
    'Check Box42': 'Yes', 'Check Box43': 'Yes', 'Check Box48': 'Yes',
    'Check Box51': 'Yes', 'Check Box52': 'Yes', 'Check Box55': 'Yes',
    'Check Box58': 'Yes', 'Check Box63': 'Yes', 'Check Box66': 'Yes',
    'Check Box67': 'Yes', 'Check Box70': 'Yes', 'Check Box73': 'Yes',
    'Check Box78': 'Yes', 'Check Box79': 'Yes', 'Check Box80': 'Off', 'Check Box84': 'Off',
    'Check Box90': 'Off', 'Check Box91': 'Off', 'Check Box92': 'Off',
}

calibrated_font_sizes = {
    'Date Field0': 10.0,
    'Date Field1': 10.0,
    'Date Field2': 10.0,
    'Date Field3': 10.0,
    'Text Field1': 7.5,
    'Text Field2': 11.0,
    'Text Field3': 6.0,
    'Text Field4': 7.5,
    'Text Field5': 10.0,
    'Text Field6': 10.0,
    'Text Field7': 7.0,
    'Text Field8': 5.175,
    'Text Field9': 5.175,
    'Text Field10': 5.175,
    'Text Field11': 9.5,
    'Text Field12': 9.0,
    'Text Field13': 4.5,
    'Text Field14': 8.0,
    'Text Field15': 6.0,
    'Text Field16': 6.0,
    'Text Field17': 6.5,
    'Text Field18': 5.0,
    'Text Field19': 8.0,
    'Text Field20': 8.0,
    'Text Field21': 5.0,
    'Text Field22': 5.2,
    'Text Field23': 5.2,
    'Text Field24': 8.5,
    'Text Field25': 8.0,
    'Text Field26': 8.5,
    'Text Field27': 8.5,
    'Text Field28': 8.5,
    'Text Field29': 8.5,
    'Text Field30': 8.0,
    'Text Field31': 8.0,
    'Text Field32': 4.4,
    'Text Field33': 4.4,
    'Text Field34': 8.0,
    'Text Field35': 8.0,
    'Text Field36': 8.0,
    'Text Field37': 8.0,
    'Text Field38': 8.0,
    'Text Field39': 8.0,
    'Text Field40': 7.5,
    'Text Field41': 7.5,
    'Text Field42': 7.5,
    'Text Field43': 5.4,
    'Text Field44': 5.4,
    'Text Field45': 7.5,
    'Text Field46': 7.0,
    'Text Field47': 7.0,
    'Text Field48': 7.0,
    'Text Field49': 4.825,
    'Text Field50': 4.825,
    'Text Field51': 8.0,
    'Text Field53': 20.3,
    'Text Field57': 8.5,
}

print("=== 1. Filling 6 Pages of Authentic ZenQMS Template ===")
assert os.path.exists(FORM9_PATH), f"Form (9) template missing at: {FORM9_PATH}"
doc = fitz.open(FORM9_PATH)

for pno in range(min(6, len(doc))):
    page = doc[pno]
    for w in page.widgets():
        fname = w.field_name
        ftype = w.field_type_string

        if ftype == 'CheckBox':
            if fname in cb_whitelists:
                w.field_value = cb_whitelists[fname]
                w.update()
            else:
                w.field_value = 'Off'
                w.update()

        elif ftype == 'Text':
            if fname in pdf_field_values:
                val = pdf_field_values[fname]
                w.field_value = str(val)
                if fname in calibrated_font_sizes:
                    w.text_fontsize = calibrated_font_sizes[fname]
                w.update()

temp_6page = os.path.join(SCRIPT_DIR, f"temp_6page_{OOS_ID}.pdf")
doc.save(temp_6page)
doc.close()
print("Saved filled 6-page PDF to:", temp_6page)

print("=== 2. Appending Standalone Table Page (Page 7) ===")
writer = PdfWriter()

# Read 6 pages from filled form
reader_6p = PdfReader(temp_6page)
for page in reader_6p.pages:
    writer.add_page(page)

# Read Page 1 from standalone table PDF
table_pdf_path = os.path.join(SCRIPT_DIR, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
assert os.path.exists(table_pdf_path), f"Standalone table PDF missing: {table_pdf_path}"
reader_tbl = PdfReader(table_pdf_path)
writer.add_page(reader_tbl.pages[0])

final_master_pdf = os.path.join(ROOT_DIR, f"CORP-FORM-21- 22 Oct 2026 OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
with open(final_master_pdf, "wb") as f_out:
    writer.write(f_out)

print("Master 7-Page PDF generated:", final_master_pdf)

# 3. Verify Page Count & Zero-Overflow
chk_doc = fitz.open(final_master_pdf)
print(f"Verified Final Master PDF Page Count: {len(chk_doc)} (Strict requirement: 7)")
assert len(chk_doc) == 7, f"ERROR: Master PDF must be 7 pages, got {len(chk_doc)}"

overflow_count = 0
for p_no in range(6):
    p = chk_doc[p_no]
    for w in p.widgets():
        if w.field_type_string == 'Text' and w.field_value:
            xref = w.xref
            ap_xref = chk_doc.xref_get_key(xref, 'AP/N')
            if ap_xref[0] == 'xref':
                sx = int(ap_xref[1].split()[0])
                stream = chk_doc.xref_stream(sx).decode('latin1', errors='ignore')
                re_match = re.search(r'([0-9\.\-]+)\s+([0-9\.\-]+)\s+([0-9\.\-]+)\s+([0-9\.\-]+)\s+re\s+W', stream)
                if re_match:
                    rw = float(re_match.group(3))
                    rh = float(re_match.group(4))
                    td_matches = re.findall(r'([0-9\.\-]+)\s+([0-9\.\-]+)\s+Td', stream)
                    cur_y = 0.0
                    min_y = 999.0
                    for td in td_matches:
                        cur_y += float(td[1])
                        if cur_y < min_y:
                            min_y = cur_y
                    if min_y < 0:
                        print(f"  [ERROR] OVERFLOW on Page {p_no+1} {w.field_name}: min_y={min_y:.2f}, rh={rh:.2f}, text={repr(w.field_value[:35])}")
                        overflow_count += 1

print(f"Zero-Overflow Verification: {overflow_count} overflowing fields detected.")
assert overflow_count == 0, f"Detected {overflow_count} overflowing fields!"

# 4. Sync to Desktop per Rule 0M
dst_desktop = os.path.join(DESKTOP_DIR, os.path.basename(final_master_pdf))
try:
    shutil.copy2(final_master_pdf, dst_desktop)
    print("Master PDF synced to Desktop:", dst_desktop)
except Exception as e:
    print(f"Direct overwrite of {dst_desktop} failed (locked by PDF reader): {e}")
    dst_desktop_v = os.path.join(DESKTOP_DIR, f"CORP-FORM-21- 22 Oct 2026 OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI (v12.4.28).pdf")
    shutil.copy2(final_master_pdf, dst_desktop_v)
    print("Master PDF synced to versioned Desktop copy:", dst_desktop_v)

# 5. Render preview images of Page 1 and Page 2
preview_dir = os.path.join(SCRIPT_DIR, f"preview_oos_{OOS_ID}")
os.makedirs(preview_dir, exist_ok=True)
for pno in range(min(2, len(chk_doc))):
    page = chk_doc[pno]
    pix = page.get_pixmap(dpi=200)
    out_img = os.path.join(preview_dir, f"page_{pno+1}_v12428.png")
    pix.save(out_img)
    print(f"Saved preview: {out_img}")

print(f"=== Rebuild OOS-{OOS_ID} Complete! ===")
