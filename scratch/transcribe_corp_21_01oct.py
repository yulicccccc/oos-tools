import os
import sys
import fitz
from pypdf import PdfWriter, PdfReader

src_pdf = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21- 01 Oct 2026.pdf"
desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docs_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"

# Ground truth values
OOS_ID = "262289"
CLIENT_NAME = "Link Pharmacy (E19112)"
SAMPLE_ID = "ETX-260929-0931"
SAMPLE_NAME = "Tirzepatide/Levocarnitine 17mg/100mg/mL (1ml)"
LOT_NUMBER = "20260928@54"
DOSAGE_FORM = "Injectable"
TEST_DATE = "01Oct26"
TEST_DATE_FULL = "01-Oct-2026"
TEST_RECORD = "100126-1040-2"

PREPPER_NAME = "Sasha Allen"
PREPPER_INIT = "SAA"
PROCESSOR_NAME = "Ellen Bennett"
PROCESSOR_INIT = "ELB"
CHANGEOVER_NAME = "Karla Silva"
CHANGEOVER_INIT = "KSM"
READER_NAME = "Varsha Subramanian"
READER_INIT = "VV"

BSC_PROC = "1938"
BSC_CHG = "1937"
SCAN_ID = "1040"
MONTHLY_CLEANING_DATE = "27Sep26"
EVENTS_COUNT = "15"
CONFIRMED_COUNT = "4"
MORPHOLOGY = "Cocci shaped morphology"

CONTROL_POSITIVE = "P. aeruginosa"
CONTROL_LOT = "05162026-9027-PA"
CONTROL_EXP = "15May28"

# Narrative paragraphs
p1 = (
    f"All analysts involved in the preparation, processing, changeover, and reading of the sample – "
    f"{PREPPER_NAME}, {PROCESSOR_NAME}, {CHANGEOVER_NAME}, and {READER_NAME} – were interviewed comprehensively. "
    "Their responses are documented throughout this report. This investigation was performed in "
    "accordance with MICRO-SOP-53, Sterility Test Out-of-Specification Investigation Procedure."
)

p2 = (
    "Upon arrival, the sample was stored refrigerated in accordance with the Client's instructions. "
    f"Prepping analyst - {PREPPER_NAME} - and processing analyst - {PROCESSOR_NAME} - verified the "
    "physical integrity and packaging of the sample throughout both the preparation and processing stages. "
    "No leaks, cracks, or visible turbidity were observed prior to testing, verifying container integrity."
)

p3 = (
    "All reagents and supplies mentioned in the material section above were stored according to suppliers’ "
    "recommendations, and their integrity was visually verified before utilization. Moreover, each reagent, "
    "media, and consumable possessed valid expiration dates. The functionality of all equipment was confirmed "
    "by reviewing continuous environmental and operational monitoring data."
)

p4 = (
    f"During the preparation phase, {PREPPER_NAME} disinfected the sample containers using acidified bleach "
    "and placed them into a pre-disinfected storage bin. On 01Oct26, prior to sample processing, "
    f"{PROCESSOR_NAME} performed a second disinfection with acidified bleach, allowing a minimum contact "
    "time of 10 minutes before transferring the samples into the cleanroom suites. A final disinfection step "
    "was completed immediately before introducing the sample into the ISO 5 Biological Safety Cabinet "
    f"(BSC E00{BSC_PROC}) located in the ISO 7 cleanroom (144). All activities were performed in "
    "accordance with MICRO-SOP-12, Rapid Scan RDI® Test using FIFU Method."
)

p5 = (
    "Cleanroom L-Suite, used for both sample processing and changeover, consists of ISO 7 cleanroom 145, "
    "ISO 7 buffer room 144, ISO 8 anteroom 143, and outermost ISO 8 room 142. Outward-cascading positive "
    "air pressure differentials are continuously maintained across the suite to ensure controlled "
    "unidirectional airflow."
)

p6 = (
    f"The ISO 5 BSC E00{BSC_PROC} and ISO 5 BSC E00{BSC_CHG} located in Cleanroom 144 (L-Suite) were thoroughly "
    "cleaned and disinfected prior to their respective procedures in accordance with MICRO-SOP-9, Cleaning and "
    "Disinfecting Procedure for Microbiology. Both BSCs were certified and approved by the Engineering and "
    "Quality Assurance teams prior to testing."
)

p7 = (
    f"Pre-labeling and filtration activities were performed by {PROCESSOR_NAME} on 01Oct26 within the ISO 5 "
    f"biosafety cabinet (BSC E00{BSC_PROC}) located in the ISO 7 cleanroom (144). The changeover procedure "
    f"was subsequently performed by analyst {CHANGEOVER_NAME} on the same date within the ISO 5 biosafety cabinet "
    f"(BSC E00{BSC_CHG}) located in Cleanroom 144 of L-Suite. The reading analyst, {READER_NAME}, confirmed that the "
    f"ScanRDI instrument E00{SCAN_ID} was set up as per ENG-SOP-4 (Scan RDI® System – Operations (Standard C3 Quality "
    f"Check and Microscope Setup and Maintenance)), and the negative control and positive control for analyst {PROCESSOR_NAME} "
    "yielded expected results."
)

p8 = (
    f"On 01Oct26, a rapid sterility test was performed on the sample using the ScanRDI method under test record {TEST_RECORD}. "
    f"The sample was prepared by analyst {PREPPER_NAME}, processed by {PROCESSOR_NAME} (39th sample processed), "
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
    "Weekly surface monitoring of Cleanroom L-Suite conducted on 01Oct26 demonstrated no microbial recovery across all "
    "sampled cleanroom surfaces, confirming effective sanitization and surface control."
)

p12 = (
    "Weekly active air monitoring of Cleanroom L-Suite conducted on 01Oct26 demonstrated no microbial recovery in ISO 7 "
    "cleanroom 145 and ISO 8 anteroom 143; however, 1 CFU under test sample ETX-261012-0430 was recovered from ISO 7 cleanroom "
    "144 Section II (identified as Staphylococcus epidermidis), and 5 CFUs under test sample ETX-261012-0433 were recovered from "
    "outermost ISO 8 room 142 (identified as Staphylococcus epidermidis, Pseudomonas parafulva, Bacillus safensis, Corynebacterium "
    "xerosis, and Corynebacterium ureicelerivorans). The recovered organisms are common environmental isolates with morphology "
    "distinctly different from the Cocci shaped morphology observed in the failing test sample, and no microbial match could be established."
)

p13 = (
    f"It is important to emphasize that all sample processing activities were performed strictly within the validated ISO 5 "
    f"BSC E00{BSC_PROC} and changeover activities within ISO 5 BSC E00{BSC_CHG}, both located in the ISO 7 cleanroom (144). "
    "There is minimal exposure of the test sample to ambient ISO 8 air, as samples and supplies are transferred in disinfected, "
    "closed containers on carts through the layered cleanroom suites. Furthermore, the recoveries occurred in the general cleanroom "
    "and lower-classified ISO 8 outer room environments, which are physically segregated from the ISO 5 processing zone by closed doors "
    "and an outward-cascading positive air pressure gradient. In conjunction with the complete absence of microbial growth on all "
    "analyst glove touch plates, settling plates, and ISO 5 BSC work surfaces throughout testing, there is no plausible contamination "
    "transfer pathway from the ambient environment to the test sample."
)

p14 = (
    f"Monthly cleaning and disinfection of the cleanrooms (ISO 7) and their containing biosafety cabinets (ISO 5) were performed "
    f"on {MONTHLY_CLEANING_DATE} for L-Suite, as per MICRO-SOP-9, Cleaning and Disinfecting Procedure for Microbiology. It was documented "
    "that all chemical indicators passed, verifying that environmental sanitization remained validated and in control prior to testing."
)

p15 = (
    f"Analyzing a 6-month sample history for {CLIENT_NAME}, this specific analyte \"{SAMPLE_NAME}\" has had no prior Out-of-Specification (OOS) occurrences under ScanRDI sterility testing."
)

p16 = (
    "All samples processed on the same day were also evaluated for evidence of sample-to-sample cross-contamination. "
    f"The test sample {SAMPLE_ID} was the 39th sample processed by {PROCESSOR_NAME} during the 2nd Shift on 01Oct26 "
    f"within BSC E00{BSC_PROC} (Cleanroom 144, L-Suite). All other samples processed by {PROCESSOR_NAME} during this shift "
    "yielded negative results. Review of laboratory records identified one other sample that yielded a positive result "
    "on 01Oct26, documented under OOS-262285 for CasaPharma Rx (E72585), sample ETX-260930-0635 "
    "(Tirzepatide/B12 15mg/1mg/0.5mL - 2mL). However, that sample was processed by "
    "America Alanis (11th sample processed) during the 1st Shift. The two testing events were segregated by shift, "
    f"processing analyst and primary containment equipment. Furthermore, analyst "
    f"{PROCESSOR_NAME} verified that gloves were disinfected between samples and standard aseptic techniques were strictly "
    "maintained. Consequently, there is no evidence of sample-to-sample cross-contamination or a shared laboratory vector."
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
text_field_51 = "N/A QYC 21Oct26"

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
    "Jun 2027\r \r"
    "Jul 2027"
)

personnel_block = (
    f"Prepping Analyst:\n{PREPPER_NAME} ({PREPPER_INIT})\n\n"
    f"Processing Analyst:\n{PROCESSOR_NAME} ({PROCESSOR_INIT})\n\n"
    f"Changeover Analyst:\n{CHANGEOVER_NAME} ({CHANGEOVER_INIT})\n\n"
    f"Reading Analyst:\n{READER_NAME} ({READER_INIT})"
)

doc = fitz.open(src_pdf)

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
    'Text Field0': f"{READER_NAME} (Written by: Qiyue Chen)",  # Initiator
    'Text Field1': "Scan RDI Sterility Test",
    'Text Field3': personnel_block,
    'Text Field7': f"On {TEST_DATE}, sample {SAMPLE_ID} was found positive for viable microorganisms after Scan RDI sterility testing.",
    'Text Field11': "<1 event",
    'Text Field12': "Kathan Parikh",
    'Text Field13': f"Yes, analysts {PREPPER_NAME}, {PROCESSOR_NAME}, {CHANGEOVER_NAME}, and {READER_NAME} were interviewed comprehensively.",
    'Text Field14': f"Yes, {SAMPLE_ID}",
    'Text Field17': "Yes, See 100126-1040-2 for more information.",
    'Text Field18': "Yes, all analysts are trained and qualified by quality to perform the test.",
    'Text Field19': "Not Applicable",
    'Text Field20': "Not Applicable",
    'Text Field21': f"Yes, Information is available in Eagle Trax Sample Location History under {SAMPLE_ID}",
    'Text Field22': (
        "Scan Consumables:\rSee the attached data packet\r \r"
        "Environmental Plates:\rTSA and Surface Plate: see attached environmental logs\r \r"
        "Monthly Cleaning:\rH2O2 strips and IPA: See attached monthly cleaning logs"
    ),
    'Text Field23': (
        "Scan Consumables:\rSee the attached data packet\r \r"
        "Environmental Plates:\rTSA and Surface Plate: see attached environmental logs\r \r"
        "Monthly Cleaning:\rH2O2 strips and IPA: See attached monthly cleaning logs"
    ),
    'Text Field24': CONTROL_POSITIVE,
    'Text Field25': CONTROL_LOT,
    'Text Field26': CONTROL_EXP,
    'Text Field27': "Not Applicable",
    'Text Field28': "Not Applicable",
    'Text Field29': "Not Applicable",
    'Text Field30': f"E00{SCAN_ID}",
    'Text Field31': "Nov 2026",
    'Text Field32': "E001978 (CR144)",
    'Text Field33': "Dec 2026",
    'Text Field34': f"E00{SCAN_ID}",
    'Text Field35': "Nov 2026",
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
    'Text Field48': "N/A QYC 21Oct26",
    'Text Field49': text_field_49,
    'Text Field50': text_field_50,
    'Text Field51': text_field_51,
    'Text Field53': "Qiyue Chen",
    'Text Field54': "",
    'Text Field55': "",
    'Text Field52': "",
}

cb_whitelists = {
    'Check Box0': 'Yes',
    'Check Box1': 'Yes',
    'Check Box2': 'Yes',
    'Check Box3': 'Off',
    'Check Box4': 'Yes', 'Check Box5': 'Off', 'Check Box6': 'Off',
    'Check Box7': 'Yes', 'Check Box8': 'Off', 'Check Box9': 'Off',
    'Check Box10': 'Yes', 'Check Box11': 'Off', 'Check Box12': 'Off',
    'Check Box13': 'Yes', 'Check Box14': 'Off', 'Check Box15': 'Off',
    'Check Box16': 'Yes', 'Check Box17': 'Off', 'Check Box18': 'Off',
    'Check Box19': 'Yes', 'Check Box20': 'Off', 'Check Box21': 'Off',
    'Check Box22': 'Off', 'Check Box23': 'Off', 'Check Box24': 'Yes',
    'Check Box25': 'Off', 'Check Box26': 'Off', 'Check Box27': 'Yes',
    'Check Box28': 'Yes', 'Check Box29': 'Off', 'Check Box30': 'Off',
    'Check Box31': 'Off', 'Check Box32': 'Yes', 'Check Box33': 'Off',
    'Check Box34': 'Yes', 'Check Box35': 'Off', 'Check Box36': 'Off',
    'Check Box37': 'Off', 'Check Box38': 'Yes', 'Check Box39': 'Off',
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
    'Check Box78': 'Yes',
    'Check Box79': 'Yes',
    'Check Box87': 'Off',
    'Check Box88': 'Yes',
    'Check Box89': 'Off',
    'Check Box90': 'Off',
    'Check Box91': 'Off',
    'Check Box92': 'Off',
}

robin_fonts = {
    'Date Field0': 10.0, 'Date Field1': 10.0, 'Date Field2': 10.0, 'Date Field3': 10.0,
    'Text Field0': 4.75, 'Text Field1': 7.5, 'Text Field2': 9.65, 'Text Field3': 6.0,
    'Text Field4': 8.5, 'Text Field5': 9.5, 'Text Field6': 9.5, 'Text Field7': 7.5,
    'Text Field8': 5.18, 'Text Field9': 5.18, 'Text Field10': 5.18, 'Text Field11': 9.5,
    'Text Field12': 10.0, 'Text Field13': 4.825, 'Text Field14': 8.5, 'Text Field15': 9.0,
    'Text Field16': 9.0, 'Text Field17': 6.725, 'Text Field18': 4.825, 'Text Field19': 9.0,
    'Text Field20': 9.0, 'Text Field21': 4.825, 'Text Field22': 6.0, 'Text Field23': 6.0,
    'Text Field24': 10.5, 'Text Field25': 8.0, 'Text Field26': 9.0, 'Text Field27': 9.5,
    'Text Field28': 9.5, 'Text Field29': 9.5, 'Text Field30': 9.5, 'Text Field31': 9.5,
    'Text Field32': 5.5, 'Text Field33': 5.5, 'Text Field34': 9.5, 'Text Field35': 9.5,
    'Text Field36': 9.5, 'Text Field37': 9.5, 'Text Field38': 9.5, 'Text Field39': 9.5,
    'Text Field40': 7.5, 'Text Field41': 7.5, 'Text Field42': 7.5, 'Text Field43': 8.5,
    'Text Field44': 8.5, 'Text Field45': 7.5, 'Text Field46': 9.0, 'Text Field47': 9.0,
    'Text Field48': 8.5, 'Text Field49': 8.2, 'Text Field50': 7.3, 'Text Field51': 11.0,
    'Text Field52': 7.5, 'Text Field53': 7.5, 'Text Field54': 7.5, 'Text Field55': 7.5,
    'Text Field57': 6.0,
}

for page in doc:
    for widget in page.widgets():
        fn = widget.field_name
        if fn in pdf_field_values:
            widget.field_value = pdf_field_values[fn]
            if fn in robin_fonts:
                widget.text_fontsize = robin_fonts[fn]
                widget.text_font = "Helv"
            widget.update()
        elif fn in cb_whitelists:
            widget.field_value = cb_whitelists[fn]
            widget.update()
        elif fn in robin_fonts:
            widget.text_fontsize = robin_fonts[fn]
            widget.update()

temp_filled_pdf = "scratch/temp_filled_corp21.pdf"
doc.save(temp_filled_pdf)
doc.close()

# Append Table 1 & Table 2
tables_pdf_desktop = os.path.join(desktop_dir, f"Tables OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
writer = PdfWriter(clone_from=temp_filled_pdf)
if os.path.exists(tables_pdf_desktop):
    t_reader = PdfReader(tables_pdf_desktop)
    for t_page in t_reader.pages:
        writer.add_page(t_page)
    print("Appended Table 1 & Table 2 as Page 7. Total pages:", len(writer.pages))

out_pdf_desktop = os.path.join(desktop_dir, f"OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
out_pdf_docs = os.path.join(docs_dir, f"OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
out_transcribed_desktop = os.path.join(desktop_dir, f"CORP-FORM-21- 01 Oct 2026 OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")
out_transcribed_docs = os.path.join(docs_dir, f"CORP-FORM-21- 01 Oct 2026 OOS-{OOS_ID} {CLIENT_NAME} - ScanRDI.pdf")

with open(out_pdf_desktop, "wb") as f:
    writer.write(f)
with open(out_transcribed_desktop, "wb") as f:
    writer.write(f)
print("Saved Master PDF to Desktop:", out_transcribed_desktop)

try:
    with open(out_pdf_docs, "wb") as f:
        writer.write(f)
    with open(out_transcribed_docs, "wb") as f:
        writer.write(f)
    print("Saved Master PDF to Documents:", out_transcribed_docs)
except Exception as e:
    print("Documents lock status:", e)

# Render preview pages
doc_preview = fitz.open(out_pdf_desktop)
out_dir = "scratch/corp21_rendered_pages"
os.makedirs(out_dir, exist_ok=True)
for i, page in enumerate(doc_preview):
    pix = page.get_pixmap(dpi=150)
    pix.save(os.path.join(out_dir, f"page_{i+1}.png"))
print("Saved rendered preview pages to scratch/corp21_rendered_pages")
