import os
import shutil
import datetime
import fitz
import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
import win32com.client

WORKSPACE = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
DESKTOP = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"
HISTORY_DIR = os.path.join(WORKSPACE, ".history")
os.makedirs(HISTORY_DIR, exist_ok=True)

ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

# 0. Backup template and existing deliverables to .history/
TEMPLATE_DOCX = os.path.join(WORKSPACE, "scratch", "Celsis table OOS-262080.docx")
if os.path.exists(TEMPLATE_DOCX):
    shutil.copy2(TEMPLATE_DOCX, os.path.join(HISTORY_DIR, f"Celsis table OOS-262080_backup_{ts}.docx"))

DOCX_NAME = "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.docx"
PDF_NAME = "CORP-FORM-21- 12 Oct 2026 OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis.pdf"

for folder in [DESKTOP, DOCS_DIR]:
    for fname in [DOCX_NAME, PDF_NAME, "Tables OOS-262333 Optimal Balance Pharmacy (E19193) - Celsis (Draft).docx"]:
        fpath = os.path.join(folder, fname)
        if os.path.exists(fpath):
            shutil.copy2(fpath, os.path.join(HISTORY_DIR, f"{fname}_backup_{ts}"))
            if "Draft" in fname:
                try:
                    os.remove(fpath)
                except Exception:
                    pass

def set_cell_text_preserve_format(cell, text, italic=False, bold=False, size_pt=7.0):
    """Updates cell text while preserving cell borders, shading, width, and vertical alignment."""
    # Keep first paragraph, remove any extra paragraphs
    while len(cell.paragraphs) > 1:
        p_el = cell.paragraphs[-1]._p
        p_el.getparent().remove(p_el)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        p._p.remove(r._r)
    # Remove any existing hyperlinks inside paragraph
    for hl in list(p._p.findall(qn("w:hyperlink"))):
        p._p.remove(hl)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.italic = italic
    run.bold = bold

def set_Para_underlined_title(p, text):
    for r in list(p.runs):
        p._p.remove(r._r)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(9.0)
    run.underline = True
    run.bold = False

def build_tables_from_template(out_docx_path):
    doc = docx.Document(TEMPLATE_DOCX)

    # 1. Update the 3 table title paragraphs in-place
    set_Para_underlined_title(doc.paragraphs[2], "Table 1: Information for ETX-261002-0288 under investigation")
    set_Para_underlined_title(doc.paragraphs[6], "Table 2: Environmental Monitoring from Processing Performed on 05Oct26")
    set_Para_underlined_title(doc.paragraphs[10], "Table 3: Environmental Monitoring from Aliquoting Performed on 12Oct26")

    # 2. Update hyperlink targets in document relationships
    # rId7 -> Sample ID ETX-261002-0288 in Table 1
    # rId8 -> Active Air Plate ETX-261021-0578 in Table 3
    doc.part.rels["rId7"]._target = "https://etrax.eagleanalytical.com/SubmissionTest/Details/5Fwh6gV%24HhaU8JvbbaL0nQ__"
    doc.part.rels["rId8"]._target = "https://etrax.eagleanalytical.com/SubmissionTest/Details/Cw%24QZhYAogGUwoutnOavpw__"

    # 3. Update Table 0 (Table 1: Information for ETX-261002-0288 under investigation)
    t0 = doc.tables[0]
    r1_cells = [docx.table._Cell(tc, t0) for tc in t0.rows[1]._tr.tc_lst]
    set_cell_text_preserve_format(r1_cells[0], "Alex Saravia")
    set_cell_text_preserve_format(r1_cells[1], "Cuong Du")
    # Update hyperlink text in r1_cells[2]
    hl_t0 = r1_cells[2].paragraphs[0]._p.find(qn("w:hyperlink"))
    if hl_t0 is not None:
        for t_el in hl_t0.findall(".//" + qn("w:t")):
            t_el.text = "ETX-261002-0288"
    set_cell_text_preserve_format(r1_cells[3], "ETX-261012-0654")
    set_cell_text_preserve_format(r1_cells[4], "1 x 300mL FTM")
    set_cell_text_preserve_format(r1_cells[5], "Bacillus paralicheniformis", italic=True)

    # 4. Update Table 1 (Table 2: Environmental Monitoring from Processing Performed on 05Oct26)
    t1 = doc.tables[1]
    t1_updates = {
        2: ["Personnel (Left Touch and Right Touch)", "Daily", "02Oct26", "ES", "Date Before Testing", "No growth", "N/A", "N/A", "None"],
        3: ["Personnel (Left Touch and Right Touch)", "Daily", "05Oct26", "ES", "Date of Testing", "No growth", "N/A", "N/A", "None"],
        4: ["Personnel (Left Touch and Right Touch)", "Daily", "06Oct26", "ES", "Date After Testing", "No growth", "N/A", "N/A", "None"],
        6: ["Surface Sampling of ISO 5 (4 locations)", "Daily", "02Oct26", "ALA", "Date Before Testing", "No growth", "N/A", "N/A", "None"],
        7: ["Surface Sampling of ISO 5 (4 locations)", "Daily", "05Oct26", "ES", "Date of Testing", "No growth", "N/A", "N/A", "None"],
        8: ["Surface Sampling of ISO 5 (4 locations)", "Daily", "06Oct26", "ES", "Date After Testing", "No growth", "N/A", "N/A", "None"],
        9: ["Settling Sampling of ISO 5 (2 locations)", "Daily", "02Oct26", "ALA", "Date Before Testing", "No growth", "N/A", "N/A", "None"],
        10: ["Settling Sampling of ISO 5 (2 locations)", "Daily", "05Oct26", "ES", "Date of Testing", "No growth", "N/A", "N/A", "None"],
        11: ["Settling Sampling of ISO 5 (2 locations)", "Daily", "06Oct26", "ES", "Date After Testing", "No growth", "N/A", "N/A", "None"],
        13: ["Active Air Sampling of Cleanrooms", "Weekly", "05Oct26", "SMO", "Week of Testing", "No growth", "N/A", "N/A", "None"],
        15: ["Surface Sampling of Cleanrooms", "Weekly", "05Oct26", "SMO", "Week of Testing", "No growth", "N/A", "N/A", "None"],
    }
    for row_idx, vals in t1_updates.items():
        rcells = [docx.table._Cell(tc, t1) for tc in t1.rows[row_idx]._tr.tc_lst]
        for ci, val in enumerate(vals):
            set_cell_text_preserve_format(rcells[ci], val)

    # 5. Update Table 2 (Table 3: Environmental Monitoring from Aliquoting Performed on 12Oct26)
    t2 = doc.tables[2]
    t2_updates = {
        2: ["Personnel (Left Touch and Right Touch)", "Daily", "09Oct26", "CCD", "Date Before Aliquoting", "No growth", "N/A", "N/A", "None"],
        3: ["Personnel (Left Touch and Right Touch)", "Daily", "12Oct26", "CCD", "Date of Aliquoting", "No growth", "N/A", "N/A", "None"],
        4: ["Personnel (Left Touch and Right Touch)", "Daily", "13Oct26", "CCD", "Date After Aliquoting", "No growth", "N/A", "N/A", "None"],
        6: ["Surface Sampling of ISO 5 (4 locations)", "Daily", "09Oct26", "ALA", "Date Before Aliquoting", "No growth", "N/A", "N/A", "None"],
        7: ["Surface Sampling of ISO 5 (4 locations)", "Daily", "12Oct26", "CCD", "Date of Aliquoting", "No growth", "N/A", "N/A", "None"],
        8: ["Surface Sampling of ISO 5 (4 locations)", "Daily", "13Oct26", "CCD", "Date After Aliquoting", "No growth", "N/A", "N/A", "None"],
        9: ["Settling Sampling of ISO 5 (2 locations)", "Daily", "09Oct26", "ALA", "Date Before Aliquoting", "No growth", "N/A", "N/A", "None"],
        10: ["Settling Sampling of ISO 5 (2 locations)", "Daily", "12Oct26", "CCD", "Date of Aliquoting", "No growth", "N/A", "N/A", "None"],
        11: ["Settling Sampling of ISO 5 (2 locations)", "Daily", "13Oct26", "CCD", "Date After Aliquoting", "No growth", "N/A", "N/A", "None"],
        15: ["Surface Sampling of Cleanrooms", "Weekly", "12Oct26", "SMO", "Week of Testing", "No growth", "N/A", "N/A", "None"],
    }
    for row_idx, vals in t2_updates.items():
        rcells = [docx.table._Cell(tc, t2) for tc in t2.rows[row_idx]._tr.tc_lst]
        for ci, val in enumerate(vals):
            set_cell_text_preserve_format(rcells[ci], val)

    # Row 13 of Table 3 (Active Air Sampling of Cleanrooms with 6 CFUs and 7 Nanopore species)
    r13_cells = [docx.table._Cell(tc, t2) for tc in t2.rows[13]._tr.tc_lst]
    set_cell_text_preserve_format(r13_cells[0], "Active Air Sampling of Cleanrooms")
    set_cell_text_preserve_format(r13_cells[1], "Weekly")
    set_cell_text_preserve_format(r13_cells[2], "12Oct26")
    set_cell_text_preserve_format(r13_cells[3], "SMO")
    set_cell_text_preserve_format(r13_cells[4], "Week of Testing")
    set_cell_text_preserve_format(r13_cells[5], "6 CFUs (ISO 8 114)")
    # Update hyperlink text in r13_cells[6]
    hl_t2 = r13_cells[6].paragraphs[0]._p.find(qn("w:hyperlink"))
    if hl_t2 is not None:
        for t_el in hl_t2.findall(".//" + qn("w:t")):
            t_el.text = "ETX-261021-0578"
    # Populate 7 italicized species separated by <w:br/><w:br/> inside a single paragraph (matching OOS-262080 XML)
    species_list = [
        "Moraxella osloensis",
        "Sporisorium graminicola",
        "Staphylococcus haemolyticus",
        "Staphylococcus hominis",
        "Kocuria indica",
        "Ustilago maydis",
        "Kocuria rhizophila/ sp. BT304",
    ]
    p_sp = r13_cells[7].paragraphs[0]
    for r in list(p_sp.runs):
        p_sp._p.remove(r._r)
    for idx, sp in enumerate(species_list):
        run = p_sp.add_run(sp)
        run.font.name = "Times New Roman"
        run.font.size = Pt(6.5)
        run.italic = True
        if idx < len(species_list) - 1:
            run._r.append(parse_xml(f"<w:br {nsdecls('w')}/>"))
            run._r.append(parse_xml(f"<w:br {nsdecls('w')}/>"))
    set_cell_text_preserve_format(r13_cells[8], "None")

    doc.save(out_docx_path)

def convert_docx_to_pdf(docx_path, pdf_path):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close(False)
    finally:
        word.Quit()

def fill_form_and_merge(tables_pdf_path, final_pdf_path):
    blank_form = os.path.join(WORKSPACE, "scratch", "qa_celsis_emails", "262333_CORP-FORM-21 Laboratory OOS Investigation Form (v11.1) (5).pdf")
    doc = fitz.open(blank_form)

    p3_narrative = (
        "All analysts involved in the preparation (Mukyung Jang), processing (Alex Saravia), aliquoting (Cuong Du), and "
        "reading of the samples (Cuong Du) were interviewed comprehensively. Their responses are documented throughout "
        "this report. This investigation was performed as per MICRO-SOP-53 Sterility Test Out-of-Specification (OOS) Investigation "
        "Procedure.\r \r"
        "Upon arrival, the sample was stored refrigerated in accordance with the client instructions. Both sample prepping analyst "
        "Mukyung Jang and processing analyst (Alex Saravia) inspected the sample during preparation and processing. No "
        "leaks or visible turbidity were observed, and no evidence of compromised container integrity was identified.\r \r"
        "All reagents and supplies mentioned in the material section above were stored according to supplier recommendations, "
        "and their integrity was visually verified before utilization. Moreover, all reagents and supplies had valid expiration dates. "
        "The functionality of all equipment was confirmed by reviewing data generated by our comprehensive in-house continuous "
        "monitoring system.\r \r"
        "The cleanroom suite used for processing procedures (CR114) comprises three interconnected sections: the innermost "
        "ISO 7 cleanroom (114B), which connects to the middle ISO 7 buffer room (114A), and then to the outermost ISO 8 "
        "anteroom (114). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional "
        "airflow from 114B through 114A and into 114.\r \r"
        "The innermost ISO 7 cleanroom 114B and its ISO 5 BSC E001316 were thoroughly cleaned and prepared before initiating "
        "testing by analyst Alex Saravia as per MICRO-SOP-2 (Environmental Monitoring of the Clean Room Facility) and "
        "MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology). For Celsis aliquoting, intermediate ISO 7 buffer "
        "room 114A and its ISO 5 BSC E001798 were thoroughly cleaned and prepared by analyst (Cuong Du) before initiating "
        "testing as per MICRO-SOP-44 and MICRO-SOP-9. Both BSCs, E001316 in 114B and E001798 in 114A, were certified "
        "and approved by the Engineering and Quality Assurance teams prior to use.\r \r"
        "On 05 Oct 2026, the sample vials for ETX-261002-0288 were received from the Sample Submissions team and brought "
        "into the Sterile Microbiology lab. Upon arrival, each vial was sprayed with an acidified bleach disinfectant, placed into pre-"
        "disinfected bins, and allowed a 10-minute contact time. Secondary disinfection was performed in the ISO 8 anteroom "
        "(114), where the vials were again treated with acidified bleach with a 10-minute contact time before transfer into "
        "cleanroom 114B. Inside 114B, the processing analyst (Alex Saravia) performed a final disinfection step with a 10-"
        "minute contact time. Once fully disinfected, the vials were transferred into the ISO 5 BSC E001316, placed on the "
        "disinfected work surface, aseptically opened, and tested in accordance with MICRO-SOP-44. Direct inoculation was "
        "performed by adding 15 mL of sample per media type, using one 300 mL FTM jar and one 300 mL TSB jar. In "
        "accordance with the method suitability (ETX-250826-0390), 3 mL of sterile Tween 80 was added to each jar. "
        "Accordingly, a negative control with the same modifications was made. Following testing, the media bottles were "
        "transferred into incubators E001356 and E001357 to initiate incubation."
    )

    p4_narrative = (
        "Upon completion of incubation on 12 Oct 2026, the media bottles for ETX-261002-0288 were disinfected and transferred "
        "to the middle ISO 7 buffer room (114A) for the aliquoting step per MICRO-SOP-44 (Celsis Sterility Testing). In cleanroom "
        "114A, the media bottles were disinfected one more time before transfer into the ISO 5 BSC E001798 located in cleanroom "
        "114A. In ISO 5 BSC E001798, the sample was aliquoted into assay cuvettes by aliquoting analyst (Cuong Du). Following "
        "aliquoting, Celsis Sterility Reading was performed in accordance with MICRO-SOP-44 by reading analyst (Cuong Du).\r \r"
        "Following the reading, sample ETX-261002-0288 was found to yield a positive reading in the 300 mL FTM media "
        "jar. The confirmed average Relative Luminescence Units (RLU) from the duplicate reading tubes, "
        "originating from the positive FTM sample bottle, yielded 3,647,568 RLU, which "
        "exceeded the positive cutoff of 2,596.5 RLU, where the FTM negative control was 866 RLU. All other FTM and TSB sample "
        "bottles in the same batch tested negative. The %CV from the duplicate reading tubes for the positive FTM bottle was well "
        "within the specification (< 30%). Additionally, all Daily Controls, including the Instrument Blank, Reagent Blank, and ATP "
        "Positive Control, were within the defined specifications, each with a %CV below 30%.\r \r"
        "Following the OOS result, the positive FTM bottle for ETX-261002-0288 was submitted for Differential Staining and "
        "Microbial Identification under ETX-261012-0654. Microbial growth was recovered by subculture as shown in Table 1, "
        "and was identified as Bacillus paralicheniformis (Gram-positive rod).\r \r"
        "The culture media utilized were within their expiry period. The negative culture media bottles for the direct inoculation "
        "method for the original culture were handled, processed, and incubated in a manner identical to that of actual samples. No "
        "microbial growth was observed in the corresponding negative control.\r \r"
        "Monthly cleaning and disinfection of the outermost ISO 8 Anteroom, the middle ISO 7 Buffer room, the innermost ISO 7 "
        "cleanroom, and its containing ISO 5 Biosafety Cabinets for CR114 was performed on 27 Sep 2026, as per MICRO-SOP-9 "
        "(Cleaning and Disinfecting Procedure for Microbiology). It was documented that all H2O2 indicators passed. This confirms "
        "the efficient monthly cleaning of all three parts of Cleanroom 114.\r \r"
        "Attached Tables 2 and 3 present the environmental monitoring results for the duration of testing. The environmental "
        "monitoring (EM) plates were incubated for no less than 48 hours at 30-35\u00b0C and no less than an additional five days at "
        "20-25\u00b0C as per MICRO-SOP-2 (Environmental Monitoring of the Cleanroom Facility). Table 2 pertains to Environmental "
        "Monitoring performed during Celsis Sterility Processing, and Table 3 pertains to Environmental Monitoring performed "
        "during Celsis Sterility Aliquoting.\r \r"
        "A review of the Environmental Monitoring (EM) results for the relevant Celsis Sterility processing period showed no "
        "microbial growth on personnel monitoring plates, settling plates, or ISO 5 BSC E001316 surface sampling plates on the "
        "date of testing (05Oct2026), the preceding sampling date (02Oct2026), or the subsequent sampling date (06Oct2026). "
        "Weekly active air and surface sampling of the cleanroom suite associated with processing was performed on 05Oct2026. "
        "No microbial recovery was observed from weekly active air or surface sampling of the cleanroom suite."
    )

    p5_narrative = (
        "After reviewing the Environmental Monitoring (EM) results for the relevant Celsis Sterility aliquoting period, no microbial "
        "growth was observed on personnel monitoring plates, ISO 5 BSC E001798 surface sampling plates, or settling plates on "
        "the date of aliquoting (12Oct2026), the preceding sampling date (09Oct2026), or the subsequent sampling date "
        "(13Oct2026). Weekly active air sampling of the cleanroom suite performed on 12Oct2026 showed recovery of 6 CFUs "
        "(ETX-261021-0578) in the ISO 8 area, identified as Moraxella osloensis, Sporisorium graminicola, Staphylococcus "
        "haemolyticus, Staphylococcus hominis, Kocuria indica, Ustilago maydis, and Kocuria rhizophila/ sp. BT304, with zero "
        "recovery in the ISO 7 buffer or cleanrooms, and weekly surface sampling of the cleanroom suite showed no growth.\r \r"
        "Evaluation of Environmental Monitoring Results:\r \r"
        "It is important to note that, for both processing and aliquoting activities, the samples are processed within validated ISO 5 "
        "BSCs (E001316 and E001798 respectively). Although microbial recovery was observed from weekly active air monitoring "
        "in the outermost ISO 8 room 114 background environment in cleanroom suite (CR114), these recoveries were not "
        "obtained from the ISO 5 processing environment or the ISO 7 cleanroom environment.\r \r"
        "The absence of microbial recovery from both the processing and the aliquoting analyst glove plates, ISO 5 BSC surface "
        "monitoring plates, and ISO 5 settling plates indicates that no identifiable contamination transfer pathway existed between "
        "the monitored background cleanroom environment and the ISO 5 critical processing areas. Samples were also prepared "
        "and transferred into disinfected, lidded containers and transported on carts within the controlled cleanroom environment. "
        "Therefore, while the active air results demonstrate microbial recovery in the surrounding ISO 8 area, there is no evidence "
        "that the recovery originated from, or affected, the ISO 5 BSC environment where the test sample was handled and "
        "processed.\r \r"
        "All analysts confirmed full compliance with cleaning procedures as outlined in MICRO-SOP-9 (Cleaning and Disinfecting "
        "Procedure for Microbiology) and MICRO-SOP-44 (Celsis Sterility Testing). A review of the available data confirms that the "
        "cleanroom and equipment conditions remained within acceptable parameters. No deviations or obvious signs of laboratory "
        "error during processing and aliquoting were noted from analysts.\r \r"
        "No other sample processed by analyst Alex Saravia on 05 Oct 2026 and aliquoted by analyst Cuong Du on 12 Oct "
        "2026, failed Celsis Sterility testing that day. To assess the potential for sample-to-sample contamination contributing to the "
        "positive results, a comprehensive review was conducted of all samples processed in the same batch. All other samples "
        "processed by the same analyst on the day of testing were found to test negative for microbial growth. These findings "
        "suggest that cross-contamination between samples is unlikely.\r \r"
        'Analyzing a 6-month sample history for Optimal Balance Pharmacy indicates, this specific analyte "TESAMORELIN 8 MG/ML (3 '
        'ML) Injection" has had no prior failures using the Celsis Sterility testing during this period.\r \r'
        "Based on the laboratory investigation, no assignable laboratory cause was identified involving the analyst, instrument, "
        "reagents, supplies, or monitored laboratory environment. The initial Celsis OOS result therefore remains valid in "
        "accordance with the applicable laboratory OOS procedure."
    )

    # Exact field values and font sizes from OOS-262080
    field_specs = {
        "Text Field57": ("262333", 6.5),
        "Text Field58": ("262333", 6.5),
        "Text Field59": ("262333", 6.5),
        "Text Field60": ("262333", 6.5),
        "Text Field61": ("262333", 6.5),
        "Text Field62": ("262333", 6.5),
        "Text Field0": ("Cuong Du (Written by Qiyue Chen)", 5.5),
        "Date Field0": ("05-Oct-2026", 10.0),
        "Date Field1": ("12-Oct-2026", 10.0),
        "Date Field2": ("12-Oct-2026", 10.0),
        "Text Field1": ("Celsis Sterility Test", 8.0),
        "Text Field2": ("ETX-261002-0288", 10.0),
        "Text Field3": (
            "Prepping Analyst:\rMukyung Jang (MJ)\r \r"
            "Processing Analyst: \rAlex Saravia (ES)\r \r"
            "Aliquoting Analyst:\rCuong Du (CCD)\r \r"
            "Reading Analyst:\rCuong Du (CCD)",
            5.8,
        ),
        "Text Field4": ("TESAMORELIN 8 MG/ML (3 ML) Injection", 8.0),
        "Text Field5": ("Injection", 9.5),
        "Text Field6": ("LG342011076", 10.0),
        "Text Field7": ("On 12 Oct 2026, sample ETX-261002-0288 was found to be positive for growth after Celsis Sterility Testing.", 8.0),
        "Text Field8": ("MICRO-SOP-44", 9.5),
        "Text Field9": ("03-Jan-2025", 9.5),
        "Text Field10": ("01", 9.5),
        "Text Field11": ("Pass/Fail", 9.5),
        "Text Field12": ("Kathan Parikh", 10.0),
        "Date Field3": ("12-Oct-2026", 10.0),
        "Text Field13": ("Yes, analysts Mukyung Jang, Alex Saravia, and Cuong Du were comprehensively interviewed.", 4.825),
        "Text Field14": ("Yes, Sample ID ETX-261002-0288", 8.5),
        "Text Field15": ("Yes, as per MICRO-SOP-44", 9.0),
        "Text Field16": ("Yes, as per MICRO-SOP-44", 9.0),
        "Text Field17": ("Yes, Information is available in EagleTrax under ETX-261002-0288", 6.725),
        "Text Field18": ("Yes, all analysts are trained and qualified by the quality department to perform the test", 4.825),
        "Text Field19": ("Not Applicable", 9.0),
        "Text Field20": ("Not Applicable", 9.0),
        "Text Field21": ("Yes, the sample was stored refrigerated as per client's instructions", 5.85),
        "Text Field22": (
            "Materials and Reagents\r"
            "TSB: 08132026-1\r"
            "FTM: 08182026-6\r"
            "Tween 80: 0000743947\r \r"
            "Celsis Reagents and Kits:\r"
            "AmpiScreen: 022601-1400\r"
            "Daily Wash: 012503-1142N\r \r"
            "TSA / Surface Plates: See EM logs\r \r"
            "Monthly Cleaning:\r"
            "H2O2 strips: 260620\r"
            "HaloMist: 2606A",
            4.35,
        ),
        "Text Field23": (
            "Materials and Reagents\r"
            "TSB: 11 Nov 2026\r"
            "FTM: 16 Nov 2026\r"
            "Tween 80: 31 Jan 2030\r \r"
            "Celsis Reagents and Kits:\r"
            "AmpiScreen: Jan 2027\r"
            "Daily Wash: Jan 2027\r \r"
            "TSA / Surface Plates: See EM logs\r \r"
            "Monthly Cleaning:\r"
            "H2O2 strips: Jun 2028\r"
            "HaloMist: Jun 2028",
            4.35,
        ),
        "Text Field24": ("Celsis ATP Positive Control", 9.0),
        "Text Field25": ("022601-1483", 9.0),
        "Text Field26": ("31 Jan 2027", 9.0),
        "Text Field27": ("Not Applicable", 9.0),
        "Text Field28": ("Not Applicable", 9.0),
        "Text Field29": ("Not Applicable", 9.0),
        "Text Field30": ("E002222", 9.5),
        "Text Field31": ("Sep 2027", 9.5),
        "Text Field32": ("E001736 (CR 114), BSC E001316\rE001736 (CR 114), BSC E001798", 4.825),
        "Text Field33": ("Dec 2026, Jun 2027\rDec 2026, Jun 2027", 4.825),
        "Text Field34": ("E002222", 9.5),
        "Text Field35": ("Sep 2027", 9.5),
        "Text Field36": ("Not Applicable", 9.0),
        "Text Field37": ("Not Applicable", 9.0),
        "Text Field38": ("Not Applicable", 9.0),
        "Text Field39": ("Not Applicable", 9.0),
        "Text Field40": ("See Phase I Summary", 7.6),
        "Text Field41": ("See Phase I Summary", 7.6),
        "Text Field42": ("See Phase I Summary", 7.6),
        "Text Field45": ("See Phase I Summary", 7.6),
        "Text Field43": (
            "Incubator E001356 (Sensor E001450)\r \r"
            "Incubator E001357 (Sensor E001449)\r \r"
            "Incubator E001034 (Sensor E001501)\r \r"
            "Incubator E001031 (Sensor E001505)",
            7.05,
        ),
        "Text Field44": (
            "Jan 2027\rFeb 2027\r \r"
            "Jan 2027\rFeb 2027\r \r"
            "Aug 2027 \rFeb 2027\r \r"
            "Aug 2027 \rFeb 2027",
            7.05,
        ),
        "Text Field46": ("Not Applicable", 8.5),
        "Text Field47": ("Not Applicable", 8.5),
        "Text Field48": ("N/A QYC 12Oct26", 8.0),
        "Text Field49": (p3_narrative, 9.2),
        "Text Field50": (p4_narrative, 9.2),
        "Text Field51": (p5_narrative, 9.2),
        "Text Field52": ("", 10.0),
        "Text Field53": ("Qiyue Chen", 10.0),
        "Text Field54": ("", 10.0),
        "Text Field55": ("", 10.0),
    }

    checked_boxes = {
        "Check Box0", "Check Box1", "Check Box2", "Check Box4", "Check Box7",
        "Check Box10", "Check Box13", "Check Box16", "Check Box19", "Check Box24",
        "Check Box27", "Check Box28", "Check Box32", "Check Box34", "Check Box38",
        "Check Box42", "Check Box43", "Check Box48", "Check Box51", "Check Box52",
        "Check Box55", "Check Box58", "Check Box63", "Check Box66", "Check Box67",
        "Check Box70", "Check Box73", "Check Box78", "Check Box79", "Check Box88"
    }

    for page in doc:
        for w in page.widgets():
            name = w.field_name
            if name in field_specs:
                val, fsize = field_specs[name]
                w.field_value = val
                w.text_fontsize = fsize
                w.update()
            elif w.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                w.field_value = True if name in checked_boxes else False
                w.update()

    tables_doc = fitz.open(tables_pdf_path)
    doc.insert_pdf(tables_doc)
    tables_doc.close()

    doc.save(final_pdf_path)
    doc.close()

if __name__ == "__main__":
    tmp_docx = os.path.join(WORKSPACE, "scratch", "tmp_tables_262333.docx")
    tmp_tables_pdf = os.path.join(WORKSPACE, "scratch", "tmp_tables_262333.pdf")
    tmp_master_pdf = os.path.join(WORKSPACE, "scratch", "tmp_master_262333.pdf")

    print("1. Building DOCX from OOS-262080 template...")
    build_tables_from_template(tmp_docx)

    print("2. Exporting DOCX to PDF via Word COM...")
    convert_docx_to_pdf(tmp_docx, tmp_tables_pdf)

    print("3. Filling AcroForm & appending tables...")
    fill_form_and_merge(tmp_tables_pdf, tmp_master_pdf)

    for folder in [WORKSPACE, DESKTOP, DOCS_DIR]:
        shutil.copy2(tmp_docx, os.path.join(folder, DOCX_NAME))
        shutil.copy2(tmp_master_pdf, os.path.join(folder, PDF_NAME))

    # Render all 8 pages to PNG for visual inspection
    doc_check = fitz.open(tmp_master_pdf)
    print(f"Done! Total Master PDF pages: {len(doc_check)}")
    for i, page in enumerate(doc_check):
        page.get_pixmap(dpi=150).save(os.path.join(WORKSPACE, "scratch", f"final_check_p{i+1}.png"))
    doc_check.close()
