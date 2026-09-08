import os, sys, json, re, io
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import docx
from docxtpl import DocxTemplate
from pypdf import PdfWriter, PdfReader
from datetime import datetime
from utils import get_room_logic, get_cleanroom_narrative, ordinal, num_to_words, get_full_name

SAVE_FILE = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\SAVE_OOS-261814 GoGoMeds Select (E10747) - ScanRDI.txt"
EM_DOCX_FILE = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\EM table 07AUG2026 (1).docx"
OUTPUT_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\scratch"

# 1. Load OOS base save
with open(SAVE_FILE, "r", encoding="utf-8") as f:
    data = json.load(f)

# 2. Parse EM docx
doc_em = docx.Document(EM_DOCX_FILE)
table = doc_em.tables[0]

em_raw = {}
for row in table.rows:
    c = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
    site = c[0]
    if 'Personal (Left' in site:
        em_raw['pers'] = {'site': site, 'freq': c[1], 'date': c[2], 'analyst': c[3], 'timing': c[4], 'obs': c[5], 'etx': c[6], 'id': c[7], 'note': c[8]}
    elif 'Surface Sampling of ISO 5' in site:
        em_raw['surf'] = {'site': site, 'freq': c[1], 'date': c[2], 'analyst': c[3], 'timing': c[4], 'obs': c[5], 'etx': c[6], 'id': c[7], 'note': c[8]}
    elif 'Settling Sampling of ISO 5' in site:
        em_raw['sett'] = {'site': site, 'freq': c[1], 'date': c[2], 'analyst': c[3], 'timing': c[4], 'obs': c[5], 'etx': c[6], 'id': c[7], 'note': c[8]}
    elif 'Active Air Sampling' in site:
        em_raw['air'] = {'site': site, 'freq': c[1], 'date': c[2], 'analyst': c[3], 'timing': c[4], 'obs': c[5], 'etx': c[6], 'id': c[7], 'note': c[8]}
    elif 'Surface Sampling of Cleanrooms' in site:
        em_raw['room'] = {'site': site, 'freq': c[1], 'date': c[2], 'analyst': c[3], 'timing': c[4], 'obs': c[5], 'etx': c[6], 'id': c[7], 'note': c[8]}

print("Parsed EM Raw:", json.dumps(em_raw, indent=2))

# 3. Update session state fields
data['obs_pers'] = em_raw['pers']['obs']
data['etx_pers'] = em_raw['pers']['etx']
data['id_pers'] = em_raw['pers']['id']
data['obs_pers_dur'] = em_raw['pers']['obs']
data['etx_pers_dur'] = em_raw['pers']['etx']
data['id_pers_dur'] = em_raw['pers']['id']

data['obs_surf'] = em_raw['surf']['obs']
data['etx_surf'] = em_raw['surf']['etx']
data['id_surf'] = em_raw['surf']['id']
data['obs_surf_dur'] = em_raw['surf']['obs']
data['etx_surf_dur'] = em_raw['surf']['etx']
data['id_surf_dur'] = em_raw['surf']['id']

data['obs_sett'] = em_raw['sett']['obs']
data['etx_sett'] = em_raw['sett']['etx']
data['id_sett'] = em_raw['sett']['id']
data['obs_sett_dur'] = em_raw['sett']['obs']
data['etx_sett_dur'] = em_raw['sett']['etx']
data['id_sett_dur'] = em_raw['sett']['id']

data['obs_air'] = em_raw['air']['obs']
data['etx_air_weekly'] = em_raw['air']['etx']
data['id_air_weekly'] = em_raw['air']['id']
data['obs_air_wk_of'] = em_raw['air']['obs']
data['etx_air_wk_of'] = em_raw['air']['etx']
data['id_air_wk_of'] = em_raw['air']['id']

data['obs_room'] = em_raw['room']['obs']
data['etx_room_weekly'] = em_raw['room']['etx']
data['id_room_wk_of'] = em_raw['room']['id']
data['obs_room_wk_of'] = em_raw['room']['obs']
data['etx_room_wk_of'] = em_raw['room']['etx']

# Dates & Initials
clean_weekly_d = em_raw['air']['date'].replace(' ', '')
try:
    d_weekly_obj = datetime.strptime(clean_weekly_d, "%d%b%Y")
    formatted_weekly_d = d_weekly_obj.strftime("%d%b%y")
except:
    formatted_weekly_d = clean_weekly_d

data['date_weekly'] = formatted_weekly_d
data['date_of_weekly'] = formatted_weekly_d
data['weekly_init'] = em_raw['air']['analyst']
data['weekly_initial'] = em_raw['air']['analyst']

# Dynamic EM UI list in Session State
data['em_growth_observed'] = "Yes"
failures = []
if data['obs_pers'].lower() != 'no growth':
    failures.append({"cat": "personnel sampling (right touch)", "obs": data['obs_pers'], "etx": data['etx_pers'], "id": data['id_pers'], "time": "daily", "note": em_raw['pers']['note']})
if data['obs_surf'].lower() != 'no growth':
    failures.append({"cat": "surface sampling", "obs": data['obs_surf'], "etx": data['etx_surf'], "id": data['id_surf'], "time": "daily", "note": em_raw['surf']['note']})
if data['obs_sett'].lower() != 'no growth':
    failures.append({"cat": "settling plates", "obs": data['obs_sett'], "etx": data['etx_sett'], "id": data['id_sett'], "time": "daily", "note": em_raw['sett']['note']})
if data['obs_air'].lower() != 'no growth':
    failures.append({"cat": "weekly active air sampling", "obs": data['obs_air'], "etx": data['etx_air_weekly'], "id": data['id_air_weekly'], "time": "weekly", "note": em_raw['air']['note']})
if data['obs_room'].lower() != 'no growth':
    failures.append({"cat": "weekly surface sampling", "obs": data['obs_room'], "etx": data['etx_room_weekly'], "id": data['id_room_wk_of'], "time": "weekly", "note": em_raw['room']['note']})

data['em_growth_count'] = len(failures)
cat_map = {
    "personnel sampling (right touch)": "Personnel Obs",
    "surface sampling": "Surface Obs",
    "settling plates": "Settling Obs",
    "weekly active air sampling": "Weekly Air Obs",
    "weekly surface sampling": "Weekly Surf Obs"
}
for i, f in enumerate(failures):
    data[f'em_cat_{i}'] = cat_map.get(f['cat'], "")
    data[f'em_obs_{i}'] = f['obs']
    data[f'em_etx_{i}'] = f['etx']
    data[f'em_id_{i}'] = f['id']

# 4. Generate Equipment Text
t_room, t_suite, t_suffix, t_loc = get_room_logic(data['bsc_id'])
data['cr_id'] = t_room
data['cr_suit'] = t_suite
data['suit'] = t_suffix
data['bsc_location'] = t_loc
data['organism_morphology'] = 'Rod'
data['control_positive'] = data.get('control_pos', 'A. brasiliensis')
data['control_data'] = data.get('control_exp', '22May28')

d_obj = datetime.strptime(data['test_date'], "%d%b%y")
tr_id = f"{d_obj.strftime('%m%d%y')}-{data['scan_id']}-{data['shift_number']}"
pdf_date_str = d_obj.strftime("%d-%b-%Y")
data['test_record'] = tr_id

part1 = get_cleanroom_narrative(t_suite, action_text="testing and changeover procedures", verb="comprises")
part2 = f"The ISO 5 BSC E00{data['bsc_id']}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both testing and changeover steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{data['bsc_id']} was certified and approved by both the Engineering and Quality Assurance teams. Sample processing and changeover were conducted in the ISO 5 BSC E00{data['bsc_id']} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {data['analyst_name']} on {data['test_date']}."
data['equipment_summary'] = f"{part1}\n\n{part2}"

# 5. History & Cross-Contamination
data['sample_history_paragraph'] = f"Analyzing a 6-month sample history for {data['client_name']}, this specific analyte \"{data['sample_name']}\" has had no prior failures using the Scan RDI method during this period."
# 5. History & Cross-Contamination
data['sample_history_paragraph'] = f"Analyzing a 6-month sample history for {data['client_name']}, this specific analyte \"{data['sample_name']}\" has had no prior failures using the Scan RDI method during this period."
data['cross_contamination_summary'] = "The microbiological findings were also assessed for evidence of a broader contamination pattern. All other samples processed by the same analyst, as well as samples processed by other laboratory personnel on the date of testing, were negative for microbial growth. The absence of additional positive samples or a clustering pattern provides further evidence against a systemic environmental, procedural, or sample-to-sample cross-contamination event."

# 6. Narrative & EM Details + FDA-Aligned cGMP Defense Engine
p_em_intro = f"Upon review of the environmental monitoring data associated with the sterility test, no microbial growth was recovered from any of the four ISO 5 work-surface monitoring locations within BSC E00{data['bsc_id']} on the date of testing. Low-level microbial recoveries were identified from personnel and settling-plate monitoring performed in association with testing, as well as from routine weekly monitoring of the surrounding cleanroom areas."

p_personnel = "One CFU was recovered from the analyst’s right-hand touch plate and submitted for microbial identification under sample ID ETX-260817-0447. The isolate was identified as Micrococcus luteus. This recovery was not microbiologically consistent with the organism observed in the test sample: M. luteus is coccal in morphology, whereas the microorganism recovered from the test sample exhibited rod-shaped morphology. Accordingly, the personnel monitoring result does not support direct transfer of the recovered right-glove organism to the test sample."

p_settling = f"Two CFUs were also recovered from settling plate Sett 2 and submitted for differential staining under sample ID ETX-260817-0507. The organisms were characterized as Gram-positive rods; however, definitive identification could not be obtained because the plate was documented as desiccated at the 5-day read. Therefore, an organism-level microbiological match between the settling-plate recovery and the test-sample isolate could not be established. This finding was evaluated in conjunction with the remaining contemporaneous environmental monitoring data, including the absence of microbial recovery from all four ISO 5 work surfaces within BSC E00{data['bsc_id']}."

p_weekly = f"Routine weekly facility monitoring during the week of testing additionally recovered 1 CFU of Gram-positive short rods from active air monitoring in ISO 8 Cleanroom 115 (ETX-260817-0370) and 2 CFUs identified as Bacillus megaterium from the floor of ISO 7 Suite 115A (ETX-260817-0366). These recoveries occurred in lower-classified background areas physically separated from the critical ISO 5 testing zone. Sample manipulation was performed within BSC E00{data['bsc_id']}, and samples were transported into the testing area in disinfected, lidded containers. No corresponding microbial recovery was observed from the ISO 5 work surfaces within the BSC that would support transfer of contamination from these surrounding areas into the critical testing environment."

p_monthly = f"Monthly cleaning and disinfection, using H2O2, of the cleanroom (ISO 7) and its containing Biosafety Cabinets (BSCs, ISO 5) were performed on {data['monthly_cleaning_date']}, as per SOP 2.600.018 Cleaning and Disinfection Procedure. It was documented that all H2O2 indicators passed."

p_history = data['sample_history_paragraph']

p_cross = data['cross_contamination_summary']

p_conclusion = "Taken together, the environmental monitoring results demonstrate isolated, low-level recoveries at discrete monitoring locations, with no microbiological match established between the recovered monitoring organisms and the test-sample isolate, no microbial recovery from the ISO 5 work surfaces used for testing, and no broader pattern of contamination among concurrently processed samples. Therefore, the available environmental and personnel monitoring data do not identify an assignable laboratory source for the microbial growth observed in the test sample and do not support laboratory-introduced contamination as the cause of the positive sterility result."

# Table 3 fields for other positives
data['oos1_analyst_name'] = data['analyst_name']
data['oos1_sample_id'] = data['sample_id']
data['oos1_sample_name'] = data['sample_name']
data['oos1_organism_morphology'] = data['organism_morphology']

# Paragraphs for Form 3.100.019.F01
suffix = "microorganism" if str(data.get('confirm_number','1')).strip() == "1" else "microorganisms"
org_lower = str(data.get('organism_morphology', 'rod')).strip().lower()
p7 = f"On {data['test_date']}, a rapid sterility test was conducted on the sample using the ScanRDI method. The sample was initially prepared by Analyst {data['prepper_name']}, processed by {data['analyst_name']}, and subsequently read by {data['reader_name']}. The test revealed {data['confirm_number']} {org_lower}-shaped viable {suffix}, see table 1."
p8 = f"Table 2 (see attached tables) presents the environmental monitoring results for {data['sample_id']}. The environmental monitoring (EM) plates were incubated for no less than 48 hours at 30-35°C and no less than an additional five days at 20-25°C as per SOP 2.600.002 (Environmental Monitoring of the Clean-room Facility)."

smart_phase1_part2 = "\n\n".join([p7, p8, p_em_intro, p_personnel, p_settling, p_weekly, p_monthly, p_history, p_cross, p_conclusion])
data['narrative_summary'] = "\n\n".join([p_em_intro, p_personnel, p_settling, p_weekly])
data['smart_justification'] = p_conclusion
data['smart_phase1_part2'] = smart_phase1_part2
data['Text Field50'] = smart_phase1_part2

# Collect and deduplicate all analyst names
analysts_raw = [
    data.get("prepper_name", ""),
    data.get("analyst_name", ""),
    data.get("changeover_name", ""),
    data.get("reader_name", "")
]
analysts_clean = [str(x).strip() for x in analysts_raw if str(x).strip() and str(x).strip() != "N/A"]
analysts_unique = list(dict.fromkeys(analysts_clean))

if not analysts_unique:
    names_only_phrase = "N/A"
    analysts_with_prefix_phrase = "the analysts"
elif len(analysts_unique) == 1:
    names_only_phrase = analysts_unique[0]
    analysts_with_prefix_phrase = f"analyst {analysts_unique[0]}"
elif len(analysts_unique) == 2:
    names_only_phrase = f"{analysts_unique[0]} and {analysts_unique[1]}"
    analysts_with_prefix_phrase = f"analysts {analysts_unique[0]} and {analysts_unique[1]}"
else:
    names_only_phrase = ", ".join(analysts_unique[:-1]) + ", and " + analysts_unique[-1]
    analysts_with_prefix_phrase = "analysts " + ", ".join(analysts_unique[:-1]) + ", and " + analysts_unique[-1]

smart_comment_interview = f"Yes, {analysts_with_prefix_phrase} were interviewed comprehensively."
data['smart_comment_interview'] = smart_comment_interview

# Form Phase 1 Part 1
analyst_sig_text = f"{data['analyst_name']} (Written by: Qiyue Chen)"
smart_personnel_block = f"Prepper: \n{data['prepper_name']} ({data['prepper_initial']})\n\nProcessor:\n{data['analyst_name']} ({data['analyst_initial']})\n\nChangeover\nProcessor:\n{data['changeover_name']} ({data['changeover_initial']})\n\nReader:\n{data['reader_name']} ({data['reader_initial']})"
smart_incident_opening = f"On {data['test_date']}, sample {data['sample_id']} was found positive for viable microorganisms after ScanRDI testing."
smart_comment_samples = f"Yes, {data['sample_id']}"
smart_comment_records = f"Yes, See {tr_id} for more information."
smart_comment_storage = f"Yes, Information is available in Eagle Trax Sample Location History under {data['sample_id']}"

p1 = f"All analysts involved in the prepping, processing, and reading of the samples – {names_only_phrase} – were interviewed and their answers are recorded throughout this document."
p2 = f"The sample was stored upon arrival according to the Client’s instructions. Analysts {data['prepper_name']} and {data['analyst_name']} confirmed the integrity of the samples throughout both the preparation and processing stages. No leaks or turbidity were observed at any point, verifying the integrity of the sample."
p3 = "All reagents and supplies mentioned in the material section above were stored according to the suppliers’ recommendations, and their integrity was visually verified before utilization. Moreover, each reagent and supply had valid expiration dates."
p4 = f"During the preparation phase, {data['prepper_name']} disinfected the samples using acidified bleach and placed them into a pre-disinfected storage bin. On {data['test_date']}, prior to sample processing, {data['analyst_name']} performed a second disinfection with acidified bleach, allowing a minimum contact time of 10 minutes before transferring the samples into the cleanroom suites. A final disinfection step was completed immediately before the samples were introduced into the ISO 5 Biological Safety Cabinet (BSC), E00{data['bsc_id']}, located within the {t_loc}, (Suite {t_suite}{t_suffix}), All activities were performed in accordance with MICRO-SOP-12, Rapid Scan RDI® Test using FIFU Method."
p5 = data['equipment_summary']
p6 = f"The analyst, {data['reader_name']}, confirmed that the equipment was set up as per ENG-SOP-4 (Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup) and Maintenance), and the negative control and the positive control for the analyst, {data['reader_name']}, yielded expected results."
smart_phase1_part1 = "\n\n".join([p1, p2, p3, p4, p5, p6])

data['smart_phase1_part1'] = smart_phase1_part1
data['smart_phase1_summary'] = smart_phase1_part1
data['smart_phase1_continued'] = smart_phase1_part2
data['smart_cr_id'] = f"CR{t_suite} (E00{t_room})"
data['smart_scan_id'] = f"E00{data['scan_id']}"
data['analyst_signature'] = analyst_sig_text
data['report_header'] = f"{data['sample_id']}\n\n{data['client_name']}"
data['smart_personnel_block'] = smart_personnel_block
data['smart_incident_opening'] = smart_incident_opening
data['smart_comment_samples'] = smart_comment_samples
data['smart_comment_records'] = smart_comment_records
data['smart_comment_storage'] = smart_comment_storage

# Save updated JSON state
out_json_path = os.path.join(OUTPUT_DIR, f"UPDATED_SAVE_OOS-{data['oos_id']} {data['client_name']} - ScanRDI.txt")
with open(out_json_path, "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2)
print("Saved updated JSON state to:", out_json_path)

# Render Word Report (Target primary template: ScanRDI OOS P1 template.docx)
primary_tpl = "ScanRDI OOS P1 template.docx" if os.path.exists("ScanRDI OOS P1 template.docx") else "ScanRDI OOS template 0.docx"
tpl_report = DocxTemplate(primary_tpl)
tpl_report.render(data)
out_doc_report = os.path.join(OUTPUT_DIR, f"OOS-{data['oos_id']} {data['client_name']} - ScanRDI.docx")
tpl_report.save(out_doc_report)
print(f"Saved Word Report (using {primary_tpl}) to: {out_doc_report}")

# Render Word Tables
tpl_tables = DocxTemplate("tables for scan.docx")
tpl_tables.render(data)
out_doc_tables = os.path.join(OUTPUT_DIR, f"Tables OOS-{data['oos_id']} {data['client_name']} - ScanRDI.docx")
tpl_tables.save(out_doc_tables)
print("Saved Word Tables to:", out_doc_tables)

pdf_map = {
    'Text Field57': data['oos_id'],
    'Date Field0': pdf_date_str,
    'Date Field1': pdf_date_str,
    'Date Field2': pdf_date_str,
    'Date Field3': pdf_date_str,
    'Text Field2': data['sample_id'],
    'Text Field6': data['lot_number'],
    'Text Field4': data['sample_name'] + "\n\n\n\n",
    'Text Field5': data['dosage_form'],
    'Text Field8': "MICRO-SOP-12 (16)\rENG-SOP-4 (05)",
    'Text Field9': "24Jul26\r24Jul26",
    'Text Field10': "Rev: 16\rRev: 05",
    'Text Field15': "Yes, as per MICRO-SOP-12, ENG-SOP-4",
    'Text Field16': "Yes, as per MICRO-SOP-12, ENG-SOP-4",
    'Text Field0': analyst_sig_text,
    'Text Field3': smart_personnel_block,
    'Text Field7': smart_incident_opening + "\n\n",
    'Text Field13': smart_comment_interview,
    'Text Field14': smart_comment_samples,
    'Text Field17': smart_comment_records,
    'Text Field21': smart_comment_storage,
    'Text Field30': f"E00{data['scan_id']}",
    'Text Field32': f"E00{t_room} (CR{t_suite})",
    'Text Field34': f"E00{data['scan_id']}",
    'Text Field24': data['control_pos'],
    'Text Field25': data['control_lot'] + "\n\n\n",
    'Text Field26': data['control_exp'],
    'Text Field49': smart_phase1_part1,
    'Text Field50': smart_phase1_part2
}

if os.path.exists("ScanRDI OOS template.pdf"):
    writer = PdfWriter(clone_from="ScanRDI OOS template.pdf")
    for p in writer.pages:
        writer.update_page_form_field_values(p, pdf_map)
    out_pdf_report = os.path.join(OUTPUT_DIR, f"OOS-{data['oos_id']} {data['client_name']} - ScanRDI.pdf")
    with open(out_pdf_report, "wb") as f:
        writer.write(f)
    print("Saved PDF Report to:", out_pdf_report)

print("\n--- ALL GENERATION COMPLETED SUCCESSFULLY! ---")
