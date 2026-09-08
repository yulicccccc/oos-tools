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
data['cross_contamination_summary'] = "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."

# 6. Narrative & EM Details + Smart Justification Engine (4-Step Shielding)
pass_daily_clean = ['surface sampling']
narr_parts = ["no microbial growth was observed in surface sampling"]
narr = "Upon analyzing the environmental monitoring results, " + ". ".join(narr_parts) + "."

daily_fails = [f['cat'] for f in failures if f['time'] == 'daily']
weekly_fails = [f['cat'] for f in failures if f['time'] == 'weekly']
intro_parts = []
if daily_fails: intro_parts.append(f"{' and '.join(daily_fails)} on the date of testing")
if weekly_fails: intro_parts.append(f"{' and '.join(weekly_fails)} from the week of testing")
fail_intro = f"However, microbial growth was observed during {' and '.join(intro_parts)}."

detail_sentences = []
for i, f in enumerate(failures):
    is_plural = bool(re.search(r'\d+', f['obs']) and int(re.search(r'\d+', f['obs']).group()) > 1)
    verb_detect = 'were' if is_plural else 'was'
    noun_id = 'organisms were' if is_plural else 'organism was'
    method_text = 'differential staining' if 'gram' in f['id'].lower() else 'microbial identification'
    note_txt = f" (Note: {f['note']})" if f.get('note') and f['note'].strip() and f['note'].strip().lower() != 'none' else ""
    base_sentence = f"{f['obs']} {verb_detect} detected during {f['cat']} and was submitted for {method_text} under sample ID {f['etx']}, where the {noun_id} identified as {f['id']}{note_txt}"
    lead = "Specifically" if i==0 else "Additionally" if i==1 else "Furthermore" if i==2 else "Also"
    detail_sentences.append(f"{lead}, {base_sentence}.")

det = f"{fail_intro} {' '.join(detail_sentences)}"

# Smart Justification Shielding Block
smart_just_blocks = [
    "Notably, the colony morphology of the recovered microorganisms differed between monitoring locations and the test sample. While the test sample exhibited rod-shaped morphology, personnel touch plates recovered Micrococcus luteus (cocci), and settling plates showed insufficient growth of Gram (+) rods with plate desiccation documented on the 5-day read. This indicates that the personnel monitoring findings were isolated, unrelated events.",
    "Also, while microbial growth was detected during weekly monitoring (Gram (+) short rods in air and Bacillus megaterium on the floor), these organisms were detected strictly in the ISO 8 background room environment (Cleanroom 115) and ISO 7 floor (Suite 115A), whereas the sample manipulation occurred strictly within the ISO 5 Primary Engineering Control (BSC E001313).",
    "Additionally, the absence of contamination on ISO 5 work surface monitoring inside BSC E001313 indicates that no viable transfer pathway existed from outer cleanroom areas to the critical ISO 5 testing zone.",
    "Furthermore, all other samples processed by the analyst that day tested negative for microbial growth, confirming that the testing environment operated under optimal conditions and cross-contamination did not occur.",
    "Based on the observations outlined above, it is unlikely that the failing results were due to reagents, supplies, the cleanroom environment, the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is minimal and the original result is deemed to be valid."
]
smart_just_text = "\n\n".join(smart_just_blocks)

data['narrative_summary'] = narr
data['em_details'] = det
data['smart_justification'] = smart_just_text

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
p9 = narr + "\n\n" + det + "\n\n" + smart_just_text
p10 = f"Monthly cleaning and disinfection, using H2O2, of the cleanroom (ISO 7) and its containing Biosafety Cabinets (BSCs, ISO 5) were performed on {data['monthly_cleaning_date']}, as per SOP 2.600.018 Cleaning and Disinfection Procedure. It was documented that all H2O2 indicators passed."
p11 = data['sample_history_paragraph']
p12 = f"To assess the potential for sample-to-sample contamination contributing to the positive results, a comprehensive review was conducted of all samples processed on the same day. {data['cross_contamination_summary']}"
p13 = "Based on the observations outlined above, it is unlikely that the failing results were due to reagents, supplies, the cleanroom environment, the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is minimal and the original result is deemed to be valid."

smart_phase1_part2 = "\n\n".join([p7, p8, p9, p10, p11, p12, p13])
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

# Save updated JSON state
out_json_path = os.path.join(OUTPUT_DIR, f"UPDATED_SAVE_OOS-{data['oos_id']} {data['client_name']} - ScanRDI.txt")
with open(out_json_path, "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2)
print("Saved updated JSON state to:", out_json_path)

# Render Word Report
tpl_report = DocxTemplate("ScanRDI OOS template 0.docx")
tpl_report.render(data)
out_doc_report = os.path.join(OUTPUT_DIR, f"OOS-{data['oos_id']} {data['client_name']} - ScanRDI.docx")
tpl_report.save(out_doc_report)
print("Saved Word Report to:", out_doc_report)

# Render Word Tables
tpl_tables = DocxTemplate("tables for scan.docx")
tpl_tables.render(data)
out_doc_tables = os.path.join(OUTPUT_DIR, f"Tables OOS-{data['oos_id']} {data['client_name']} - ScanRDI.docx")
tpl_tables.save(out_doc_tables)
print("Saved Word Tables to:", out_doc_tables)

# Render PDF Form
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
