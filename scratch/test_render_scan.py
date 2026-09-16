import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import json, docx
from docxtpl import DocxTemplate
from utils import get_room_logic, get_cleanroom_narrative

with open(r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\SAVE_OOS-261814 GoGoMeds Select (E10747) - ScanRDI.txt', 'r', encoding='utf-8') as f:
    data = json.load(f)

doc_em = docx.Document(r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\EM table 07AUG2026 (1).docx')
table = doc_em.tables[0]

for row in table.rows:
    c = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
    site = c[0]
    if 'Personal (Left' in site:
        data['obs_pers_dur'] = c[5]; data['etx_pers_dur'] = c[6]; data['id_pers_dur'] = c[7]; data['note_pers'] = c[8]
    elif 'Surface Sampling of ISO 5' in site:
        data['obs_surf_dur'] = c[5]; data['etx_surf_dur'] = c[6]; data['id_surf_dur'] = c[7]; data['note_surf'] = c[8]
    elif 'Settling Sampling of ISO 5' in site:
        data['obs_sett_dur'] = c[5]; data['etx_sett_dur'] = c[6]; data['id_sett_dur'] = c[7]; data['note_sett'] = c[8]
    elif 'Active Air Sampling' in site:
        data['obs_air_wk_of'] = c[5]; data['etx_air_wk_of'] = c[6]; data['id_air_wk_of'] = c[7]; data['date_of_weekly'] = c[2]; data['weekly_initial'] = c[3]
    elif 'Surface Sampling of Cleanrooms' in site:
        data['obs_room_wk_of'] = c[5]; data['etx_room_wk_of'] = c[6]; data['id_room_wk_of'] = c[7]

t_room, t_suite, t_suffix, t_loc = get_room_logic(data['bsc_id'])
data['cr_id'] = t_room
data['cr_suit'] = t_suite
data['suit'] = t_suffix
data['bsc_location'] = t_loc
data['organism_morphology'] = 'Rod'
data['control_positive'] = data.get('control_pos', 'A. brasiliensis')
data['control_data'] = data.get('control_exp', '22May28')
data['test_record'] = '080726-1040-2'

# Equipment summary
p1 = get_cleanroom_narrative(t_suite, action_text='testing and changeover procedures', verb='comprises')
p2 = f"The ISO 5 BSC E00{data['bsc_id']}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both testing and changeover steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{data['bsc_id']} was certified and approved by both the Engineering and Quality Assurance teams. Sample processing and changeover were conducted in the ISO 5 BSC E00{data['bsc_id']} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {data['analyst_name']} on {data['test_date']}."
data['equipment_summary'] = f"{p1}\n\n{p2}"

# Cross contamination
data['cross_contamination_summary'] = "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."

# History
data['sample_history_paragraph'] = f"Analyzing a 6-month sample history for {data['client_name']}, this specific analyte \"{data['sample_name']}\" has had no prior failures using the Scan RDI method during this period."

# Narrative summary
data['narrative_summary'] = (
    "Upon analyzing the environmental monitoring results, no microbial growth was observed in surface sampling. "
    "However, microbial growth was observed during personnel sampling (right touch) and settling plates on the date of testing, and weekly active air sampling and weekly surface sampling from the week of testing. "
    "Specifically, 1 CFU on RT was detected during personnel sampling and was submitted for microbial identification under sample ID ETX-260817-0447, where the organism was identified as Micrococcus luteus (plate was desiccated on 5 day read). "
    "Additionally, 2 CFUs on sett 2 were detected during settling plates and was submitted for differential staining under sample ID ETX-260817-0507, where the organisms were identified as Insufficient read Gram (+) rods (plate was desiccated on 5 day read). "
    "Furthermore, 1 CFU in ISO 8 115 was detected during weekly active air sampling under sample ID ETX-260817-0370 (Gram (+) short rods). "
    "Also, 2 CFUs on floor in 115A ISO 7 were detected during weekly surface sampling under sample ID ETX-260817-0366 (Bacillus megaterium)."
)

tpl = DocxTemplate('ScanRDI OOS P1 template 0.docx')
tpl.render(data)
out_path = 'scratch/test_report_p1_rendered.docx'
tpl.save(out_path)
print(f'Successfully rendered {out_path}!')
