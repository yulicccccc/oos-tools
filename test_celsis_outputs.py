import streamlit as st
import celsis_logic as cl

# Mock session state
st.session_state = {
    'prepper_name': 'Cuong Du', 'prepper_initial': 'CCD',
    'analyst_name': 'Devanshi Shah', 'analyst_initial': 'DS',
    'aliquoting_name': 'America Alanis', 'aliquoting_initial': 'ALA',
    'test_date': '14May26', 'process_date': '07May26',
    'bsc_id': '1314', 'positive_org': 'Staphylococcus Capitis',
    'other_positives': 'No',
    # Processing daily
    'pro_be_obs_pers': 'No growth', 'pro_on_obs_pers': 'No growth', 'pro_af_obs_pers': 'No growth',
    'pro_be_obs_surf': 'No growth', 'pro_on_obs_surf': 'No growth', 'pro_af_obs_surf': 'No growth',
    'pro_be_obs_sett': 'No growth', 'pro_on_obs_sett': 'No growth', 'pro_af_obs_sett': 'No growth',
    # Processing weekly
    'pro_be_obs_air_wk': '1 CFU', 'pro_be_etx_air_wk': 'ETX-260506-0280', 'pro_be_id_air_wk': 'Micrococcus luteus',
    'pro_on_obs_air_wk': '1 CFU', 'pro_on_etx_air_wk': 'ETX-260518-0249', 'pro_on_id_air_wk': 'Staphylococcus aureus',
    # Aliquoting daily
    'alq_be_obs_pers': 'No growth', 'alq_on_obs_pers': 'No growth', 'alq_af_obs_pers': 'No growth',
    'alq_be_obs_surf': 'No growth', 'alq_on_obs_surf': 'No growth', 'alq_af_obs_surf': 'No growth',
    'alq_be_obs_sett': 'No growth', 'alq_on_obs_sett': 'No growth', 'alq_af_obs_sett': 'No growth',
    # Aliquoting weekly
    'alq_be_obs_room_wk': '1 CFU', 'alq_be_etx_room_wk': 'ETX-260518-0243', 'alq_be_id_room_wk': 'Bacillus subtilis',
}

em_pro_narrative, em_alq_narrative, smart_just = cl.generate_celsis_narrative_and_details()
print(em_pro_narrative)
print("----")
print(em_alq_narrative)
print("----")
print(smart_just)
