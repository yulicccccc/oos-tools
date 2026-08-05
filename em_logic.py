# filename: em_logic.py
import streamlit as st
import os
import re
import json
import io
import sys
import subprocess
import time
from datetime import datetime, timedelta

# --- 1. Central Utilities ---
try:
    from utils import get_room_logic as u_grl, get_full_name, ordinal, num_to_words, get_cleanroom_narrative
except ImportError:
    def u_grl(i): return "Unknown", "000", "", "Unknown"
    def get_full_name(i): return i
    def ordinal(n): return str(n)
    def num_to_words(n): return str(n)
    def get_cleanroom_narrative(s, r=None, a="", v=""): return ""

# --- 2. CONFIG & KEYS ---
FIELD_KEYS = [
    "oos_id", "client_name", "sample_id", "test_date", "sample_name", "lot_number", 
    "dosage_form", "monthly_cleaning_date", 
    "analyst_initial", "analyst_name", "reader_initial", "reader_name",
    "writer_name", "bsc_id", "cr_id", "shift_number", "sampling_type",
    "org_choice", "manual_org", "action_level", "cfu_count", "event_number", "confirm_number",
    "obs_pers_dur", "etx_pers_dur", "id_pers_dur", 
    "obs_surf_dur", "etx_surf_dur", "id_surf_dur", 
    "obs_sett_dur", "etx_sett_dur", "id_sett_dur", 
    "obs_air_wk_of", "etx_air_wk_of", "id_air_wk_of", 
    "obs_room_wk_of", "etx_room_wk_of", "id_room_wk_of",
    "date_of_weekly", "weekly_initial"
]

def auto_fill_name(initial_key, name_key):
    initial = st.session_state.get(initial_key, "")
    current_name = st.session_state.get(name_key, "")
    if initial:
        calculated_name = get_full_name(initial)
        if calculated_name and not current_name:
            st.session_state[name_key] = calculated_name
            st.rerun()

def validate_inputs():
    errors, warnings = [], []
    reqs = {
        "OOS Number": "oos_id", "Sample / Plate Name": "sample_name", 
        "Test Date": "test_date", "Setup Analyst Name": "analyst_name",
        "Reader Name": "reader_name", "BSC / Cleanroom ID": "bsc_id"
    }
    for label, key in reqs.items():
        if not st.session_state.get(key, "").strip(): warnings.append(label)
    date_val = st.session_state.get("test_date", "").strip()
    if date_val:
        try: datetime.strptime(date_val, "%d%b%y")
        except ValueError: errors.append(f"❌ Date Error: '{date_val}' invalid. Use DDMMMYY (e.g. 17Feb26).")
    return errors, warnings

def clean_filename(text): 
    return re.sub(r'[\\/*?:"<>|]', '_', str(text)).strip() if text else ""

# --- 3. NARRATIVE GENERATION LOGIC ---
def generate_em_narrative():
    s = st.session_state
    
    analyst_name = s.get("analyst_name", "[Analyst Name]")
    reader_name = s.get("reader_name", "[Reader Name]")
    sampling_type = s.get("sampling_type", "Surface Sampling")
    bsc_id = s.get("bsc_id", "BSC 1309")
    plate_name = s.get("sample_name", "EM Plate")
    cfu_count = s.get("cfu_count", "10")
    org_identified = s.get("manual_org", "Staphylococcus epidermidis")
    
    # 1. Interview & Storage Block
    interview_block = (
        f"The analyst involved in the {sampling_type} plate setup, {analyst_name}, and the analyst(s) involved "
        f"in reading the plate, {reader_name}, were interviewed comprehensively. Their answers are recorded throughout this document.\n"
        f"The EM plates were stored in compliance with the supplier's recommendations, and their integrity was visually inspected prior to use. "
        f"Furthermore, the plates were confirmed to be within their valid expiration dates. All the supplies were thoroughly disinfected prior to use."
    )
    
    # 2. EM Records Block
    records_block = (
        f"Environmental Monitoring Summary: Personnel sampling plates for analyst {analyst_name} and "
        f"routine monitoring plates for ISO 5 {bsc_id}, for the previous date, date of, and following date of testing showed no microbial growth. "
        f"However, the {sampling_type} plate for ISO 5 {bsc_id} exhibited {cfu_count} CFUs on {plate_name} for the day of testing "
        f"performed by analyst {analyst_name}, which were identified as {org_identified}. "
        f"No growth was observed on other surface/settling plates for the date of and following date of testing."
    )
    
    # 3. Phase I Summary / Defensive Conclusion
    summary_block = (
        f"Based on the findings outlined in the preceding sections, the Out-Of-Specification (OOS) result observed for "
        f"the Environmental Monitoring (EM) {sampling_type} plate may be attributed to a potential analyst error or transient laboratory contamination.\n"
        f"No growth was observed on the analyst's personnel/surface/settling plates collected from the BSC for the following day, "
        f"indicating that the contamination was transient in nature and that routine daily disinfection procedures were effective in eliminating any residual contamination. "
        f"Furthermore, all testing in the ISO 5 {bsc_id} occurred under controlled conditions with proper physical isolation from the background environment."
    )
    
    return interview_block, records_block, summary_block
