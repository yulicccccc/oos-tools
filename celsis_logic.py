# filename: celsis_logic.py
import streamlit as st
import re
from utils import get_room_logic as u_grl, get_full_name, ordinal, num_to_words

# --- 1. CONFIG & KEYS (前后端数据契约) ---
FIELD_KEYS = [
    "oos_id", "client_name", "sample_id", "test_date", "sample_name", "lot_number", 
    "dosage_form", "monthly_cleaning_date", 
    "prepper_initial", "prepper_name", "analyst_initial", "analyst_name",
    "aliquoting_initial", "aliquoting_name", 
    "bsc_id", "celsis_id", "test_record",
    "positive_media", "positive_id", "positive_org",
    "control_lot", "control_data",
    "incidence_count", "has_prior_failures",
    "other_positives", "total_pos_count_num", "current_pos_order"
]
for i in range(20):
    FIELD_KEYS.extend([f"other_id_{i}", f"other_order_{i}", f"prior_oos_{i}"])

# --- 2. HELPER FUNCTIONS (从 Scan 移植的防呆安全锁) ---

def auto_fill_name(initial_key, name_key):
    """自动将缩写转换为全名并刷新界面"""
    initial = st.session_state.get(initial_key, "")
    current_name = st.session_state.get(name_key, "")
    if initial:
        calculated_name = get_full_name(initial)
        if calculated_name and not current_name:
            st.session_state[name_key] = calculated_name
            st.rerun()

def validate_inputs():
    """Celsis 专属的表单完整性体检仪"""
    errors, warnings = [], []
    reqs = {
        "OOS Number": "oos_id", "Client Name": "client_name", "Sample ID": "sample_id", 
        "Test Date": "test_date", "Sample Name": "sample_name", "Lot Number": "lot_number",
        "Prepper Name": "prepper_name", "Processor Name": "analyst_name",
        "Aliquoting Name": "aliquoting_name", "Processing BSC ID": "bsc_id", 
        "Celsis ID": "celsis_id", "Positive Media": "positive_media",
        "Positive ID": "positive_id"
    }
    for label, key in reqs.items():
        if not st.session_state.get(key, "").strip(): 
            warnings.append(label)
            
    date_val = st.session_state.get("test_date", "").strip()
    if date_val:
        try: 
            # 校验日期必须是 DDMMMYY 格式
            from datetime import datetime
            datetime.strptime(date_val, "%d%b%y")
        except ValueError: 
            errors.append(f"❌ Date Error: '{date_val}' invalid. Use DDMMMYY (e.g. 17Mar26).")
            
    return errors, warnings

# --- 3. TEXT GENERATION LOGIC (重型文案生成引擎) ---

def generate_celsis_equipment_text():
    t_room, t_suite, t_suffix, t_loc = u_grl(st.session_state.bsc_id)
    a_suite, a_loc, a_bsc = "114", "middle ISO 7 buffer room", "E001798" 
    
    env_desc = f"The cleanroom used for Celsis sterility testing (Suite {t_suite}) and the aliquoting step (Suite {a_suite}A) comprise interconnected ISO 7 and ISO 8 sections maintaining unidirectional airflow."
    
    if st.session_state.analyst_name == st.session_state.aliquoting_name:
        action_desc = f"Sample processing and the aliquoting step were conducted in BSC E00{st.session_state.bsc_id} ({t_loc}) and BSC {a_bsc} ({a_loc}), respectively, by {st.session_state.analyst_name} on {st.session_state.test_date} as per SOP 2.600.059."
    else:
        action_desc = f"Sample processing was conducted in BSC E00{st.session_state.bsc_id} ({t_loc}) by Analyst {st.session_state.analyst_name}, and the aliquoting step was performed in BSC {a_bsc} ({a_loc}) by Aliquoting Analyst {st.session_state.aliquoting_name} on {st.session_state.test_date} as per SOP 2.600.059."
        
    return f"{env_desc}\n\n{action_desc}"

def generate_celsis_narrative_and_details():
    def is_fail(val): return val and str(val).strip().lower() != "no growth"

    failures = []
    em_fields = [
        {"cat": "personal sampling", "obs": st.session_state.get("obs_pers", ""), "etx": st.session_state.get("etx_pers", ""), "id": st.session_state.get("id_pers", "")},
        {"cat": "surface sampling", "obs": st.session_state.get("obs_surf", ""), "etx": st.session_state.get("etx_surf", ""), "id": st.session_state.get("id_surf", "")},
        {"cat": "settling plates", "obs": st.session_state.get("obs_sett", ""), "etx": st.session_state.get("etx_sett", ""), "id": st.session_state.get("id_sett", "")}
    ]

    for f in em_fields:
        if is_fail(f['obs']): failures.append(f)

    if not failures:
        narrative = "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, and settling plates during both the processing and aliquoting steps. Additionally, weekly active air sampling and weekly surface sampling showed no microbial growth."
        details = ""
    else:
        narrative = "Microbial growth was observed during the environmental monitoring of the processing and aliquoting steps."
        detail_parts = []
        for i, f in enumerate(failures):
            is_plural = "s" in str(f['obs']).lower() or (re.search(r'\d+', str(f['obs'])) and int(re.search(r'\d+', str(f['obs'])).group()) > 1)
            verb = "were" if is_plural else "was"
            noun = "organisms were" if is_plural else "organism was"
            prefix = "Specifically, " if i == 0 else "Additionally, "
            sentence = f"{prefix}{f['obs']} {verb} detected during {f['cat']} and {verb} submitted for microbial identification under sample ID {f['etx']}, where the {noun} identified as {f['id']}."
            detail_parts.append(sentence)
        details = " ".join(detail_parts)

    return narrative, details

def generate_celsis_history_text():
    if st.session_state.get("incidence_count", 0) == 0 or st.session_state.get("has_prior_failures") == "No":
        phrase = "no prior failures"
    else:
        count = st.session_state.get("incidence_count", 0)
        pids = [st.session_state.get(f"prior_oos_{i}", "").strip() for i in range(count) if st.session_state.get(f"prior_oos_{i}")]
        if not pids: refs_str = "[Missing OOS References]"
        elif len(pids) == 1: refs_str = pids[0]
        else: refs_str = ", ".join(pids[:-1]) + " and " + pids[-1]
        phrase = f"1 incident ({refs_str})" if len(pids) == 1 else f"{len(pids)} incidents ({refs_str})"

    return f"Analyzing a 6-month sample history for {st.session_state.get('client_name', '[Client]')}, this specific analyte \"{st.session_state.get('sample_name', '[Sample]')}\" has had {phrase} using Celsis sterility testing during this period."

def generate_celsis_cross_contam_text():
    if st.session_state.get("other_positives") == "No": 
        return "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
    
    num = st.session_state.get("total_pos_count_num", 1) - 1
    other_list_ids = []; detail_sentences = []
    for i in range(num):
        oid = st.session_state.get(f"other_id_{i}", "")
        oord_num = st.session_state.get(f"other_order_{i}", 1)
        if oid:
            other_list_ids.append(oid); detail_sentences.append(f"{oid} was the {ordinal(oord_num)} sample processed")
            
    all_ids = other_list_ids + [st.session_state.get("sample_id", "")]
    if not all_ids: ids_str = ""
    elif len(all_ids) == 1: ids_str = all_ids[0]
    else: ids_str = ", ".join(all_ids[:-1]) + " and " + all_ids[-1]
    
    count_word = num_to_words(st.session_state.get("total_pos_count_num", 1))
    cur_ord_text = ordinal(st.session_state.get("current_pos_order", 1))
    current_detail = f"while {st.session_state.get('sample_id', '')} was the {cur_ord_text}"
    
    if len(detail_sentences) == 1: details_str = f"{detail_sentences[0]}, {current_detail}"
    else: details_str = ", ".join(detail_sentences) + f", {current_detail}"
    
    return f"{ids_str} were the {count_word} samples tested positive for microbial growth. The analyst confirmed that these samples were not processed concurrently, sequentially, or within the same manifold run. Specifically, {details_str}. The analyst also verified that gloves were thoroughly disinfected between samples. Furthermore, all other samples processed by the analyst that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
