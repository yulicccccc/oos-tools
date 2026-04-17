# filename: celsis_logic.py
import streamlit as st
import re
from utils import get_room_logic as u_grl, get_full_name, ordinal, num_to_words

# --- 1. CONFIG & KEYS (前后端数据契约) ---
FIELD_KEYS = [
    "oos_id", "client_name", "sample_id", "test_date", "process_date", "sample_name", "lot_number", 
    "dosage_form", "monthly_cleaning_date", 
    "prepper_initial", "prepper_name", "analyst_initial", "analyst_name",
    "aliquoting_initial", "aliquoting_name", 
    "bsc_id", "celsis_id", "test_record",
    "positive_media", "positive_id", "positive_org",
    "control_lot", "control_data",
    "incidence_count", "has_prior_failures",
    "other_positives", "total_pos_count_num", "current_pos_order",
    "pos_bottle_count"
]
for i in range(10):
    FIELD_KEYS.extend([f"pos_media_{i}", f"pos_id_{i}", f"pos_org_{i}"])
for i in range(20):
    FIELD_KEYS.extend([f"other_id_{i}", f"other_order_{i}", f"prior_oos_{i}"])

# --- 2. HELPER FUNCTIONS (安全锁) ---
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
        "OOS Number": "oos_id", "Client Name": "client_name", "Sample ID": "sample_id", 
        "Test Date": "test_date", "Process Date": "process_date", "Sample Name": "sample_name", 
        "Lot Number": "lot_number", "Prepper Name": "prepper_name", "Processor Name": "analyst_name",
        "Aliquoting Name": "aliquoting_name", "Processing BSC ID": "bsc_id", 
        "Celsis ID": "celsis_id"
    }
    for label, key in reqs.items():
        if not st.session_state.get(key, "").strip(): warnings.append(label)
            
    for date_key in ["test_date", "process_date"]:
        d_val = st.session_state.get(date_key, "").strip()
        if d_val:
            try: 
                from datetime import datetime
                datetime.strptime(d_val, "%d%b%y")
            except ValueError: 
                errors.append(f"❌ Date Error: '{d_val}' invalid. Use DDMMMYY (e.g. 17Mar26).")
    return errors, warnings

# --- 3. TEXT GENERATION LOGIC (重型文案生成引擎) ---

def generate_celsis_equipment_text():
    t_room, t_suite, t_suffix, t_loc = u_grl(st.session_state.bsc_id)
    a_room, a_suite, a_suffix, a_loc = u_grl("1798")
    a_bsc = "1798"
    
    p_date = st.session_state.get("process_date", "[Process Date]")
    t_date = st.session_state.get("test_date", "[Test Date]")

    if str(st.session_state.bsc_id).strip() == a_bsc:
        part1 = f"The cleanroom used for testing and the aliquoting step (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        part2 = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both sample processing and aliquoting steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{st.session_state.bsc_id} was certified and approved by both the Engineering and Quality Assurance teams."
        if st.session_state.analyst_name == st.session_state.aliquoting_name:
            usage_sent = f"Sample processing and the aliquoting step were conducted in the ISO 5 BSC E00{st.session_state.bsc_id} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {st.session_state.analyst_name} on {p_date} and {t_date}, respectively."
        else:
            usage_sent = f"Sample processing was conducted in the ISO 5 BSC E00{st.session_state.bsc_id} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {st.session_state.analyst_name} on {p_date}, and the aliquoting step was conducted by {st.session_state.aliquoting_name} on {t_date}."
        return f"{part1}\n\n{part2} {usage_sent}"
    else:
        if t_suite == a_suite:
             part1 = f"The cleanroom used for testing and the aliquoting step (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        else:
             p1a = f"The cleanroom used for testing (E00{t_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which opens into the middle ISO 7 buffer room ({t_suite}A), and then into the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
             p1b = f"The cleanroom used for the aliquoting step (E00{a_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({a_suite}B), which opens into the middle ISO 7 buffer room ({a_suite}A), and then into the outermost ISO 8 anteroom ({a_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {a_suite}B through {a_suite}A and into {a_suite}."
             part1 = f"{p1a}\n\n{p1b}"
        intro = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), and ISO 5 BSC E00{a_bsc}, located in the {a_loc}, (Suite {a_suite}{a_suffix}), were thoroughly cleaned and disinfected prior to their respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Furthermore, the BSCs used throughout testing, E00{st.session_state.bsc_id} for sample processing and E00{a_bsc} for the aliquoting step, were certified and approved by both the Engineering and Quality Assurance teams."
        if st.session_state.analyst_name == st.session_state.aliquoting_name:
            usage_sent = f"Sample processing was conducted within the ISO 5 BSC in the {t_loc} (Suite {t_suite}{t_suffix}, BSC E00{st.session_state.bsc_id}) on {p_date}, and the aliquoting step was conducted within the ISO 5 BSC in the {a_loc} (Suite {a_suite}{a_suffix}, BSC E00{a_bsc}) on {t_date} by {st.session_state.analyst_name}."
        else:
            usage_sent = f"Sample processing was conducted within the ISO 5 BSC in the {t_loc} (Suite {t_suite}{t_suffix}, BSC E00{st.session_state.bsc_id}) by {st.session_state.analyst_name} on {p_date}, and the aliquoting step was conducted within the ISO 5 BSC in the {a_loc} (Suite {a_suite}{a_suffix}, BSC E00{a_bsc}) by {st.session_state.aliquoting_name} on {t_date}."
        return f"{part1}\n\n{intro} {usage_sent}"

def generate_celsis_narrative_and_details():
    default_obs, default_etx, default_id = "No Growth", "N/A", "N/A"
    fixed_map = {
        "Personnel Obs": ("obs_pers", "etx_pers", "id_pers"), 
        "Surface Obs": ("obs_surf", "etx_surf", "id_surf"), 
        "Settling Obs": ("obs_sett", "etx_sett", "id_sett"), 
        "Weekly Air Obs": ("obs_air", "etx_air_weekly", "id_air_weekly"), 
        "Weekly Surf Obs": ("obs_room", "etx_room_weekly", "id_room_wk_of")
    }
    for cat, (k_obs, k_etx, k_id) in fixed_map.items():
        st.session_state[k_obs] = default_obs; st.session_state[k_etx] = default_etx; st.session_state[k_id] = default_id

    if st.session_state.get("em_growth_observed") == "Yes":
        count = st.session_state.get("em_growth_count", 1)
        for i in range(count):
            cat = st.session_state.get(f"em_cat_{i}"); obs = st.session_state.get(f"em_obs_{i}", "")
            etx = st.session_state.get(f"em_etx_{i}", ""); mid = st.session_state.get(f"em_id_{i}", "")
            if cat in fixed_map:
                k_obs, k_etx, k_id = fixed_map[cat]
                st.session_state[k_obs] = obs; st.session_state[k_etx] = etx; st.session_state[k_id] = mid

    def is_fail(val): return val and str(val).strip().lower() != "no growth"
    failures = []
    if is_fail(st.session_state.obs_pers): failures.append({"cat": "personnel sampling", "obs": st.session_state.obs_pers, "etx": st.session_state.etx_pers, "id": st.session_state.id_pers, "time": "daily"})
    if is_fail(st.session_state.obs_surf): failures.append({"cat": "surface sampling", "obs": st.session_state.obs_surf, "etx": st.session_state.etx_surf, "id": st.session_state.id_surf, "time": "daily"})
    if is_fail(st.session_state.obs_sett): failures.append({"cat": "settling plates", "obs": st.session_state.obs_sett, "etx": st.session_state.etx_sett, "id": st.session_state.id_sett, "time": "daily"})
    if is_fail(st.session_state.obs_air): failures.append({"cat": "weekly active air sampling", "obs": st.session_state.obs_air, "etx": st.session_state.etx_air_weekly, "id": st.session_state.id_air_weekly, "time": "weekly"})
    if is_fail(st.session_state.obs_room): failures.append({"cat": "weekly surface sampling", "obs": st.session_state.obs_room, "etx": st.session_state.etx_room_weekly, "id": st.session_state.id_room_wk_of, "time": "weekly"})

    pass_daily_clean, pass_wk_clean = [], []
    if not is_fail(st.session_state.obs_pers): pass_daily_clean.append("personnel sampling (left touch and right touch)")
    if not is_fail(st.session_state.obs_surf): pass_daily_clean.append("surface sampling")
    if not is_fail(st.session_state.obs_sett): pass_daily_clean.append("settling plates")
    if not is_fail(st.session_state.obs_air): pass_wk_clean.append("weekly active air sampling")
    if not is_fail(st.session_state.obs_room): pass_wk_clean.append("weekly surface sampling")

    narr_parts = []
    if pass_daily_clean:
        d_str = f"{pass_daily_clean[0]}" if len(pass_daily_clean)==1 else f"{pass_daily_clean[0]} and {pass_daily_clean[1]}" if len(pass_daily_clean)==2 else f"{pass_daily_clean[0]}, {pass_daily_clean[1]}, and {pass_daily_clean[2]}"
        narr_parts.append(f"no microbial growth was observed in {d_str} during both the processing and aliquoting steps")
    if pass_wk_clean:
        w_str = f"{pass_wk_clean[0]}" if len(pass_wk_clean)==1 else f"{pass_wk_clean[0]} and {pass_wk_clean[1]}" if len(pass_wk_clean)==2 else ", ".join(pass_wk_clean)
        narr_parts.append(f"Additionally, {w_str} showed no microbial growth" if narr_parts else f"no microbial growth was observed in {w_str}")
    
    narrative = "Upon analyzing the environmental monitoring results, " + ". ".join(narr_parts) + "." if narr_parts else "Upon analyzing the environmental monitoring results, microbial growth was observed in all sampled areas."

    details = ""
    if failures:
        daily_fails = [f["cat"] for f in failures if f['time'] == 'daily']
        weekly_fails = [f["cat"] for f in failures if f['time'] == 'weekly']
        intro_parts = []
        if daily_fails: intro_parts.append(f"{' and '.join(daily_fails)} during the processing and aliquoting steps")
        if weekly_fails: intro_parts.append(f"{' and '.join(weekly_fails)}")
        fail_intro = f"However, microbial growth was observed during { ' and '.join(intro_parts) }." if intro_parts else ""
        
        detail_sentences = []
        for i, f in enumerate(failures):
            is_plural = "s" in str(f['obs']).lower() or (re.search(r'\d+', str(f['obs'])) and int(re.search(r'\d+', str(f['obs'])).group()) > 1)
            verb = "were" if is_plural else "was"; noun = "organisms were" if is_plural else "organism was"
            prefix = "Specifically, " if i == 0 else "Additionally, "
            detail_sentences.append(f"{prefix}{f['obs']} {verb} detected during {f['cat']} and {verb} submitted for microbial identification under sample ID {f['etx']}, where the {noun} identified as {f['id']}.")
        details = f"{fail_intro} {' '.join(detail_sentences)}"

    return narrative, details

def generate_celsis_history_text():
    if st.session_state.get("incidence_count", 0) == 0 or st.session_state.get("has_prior_failures") == "No": phrase = "no prior failures"
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
    other_list_ids, detail_sentences = [], []
    for i in range(num):
        oid = st.session_state.get(f"other_id_{i}", ""); oord_num = st.session_state.get(f"other_order_{i}", 1)
        if oid: other_list_ids.append(oid); detail_sentences.append(f"{oid} was the {ordinal(oord_num)} sample processed")
    all_ids = other_list_ids + [st.session_state.get("sample_id", "")]
    if not all_ids: ids_str = ""
    elif len(all_ids) == 1: ids_str = all_ids[0]
    else: ids_str = ", ".join(all_ids[:-1]) + " and " + all_ids[-1]
    count_word = num_to_words(st.session_state.get("total_pos_count_num", 1))
    cur_ord_text = ordinal(st.session_state.get("current_pos_order", 1))
    current_detail = f"while {st.session_state.get('sample_id', '')} was the {cur_ord_text}"
    details_str = f"{detail_sentences[0]}, {current_detail}" if len(detail_sentences) == 1 else ", ".join(detail_sentences) + f", {current_detail}"
    return f"{ids_str} were the {count_word} samples tested positive for microbial growth. The analyst confirmed that these samples were not processed concurrently, sequentially, or within the same manifold run. Specifically, {details_str}. The analyst also verified that gloves were thoroughly disinfected between samples. Furthermore, all other samples processed by the analyst that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
