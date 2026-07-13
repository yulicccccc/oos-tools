# filename: celsis_logic.py
import streamlit as st
import re
from utils import get_room_logic as u_grl, get_full_name, ordinal, num_to_words, get_cleanroom_narrative

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
    "pos_bottle_count", "em_growth_observed", "em_growth_count",
]
# Add all prefixed EM keys for session save/restore
for _phase in ["pro_", "alq_"]:
    for _em in ["pers", "surf", "sett"]:
        for _day in ["be_", "", "af_"]:
            FIELD_KEYS.extend([f"{_phase}{_day}obs_{_em}", f"{_phase}{_day}etx_{_em}", f"{_phase}{_day}id_{_em}"])
    for _wk in ["air_wk", "air_wk2", "room_wk", "room_wk2"]:
        FIELD_KEYS.extend([f"{_phase}obs_{_wk}", f"{_phase}etx_{_wk}", f"{_phase}id_{_wk}"])
for i in range(10):
    FIELD_KEYS.extend([f"pos_media_{i}", f"pos_id_{i}", f"pos_org_{i}", f"em_cat_{i}", f"em_obs_{i}", f"em_etx_{i}", f"em_id_{i}"])
for i in range(20):
    FIELD_KEYS.extend([f"other_id_{i}", f"other_order_{i}", f"prior_oos_{i}"])

# --- 2. HELPER FUNCTIONS ---
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
        if not st.session_state.get(key, "").strip(): 
            warnings.append(label)
            
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
    """
    根据标准话术 (SOP 像素级复刻):
    1. 动态拆解 Cleanroom 结构。
    2. 包含清洗、认证、时间、人员。
    3. 末尾加入绝杀的 "as per SOP 2.600.059."。
    """
    t_room, t_suite, t_suffix, t_loc = u_grl(st.session_state.bsc_id)
    a_room, a_suite, a_suffix, a_loc = u_grl("1798")
    a_bsc = "1798"
    
    p_date = st.session_state.get("process_date", "[Process Date]")
    t_date = st.session_state.get("test_date", "[Test Date]")
    
    analyst = st.session_state.get("analyst_name", "[Processor Name]")
    aliquoter = st.session_state.get("aliquoting_name", "[Aliquoting Name]")

    t_suite_phrase = f"Suite {t_suite}{t_suffix}" if t_suite != "L-Suite" else "L-Suite"
    a_suite_phrase = f"Suite {a_suite}{a_suffix}" if a_suite != "L-Suite" else "L-Suite"

    if t_suite == a_suite:
        part1 = get_cleanroom_narrative(t_suite, action_text="processing and aliquoting procedures", verb="comprises")
    else:
        p1a = get_cleanroom_narrative(t_suite, action_text="processing procedures", verb="comprises")
        p1b = get_cleanroom_narrative(a_suite, action_text="aliquoting procedures", verb="comprises")
        part1 = f"{p1a}\n\n{p1b}"

    bsc_id_str = str(st.session_state.bsc_id).strip()
    
    if bsc_id_str == a_bsc:
        part2 = f"The ISO 5 BSC E00{bsc_id_str}, located in the {t_loc}, ({t_suite_phrase}), was used for both sample processing and aliquoting steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{bsc_id_str} was certified and approved by both the Engineering and Quality Assurance teams."
        
        # --- 在这里加入了绝杀的 as per SOP 2.600.059 ---
        if analyst == aliquoter:
            usage_sent = f"Sample processing and aliquoting were conducted in the ISO 5 BSC E00{bsc_id_str} in the {t_loc}, ({t_suite_phrase}) by {analyst} on {p_date} and {t_date}, respectively, as per SOP 2.600.059."
        else:
            usage_sent = f"Sample processing was conducted in the ISO 5 BSC E00{bsc_id_str} in the {t_loc}, ({t_suite_phrase}) by {analyst} on {p_date}, and the aliquoting step was conducted in the ISO 5 BSC E00{bsc_id_str} in the {t_loc}, ({t_suite_phrase}) by {aliquoter} on {t_date} as per SOP 2.600.059."
            
        return f"{part1}\n\n{part2} {usage_sent}"
        
    else:
        part2 = f"The ISO 5 BSC E00{bsc_id_str}, located in the {t_loc}, ({t_suite_phrase}), and ISO 5 BSC E00{a_bsc}, located in the {a_loc}, ({a_suite_phrase}), were thoroughly cleaned and disinfected prior to their respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, the BSCs used throughout testing, E00{bsc_id_str} for sample processing and E00{a_bsc} for the aliquoting step, were certified and approved by both the Engineering and Quality Assurance teams."
        
        # --- 在这里加入了绝杀的 as per SOP 2.600.059 ---
        usage_sent = f"Sample processing was conducted in the ISO 5 BSC E00{bsc_id_str} in the {t_loc}, ({t_suite_phrase}) by {analyst} on {p_date}, and the aliquoting step was conducted in the ISO 5 BSC E00{a_bsc} in the {a_loc}, ({a_suite_phrase}) by {aliquoter} on {t_date} as per SOP 2.600.059."
        
        return f"{part1}\n\n{part2} {usage_sent}"

def generate_celsis_narrative_and_details():
    default_obs, default_etx, default_id = "No Growth", "N/A", "N/A"
    
    # Expanded mapping: category label → (obs_key, etx_key, id_key) in session_state
    # Pro = Processing phase, Alq = Aliquoting phase
    # Daily EM: be_ = before testing, (none) = date of testing, af_ = after testing
    # Weekly EM: wk = before testing date, wk2 = on/after testing date
    fixed_map = {
        # --- Processing Daily ---
        "Pro: Personnel (Before Testing)": ("pro_be_obs_pers", "pro_be_etx_pers", "pro_be_id_pers"),
        "Pro: Personnel (Date of Testing)": ("pro_obs_pers", "pro_etx_pers", "pro_id_pers"),
        "Pro: Personnel (After Testing)": ("pro_af_obs_pers", "pro_af_etx_pers", "pro_af_id_pers"),
        "Pro: Surface (Before Testing)": ("pro_be_obs_surf", "pro_be_etx_surf", "pro_be_id_surf"),
        "Pro: Surface (Date of Testing)": ("pro_obs_surf", "pro_etx_surf", "pro_id_surf"),
        "Pro: Surface (After Testing)": ("pro_af_obs_surf", "pro_af_etx_surf", "pro_af_id_surf"),
        "Pro: Settling (Before Testing)": ("pro_be_obs_sett", "pro_be_etx_sett", "pro_be_id_sett"),
        "Pro: Settling (Date of Testing)": ("pro_obs_sett", "pro_etx_sett", "pro_id_sett"),
        "Pro: Settling (After Testing)": ("pro_af_obs_sett", "pro_af_etx_sett", "pro_af_id_sett"),
        # --- Processing Weekly ---
        "Pro: Weekly Air (Before Testing Date)": ("pro_obs_air_wk", "pro_etx_air_wk", "pro_id_air_wk"),
        "Pro: Weekly Air (On/After Testing Date)": ("pro_obs_air_wk2", "pro_etx_air_wk2", "pro_id_air_wk2"),
        "Pro: Weekly Surf (Before Testing Date)": ("pro_obs_room_wk", "pro_etx_room_wk", "pro_id_room_wk"),
        "Pro: Weekly Surf (On/After Testing Date)": ("pro_obs_room_wk2", "pro_etx_room_wk2", "pro_id_room_wk2"),
        # --- Aliquoting Daily ---
        "Alq: Personnel (Before Testing)": ("alq_be_obs_pers", "alq_be_etx_pers", "alq_be_id_pers"),
        "Alq: Personnel (Date of Testing)": ("alq_obs_pers", "alq_etx_pers", "alq_id_pers"),
        "Alq: Personnel (After Testing)": ("alq_af_obs_pers", "alq_af_etx_pers", "alq_af_id_pers"),
        "Alq: Surface (Before Testing)": ("alq_be_obs_surf", "alq_be_etx_surf", "alq_be_id_surf"),
        "Alq: Surface (Date of Testing)": ("alq_obs_surf", "alq_etx_surf", "alq_id_surf"),
        "Alq: Surface (After Testing)": ("alq_af_obs_surf", "alq_af_etx_surf", "alq_af_id_surf"),
        "Alq: Settling (Before Testing)": ("alq_be_obs_sett", "alq_be_etx_sett", "alq_be_id_sett"),
        "Alq: Settling (Date of Testing)": ("alq_obs_sett", "alq_etx_sett", "alq_id_sett"),
        "Alq: Settling (After Testing)": ("alq_af_obs_sett", "alq_af_etx_sett", "alq_af_id_sett"),
        # --- Aliquoting Weekly ---
        "Alq: Weekly Air (Before Testing Date)": ("alq_obs_air_wk", "alq_etx_air_wk", "alq_id_air_wk"),
        "Alq: Weekly Air (On/After Testing Date)": ("alq_obs_air_wk2", "alq_etx_air_wk2", "alq_id_air_wk2"),
        "Alq: Weekly Surf (Before Testing Date)": ("alq_obs_room_wk", "alq_etx_room_wk", "alq_id_room_wk"),
        "Alq: Weekly Surf (On/After Testing Date)": ("alq_obs_room_wk2", "alq_etx_room_wk2", "alq_id_room_wk2"),
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
    
    # Aggregate: for narrative purposes, check if ANY timing variant across both phases has growth for each EM type
    def any_fail(*keys):
        """Return True if any of the given session_state keys has a non-'No Growth' value."""
        return any(is_fail(st.session_state.get(k, "No Growth")) for k in keys)
    
    def first_fail(*keys):
        """Return the first failing (obs, etx, id) tuple from the given key groups."""
        for k_obs, k_etx, k_id in keys:
            if is_fail(st.session_state.get(k_obs, "No Growth")):
                return st.session_state.get(k_obs), st.session_state.get(k_etx, "N/A"), st.session_state.get(k_id, "N/A")
        return None
    
    # All obs keys for each EM type across both phases
    pers_obs_keys = [f"{p}{d}obs_pers" for p in ["pro_","alq_"] for d in ["be_","","af_"]]
    surf_obs_keys = [f"{p}{d}obs_surf" for p in ["pro_","alq_"] for d in ["be_","","af_"]]
    sett_obs_keys = [f"{p}{d}obs_sett" for p in ["pro_","alq_"] for d in ["be_","","af_"]]
    air_obs_keys = [f"{p}obs_air_wk{s}" for p in ["pro_","alq_"] for s in ["","2"]]
    room_obs_keys = [f"{p}obs_room_wk{s}" for p in ["pro_","alq_"] for s in ["","2"]]
    
    failures = []
    if any_fail(*pers_obs_keys):
        f = first_fail(*[(f"{p}{d}obs_pers", f"{p}{d}etx_pers", f"{p}{d}id_pers") for p in ["pro_","alq_"] for d in ["be_","","af_"]])
        if f: failures.append({"cat": "personnel sampling", "obs": f[0], "etx": f[1], "id": f[2], "time": "daily"})
    if any_fail(*surf_obs_keys):
        f = first_fail(*[(f"{p}{d}obs_surf", f"{p}{d}etx_surf", f"{p}{d}id_surf") for p in ["pro_","alq_"] for d in ["be_","","af_"]])
        if f: failures.append({"cat": "surface sampling", "obs": f[0], "etx": f[1], "id": f[2], "time": "daily"})
    if any_fail(*sett_obs_keys):
        f = first_fail(*[(f"{p}{d}obs_sett", f"{p}{d}etx_sett", f"{p}{d}id_sett") for p in ["pro_","alq_"] for d in ["be_","","af_"]])
        if f: failures.append({"cat": "settling plates", "obs": f[0], "etx": f[1], "id": f[2], "time": "daily"})
    if any_fail(*air_obs_keys):
        f = first_fail(*[(f"{p}obs_air_wk{s}", f"{p}etx_air_wk{s}", f"{p}id_air_wk{s}") for p in ["pro_","alq_"] for s in ["","2"]])
        if f: failures.append({"cat": "weekly active air sampling", "obs": f[0], "etx": f[1], "id": f[2], "time": "weekly"})
    if any_fail(*room_obs_keys):
        f = first_fail(*[(f"{p}obs_room_wk{s}", f"{p}etx_room_wk{s}", f"{p}id_room_wk{s}") for p in ["pro_","alq_"] for s in ["","2"]])
        if f: failures.append({"cat": "weekly surface sampling", "obs": f[0], "etx": f[1], "id": f[2], "time": "weekly"})

    pass_daily_clean, pass_wk_clean = [], []
    if not any_fail(*pers_obs_keys): pass_daily_clean.append("personnel sampling (left touch and right touch)")
    if not any_fail(*surf_obs_keys): pass_daily_clean.append("surface sampling")
    if not any_fail(*sett_obs_keys): pass_daily_clean.append("settling plates")
    if not any_fail(*air_obs_keys): pass_wk_clean.append("weekly active air sampling")
    if not any_fail(*room_obs_keys): pass_wk_clean.append("weekly surface sampling")

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
    other_list_ids, detail_sentences = [], []
    for i in range(num):
        oid = st.session_state.get(f"other_id_{i}", "")
        oord_num = st.session_state.get(f"other_order_{i}", 1)
        if oid: 
            other_list_ids.append(oid)
            detail_sentences.append(f"{oid} was the {ordinal(oord_num)} sample processed")
            
    all_ids = other_list_ids + [st.session_state.get("sample_id", "")]
    if not all_ids: ids_str = ""
    elif len(all_ids) == 1: ids_str = all_ids[0]
    else: ids_str = ", ".join(all_ids[:-1]) + " and " + all_ids[-1]
    
    count_word = num_to_words(st.session_state.get("total_pos_count_num", 1))
    cur_ord_text = ordinal(st.session_state.get("current_pos_order", 1))
    current_detail = f"while {st.session_state.get('sample_id', '')} was the {cur_ord_text}"
    
    details_str = f"{detail_sentences[0]}, {current_detail}" if len(detail_sentences) == 1 else ", ".join(detail_sentences) + f", {current_detail}"
    
    return f"{ids_str} were the {count_word} samples tested positive for microbial growth. The analyst confirmed that these samples were not processed concurrently, sequentially, or within the same manifold run. Specifically, {details_str}. The analyst also verified that gloves were thoroughly disinfected between samples. Furthermore, all other samples processed by the analyst that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
