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
    import streamlit as st
    import re
    def any_fail(*keys): return any(str(st.session_state.get(k, 'No growth')).lower() != 'no growth' and str(st.session_state.get(k, 'No growth')).strip() != '' for k in keys)
    def first_fail(variants):
        for v in variants:
            obs = str(st.session_state.get(v[0], 'No growth')).lower()
            if obs != 'no growth' and obs.strip() != '':
                return (st.session_state.get(v[0]), st.session_state.get(v[1]), st.session_state.get(v[2]), v[3], v[4])
        return None

    def get_phase_text(p): return "processing" if p == "pro_" else "aliquoting"
    def get_daily_time(d): return "the date before testing" if d == "be_" else "the date after testing" if d == "af_" else "the date of testing"
    def get_weekly_time(d): return "the week prior to testing" if d == "be_" else "the week after testing" if d == "af_" else "the week of testing"

    # Separate processing and aliquoting variants
    pro_pers = [(f"pro_{d}obs_pers", f"pro_{d}etx_pers", f"pro_{d}id_pers", "processing", get_daily_time(d)) for d in ["be_","on_","af_"]]
    pro_surf = [(f"pro_{d}obs_surf", f"pro_{d}etx_surf", f"pro_{d}id_surf", "processing", get_daily_time(d)) for d in ["be_","on_","af_"]]
    pro_sett = [(f"pro_{d}obs_sett", f"pro_{d}etx_sett", f"pro_{d}id_sett", "processing", get_daily_time(d)) for d in ["be_","on_","af_"]]
    pro_air  = [(f"pro_{d}obs_air_wk", f"pro_{d}etx_air_wk", f"pro_{d}id_air_wk", "processing", get_weekly_time(d)) for d in ["be_","on_","af_"]]
    pro_room = [(f"pro_{d}obs_room_wk", f"pro_{d}etx_room_wk", f"pro_{d}id_room_wk", "processing", get_weekly_time(d)) for d in ["be_","on_","af_"]]

    alq_pers = [(f"alq_{d}obs_pers", f"alq_{d}etx_pers", f"alq_{d}id_pers", "aliquoting", get_daily_time(d)) for d in ["be_","on_","af_"]]
    alq_surf = [(f"alq_{d}obs_surf", f"alq_{d}etx_surf", f"alq_{d}id_surf", "aliquoting", get_daily_time(d)) for d in ["be_","on_","af_"]]
    alq_sett = [(f"alq_{d}obs_sett", f"alq_{d}etx_sett", f"alq_{d}id_sett", "aliquoting", get_daily_time(d)) for d in ["be_","on_","af_"]]
    alq_air  = [(f"alq_{d}obs_air_wk", f"alq_{d}etx_air_wk", f"alq_{d}id_air_wk", "aliquoting", get_weekly_time(d)) for d in ["be_","on_","af_"]]
    alq_room = [(f"alq_{d}obs_room_wk", f"alq_{d}etx_room_wk", f"alq_{d}id_room_wk", "aliquoting", get_weekly_time(d)) for d in ["be_","on_","af_"]]
    
    def generate_phase_narrative(phase_title, pers, surf, sett, air, room, analyst_init, bsc_id):
        all_daily = pers + surf + sett
        all_weekly = air + room
        
        daily_fails = []
        for v in all_daily:
            obs = str(st.session_state.get(v[0], 'No growth')).lower()
            if obs != 'no growth' and obs.strip() != '':
                daily_fails.append(v)
                
        weekly_fails = []
        for v in all_weekly:
            obs = str(st.session_state.get(v[0], 'No growth')).lower()
            if obs != 'no growth' and obs.strip() != '':
                weekly_fails.append(v)

        if not daily_fails:
            daily_str = f"After reviewing the Environmental Monitoring results for the relevant testing dates, no microbial growth was detected on the personnel monitoring plates (analyst {analyst_init}) or on the ISO 5 BSC (E00{bsc_id}) settling and surface plates for the date of testing, the preceding date, or the subsequent date."
        else:
            daily_str = f"After reviewing the Environmental Monitoring results for the relevant testing dates, microbial growth was detected during daily sampling."
            for v in daily_fails:
                cat = "personnel sampling" if "pers" in v[0] else "surface sampling" if "surf" in v[0] else "settling plates"
                daily_str += f" Specifically, on {v[4]}, {st.session_state.get(v[0])} was detected on {cat}. The organism was submitted under ID {st.session_state.get(v[1])} and identified as {st.session_state.get(v[2])}."

        if not weekly_fails:
            weekly_str = "No growth was observed on weekly surface and active air sampling plates for either the week prior to testing or the week of testing."
        else:
            weekly_str = "However, microbial growth was observed during weekly sampling."
            for v in weekly_fails:
                cat = "active air sampling" if "air" in v[0] else "surface sampling"
                weekly_str += f" During {v[4]}, {st.session_state.get(v[0])} was detected on weekly {cat} plates. The organism was submitted under ID {st.session_state.get(v[1])} and identified as {st.session_state.get(v[2])}."

        return f"Environmental Monitoring from Celsis Sterility {phase_title.capitalize()}: {daily_str}\n\n{weekly_str}"

    a_init = st.session_state.get('analyst_initial', '').strip()
    alq_init = st.session_state.get('aliquoting_initial', '').strip()
    pro_bsc = st.session_state.get('bsc_id', '').strip()
    alq_bsc = "1798"
    
    em_pro_narrative = generate_phase_narrative("Processing", pro_pers, pro_surf, pro_sett, pro_air, pro_room, a_init, pro_bsc)
    em_alq_narrative = generate_phase_narrative("Aliquoting", alq_pers, alq_surf, alq_sett, alq_air, alq_room, alq_init, alq_bsc)
    
    # Collect ALL failures for the Smart Justification Engine
    failures = []
    all_vars = pro_pers + pro_surf + pro_sett + pro_air + pro_room + alq_pers + alq_surf + alq_sett + alq_air + alq_room
    for v in all_vars:
        obs = str(st.session_state.get(v[0], 'No growth')).lower()
        if obs != 'no growth' and obs.strip() != '':
            time_type = 'daily' if v in pro_pers+pro_surf+pro_sett+alq_pers+alq_surf+alq_sett else 'weekly'
            failures.append({"id": st.session_state.get(v[2], ""), "time": time_type, "timing": v[4]})

    # SMART JUSTIFICATION ENGINE
    smart_just = ""
    positive_org = st.session_state.get("positive_org", "N/A").strip()
    
    if not failures:
        smart_just = "Based on the observations outlined above, the cleanroom environment was in optimal condition with no microbial growth detected. Therefore, it is highly unlikely that the failing results were due to reagents, supplies, the cleanroom environment, the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is minimal, and the original result is deemed to be valid."
    else:
        just_parts = []
        
        all_em_ids = [f['id'].lower() for f in failures]
        if positive_org.lower() not in all_em_ids and "pending" not in positive_org.lower():
            just_parts.append(f"Notably, the colony morphology of all microorganisms recovered from the processing cleanroom environments differed from that of the microorganism isolated from the test sample ({positive_org}). This observation indicates that the environmental monitoring findings and the test sample contamination were likely isolated and unrelated events.")
        
        has_weekly = any(f['time'] == 'weekly' for f in failures)
        if has_weekly:
            just_parts.append("Also, while microbial growth was detected during weekly monitoring, it is important to note that these organisms were detected in the ISO 8 background room environment, whereas the sample manipulation occurred strictly within the ISO 5 primary engineering control.")
            
        has_daily_testing_day_failure = any(f['time'] == 'daily' and 'of testing' in f['timing'].lower() for f in failures)
        if not has_daily_testing_day_failure:
            just_parts.append("Also, the absence of contamination on analyst glove plates and work surface monitoring indicates that no viable transfer pathway existed from the ISO 8 areas to the ISO 5 BSCs where processing and aliquoting were performed.")
            
        just_parts.append("Furthermore, the lack of contamination in other samples supports the fact that the testing environment was operating under optimal conditions.")
        smart_just = "\\n\\n".join(just_parts)

    return em_pro_narrative, em_alq_narrative, smart_just

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
