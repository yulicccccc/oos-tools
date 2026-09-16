import re

def refactor_celsis_logic(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # The block we need to replace is from `def generate_celsis_narrative_and_details():` down to `return narrative, details, smart_just`
    start_marker = r'def generate_celsis_narrative_and_details\(\):'
    end_marker = r'return narrative, details, smart_just'
    
    pattern = re.compile(f"({start_marker}.*?{end_marker})", re.DOTALL)

    replacement = """def generate_celsis_narrative_and_details():
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

        return f"Environmental Monitoring from Celsis Sterility {phase_title.capitalize()}: {daily_str}\\n\\n{weekly_str}"

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

    return em_pro_narrative, em_alq_narrative, smart_just"""

    new_content = re.sub(pattern, replacement, content)
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(new_content)
        
    print("Updated celsis_logic.py")

refactor_celsis_logic(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\Eagle\oos-tools\celsis_logic.py")
