# filename: scan_logic.py
import streamlit as st
import os
import re
import json
import io
import sys
import subprocess
import time
from datetime import datetime, timedelta
# Import other utils if needed, handling potential errors
try:
    from utils import get_room_logic as u_grl
except ImportError:
    def u_grl(i): return "Unknown", "000", "", "Unknown"

# --- CONFIG & KEYS ---
# 所有的字段名都在这里管理
FIELD_KEYS = [
    "oos_id", "client_name", "sample_id", "test_date", "sample_name", "lot_number", 
    "dosage_form", "monthly_cleaning_date", 
    "prepper_initial", "prepper_name", "analyst_initial", "analyst_name",
    "changeover_initial", "changeover_name", "reader_initial", "reader_name",
    "writer_name", "bsc_id", "chgbsc_id", "scan_id", "shift_number", "active_platform",
    "org_choice", "manual_org", "test_record", "control_pos", "control_lot", "control_exp", 
    "weekly_init", "date_weekly", 
    "equipment_summary", "narrative_summary", "em_details", 
    "sample_history_paragraph", "incidence_count", "oos_refs",
    "other_positives", "cross_contamination_summary",
    "total_pos_count_num", "current_pos_order",
    "diff_changeover_bsc", "has_prior_failures",
    "em_growth_observed", "diff_changeover_analyst",
    "diff_reader_analyst", "em_growth_count",
    "event_number", "confirm_number",
    "obs_pers", "etx_pers", "id_pers", "obs_surf", "etx_surf", "id_surf", 
    "obs_sett", "etx_sett", "id_sett", "obs_air", "etx_air_weekly", "id_air_weekly", 
    "obs_room", "etx_room_weekly", "id_room_wk_of",
    # P2 Keys
    "include_phase2", "retest_date", "retest_sample_id", "retest_result", "retest_scan_id",
    "retest_prepper_name", "retest_prepper_initial", "retest_analyst_name", "retest_analyst_initial",
    "retest_reader_name", "retest_reader_initial", "retest_changeover_name", "retest_changeover_initial",
    "retest_bsc_id", "retest_chgbsc_id", "diff_retest_reader", "diff_retest_changeover", "diff_retest_bsc"
]

# 为 EM 观察点增加动态 key
for i in range(20):
    FIELD_KEYS.extend([f"other_id_{i}", f"other_order_{i}", f"prior_oos_{i}", f"em_cat_{i}", f"em_obs_{i}", f"em_etx_{i}", f"em_id_{i}"])

# --- HELPER FUNCTIONS ---

def ensure_dependencies():
    required = ["docxtpl", "pypdf", "reportlab"]
    missing = []
    for lib in required:
        try: __import__(lib)
        except ImportError: missing.append(lib)
    if missing:
        placeholder = st.empty()
        placeholder.warning(f"⚙️ Installing: {', '.join(missing)}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)
            placeholder.success("Installed! Reloading...")
            time.sleep(1)
            st.rerun()
        except Exception as e: placeholder.error(f"Install failed: {e}")

def get_full_name(initial):
    """Auto-converts initials to full names based on lab personnel."""
    if not initial: return ""
    mapping = {
        "KA": "Kathleen Aruta", "DH": "Domiasha Harrison", "GL": "Guanchen Li", "DS": "Devanshi Shah",
        "QC": "Qiyue Chen", "HS": "Halaina Smith", "MJ": "Mukyung Jang", "AS": "Alex Saravia",
        "CSG": "Clea S. Garza", "RS": "Robin Seymour", "CCD": "Cuong Du", "VV": "Varsha Subramanian",
        "KS": "Karla Silva", "GS": "Gabbie Surber", "PG": "Pagan Gary", "DT": "Debrework Tassew",
        "GA": "Gerald Anyangwe", "MRB": "Muralidhar Bythatagari", "TK": "Tamiru Kotisso",
        "RE": "Rey Estrada", "AO": "Ayomide Odugbesi", "KC": "Kira C"
    }
    return mapping.get(initial.strip().upper(), "") 

def auto_fill_name(initial_key, name_key):
    """Checks if initial changed and updates name if empty."""
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
        "Test Date": "test_date", "Sample Name": "sample_name", "Lot Number": "lot_number",
        "Prepper Name": "prepper_name", "Processor Name": "analyst_name",
        "Reader Name": "reader_name", "Changeover Name": "changeover_name",
        "BSC ID": "bsc_id", "ScanRDI ID": "scan_id",
        "Control Lot": "control_lot", "Control Exp": "control_exp",
        "Events Number": "event_number", "Confirmed #": "confirm_number"
    }
    for label, key in reqs.items():
        if not st.session_state.get(key, "").strip(): warnings.append(label)
    date_val = st.session_state.get("test_date", "").strip()
    if date_val:
        try: datetime.strptime(date_val, "%d%b%y")
        except ValueError: errors.append(f"❌ Date Error: '{date_val}' invalid. Use DDMMMYY (e.g. 07Jan26).")
    return errors, warnings

def clean_filename(text): 
    return re.sub(r'[\\/*?:"<>|]', '_', str(text)).strip() if text else ""

def ordinal(n):
    try:
        n = int(n); return f"{n}th" if 11<=n%100<=13 else f"{n}{['th','st','nd','rd'][n%10 if n%10<=3 else 0]}"
    except: return str(n)

def num_to_words(n):
    return {1:"one",2:"two",3:"three",4:"four",5:"five",6:"six",7:"seven",8:"eight",9:"nine",10:"ten"}.get(n, str(n))

# --- TEXT GENERATION LOGIC ---

def get_room_logic(bsc_id):
    try: return u_grl(bsc_id)
    except: return "Unknown Room", "000", "", "Unknown Loc"

def generate_equipment_text():
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
    if st.session_state.bsc_id == st.session_state.chgbsc_id:
        part1 = f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        part2 = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both testing and changeover steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{st.session_state.bsc_id} was certified and approved by both the Engineering and Quality Assurance teams. Sample processing and changeover were conducted in the ISO 5 BSC E00{st.session_state.bsc_id} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {st.session_state.analyst_name} on {st.session_state.test_date}."
        return f"{part1}\n\n{part2}"
    else:
        if t_suite == c_suite:
             part1 = f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        else:
             p1a = f"The cleanroom used for testing (E00{t_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which opens into the middle ISO 7 buffer room ({t_suite}A), and then into the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
             p1b = f"The cleanroom used for changeover (E00{c_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({c_suite}B), which opens into the middle ISO 7 buffer room ({c_suite}A), and then into the outermost ISO 8 anteroom ({c_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {c_suite}B through {c_suite}A and into {c_suite}."
             part1 = f"{p1a}\n\n{p1b}"
        intro = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), and ISO 5 BSC E00{st.session_state.chgbsc_id}, located in the {c_loc}, (Suite {c_suite}{c_suffix}), were thoroughly cleaned and disinfected prior to their respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Furthermore, the BSCs used throughout testing, E00{st.session_state.bsc_id} for sample processing and E00{st.session_state.chgbsc_id} for the changeover step, were certified and approved by both the Engineering and Quality Assurance teams."
        if st.session_state.analyst_name == st.session_state.changeover_name:
            usage_sent = f"Sample processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{st.session_state.bsc_id}) and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{st.session_state.chgbsc_id}) by {st.session_state.analyst_name} on {st.session_state.test_date}."
        else:
            usage_sent = f"Sample processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{st.session_state.bsc_id}) by {st.session_state.analyst_name} and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{st.session_state.chgbsc_id}) by {st.session_state.changeover_name} on {st.session_state.test_date}."
        return f"{part1}\n\n{intro} {usage_sent}"

def generate_history_text():
    if st.session_state.incidence_count == 0 or st.session_state.has_prior_failures == "No": phrase = "no prior failures"
    else:
        pids = [st.session_state.get(f"prior_oos_{i}","").strip() for i in range(st.session_state.incidence_count) if st.session_state.get(f"prior_oos_{i}")]
        if not pids: refs_str = "..."
        elif len(pids) == 1: refs_str = pids[0]
        else: refs_str = ", ".join(pids[:-1]) + " and " + pids[-1]
        phrase = f"1 incident ({refs_str})" if len(pids) == 1 else f"{len(pids)} incidents ({refs_str})"
    return f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte \"{st.session_state.sample_name}\" has had {phrase} using the Scan RDI method during this period."

def generate_cross_contam_text():
    if st.session_state.other_positives == "No": 
        return "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
    num = st.session_state.total_pos_count_num - 1
    other_list_ids = []; detail_sentences = []
    for i in range(num):
        oid = st.session_state.get(f"other_id_{i}", "")
        oord_num = st.session_state.get(f"other_order_{i}", 1)
        if oid:
            other_list_ids.append(oid); detail_sentences.append(f"{oid} was the {ordinal(oord_num)} sample processed")
    all_ids = other_list_ids + [st.session_state.sample_id]
    if not all_ids: ids_str = ""
    elif len(all_ids) == 1: ids_str = all_ids[0]
    else: ids_str = ", ".join(all_ids[:-1]) + " and " + all_ids[-1]
    count_word = num_to_words(st.session_state.total_pos_count_num)
    cur_ord_text = ordinal(st.session_state.current_pos_order)
    current_detail = f"while {st.session_state.sample_id} was the {cur_ord_text}"
    if len(detail_sentences) == 1: details_str = f"{detail_sentences[0]}, {current_detail}"
    else: details_str = ", ".join(detail_sentences) + f", {current_detail}"
    return f"{ids_str} were the {count_word} samples tested positive for microbial growth. The analyst confirmed that these samples were not processed concurrently, sequentially, or within the same manifold run. Specifically, {details_str}. The analyst also verified that gloves were thoroughly disinfected between samples. Furthermore, all other samples processed by the analyst that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."

def sync_dynamic_to_fixed():
    default_obs, default_etx, default_id = "No Growth", "N/A", "N/A"
    fixed_map = {"Personnel Obs": ("obs_pers", "etx_pers", "id_pers"), "Surface Obs": ("obs_surf", "etx_surf", "id_surf"), "Settling Obs": ("obs_sett", "etx_sett", "id_sett"), "Weekly Air Obs": ("obs_air", "etx_air_weekly", "id_air_weekly"), "Weekly Surf Obs": ("obs_room", "etx_room_weekly", "id_room_wk_of")}
    for cat, (k_obs, k_etx, k_id) in fixed_map.items():
        st.session_state[k_obs] = default_obs; st.session_state[k_etx] = default_etx; st.session_state[k_id] = default_id
    if st.session_state.get("em_growth_observed") == "Yes":
        count = st.session_state.get("em_growth_count", 1)
        for i in range(count):
            cat = st.session_state.get(f"em_cat_{i}"); obs = st.session_state.get(f"em_obs_{i}", ""); etx = st.session_state.get(f"em_etx_{i}", ""); mid = st.session_state.get(f"em_id_{i}", "")
            if cat in fixed_map:
                k_obs, k_etx, k_id = fixed_map[cat]; st.session_state[k_obs] = obs; st.session_state[k_etx] = etx; st.session_state[k_id] = mid

def generate_narrative_and_details():
    sync_dynamic_to_fixed()
    failures = []
    def is_fail(val): return val.strip() and val.strip().lower() != "no growth"
    if is_fail(st.session_state.obs_pers): failures.append({"cat": "personnel sampling", "obs": st.session_state.obs_pers, "etx": st.session_state.etx_pers, "id": st.session_state.id_pers, "time": "daily"})
    if is_fail(st.session_state.obs_surf): failures.append({"cat": "surface sampling", "obs": st.session_state.obs_surf, "etx": st.session_state.etx_surf, "id": st.session_state.id_surf, "time": "daily"})
    if is_fail(st.session_state.obs_sett): failures.append({"cat": "settling plates", "obs": st.session_state.obs_sett, "etx": st.session_state.etx_sett, "id": st.session_state.id_sett, "time": "daily"})
    if is_fail(st.session_state.obs_air): failures.append({"cat": "weekly active air sampling", "obs": st.session_state.obs_air, "etx": st.session_state.etx_air_weekly, "id": st.session_state.id_air_weekly, "time": "weekly"})
    if is_fail(st.session_state.obs_room): failures.append({"cat": "weekly surface sampling", "obs": st.session_state.obs_room, "etx": st.session_state.etx_room_weekly, "id": st.session_state.id_room_wk_of, "time": "weekly"})

    pass_daily_clean = []; pass_wk_clean = []
    if not is_fail(st.session_state.obs_pers): pass_daily_clean.append("personal sampling (left touch and right touch)")
    if not is_fail(st.session_state.obs_surf): pass_daily_clean.append("surface sampling")
    if not is_fail(st.session_state.obs_sett): pass_daily_clean.append("settling plates")
    if not is_fail(st.session_state.obs_air): pass_wk_clean.append("weekly active air sampling")
    if not is_fail(st.session_state.obs_room): pass_wk_clean.append("weekly surface sampling")

    narr_parts = []
    if pass_daily_clean:
        d_str = f"{pass_daily_clean[0]}" if len(pass_daily_clean)==1 else f"{pass_daily_clean[0]} and {pass_daily_clean[1]}" if len(pass_daily_clean)==2 else f"{pass_daily_clean[0]}, {pass_daily_clean[1]}, and {pass_daily_clean[2]}"
        narr_parts.append(f"no microbial growth was observed in {d_str}")
    if pass_wk_clean:
        w_str = f"{pass_wk_clean[0]}" if len(pass_wk_clean)==1 else f"{pass_wk_clean[0]} and {pass_wk_clean[1]}" if len(pass_wk_clean)==2 else ", ".join(pass_wk_clean)
        narr_parts.append(f"Additionally, {w_str} showed no microbial growth" if narr_parts else f"no microbial growth was observed in {w_str}")
    narr = "Upon analyzing the environmental monitoring results, " + ". ".join(narr_parts) + "." if narr_parts else "Upon analyzing the environmental monitoring results, microbial growth was observed in all sampled areas."

    det = ""
    if failures:
        daily_fails = [f["cat"] for f in failures if f['time'] == 'daily']
        weekly_fails = [f["cat"] for f in failures if f['time'] == 'weekly']
        intro_parts = []
        if daily_fails: intro_parts.append(f"{' and '.join(daily_fails)} on the date")
        if weekly_fails: intro_parts.append(f"{' and '.join(weekly_fails)} from week of testing")
        fail_intro = f"However, microbial growth was observed during { ' and '.join(intro_parts) }." if intro_parts else ""
        detail_sentences = []
        for i, f in enumerate(failures):
            is_plural = bool(re.search(r'\d+', f['obs']) and int(re.search(r'\d+', f['obs']).group()) > 1)
            verb_detect = "were" if is_plural else "was"; noun_id = "organisms were" if is_plural else "organism was"
            method_text = "differential staining" if "gram" in f['id'].lower() else "microbial identification"
            base_sentence = f"{f['obs']} {verb_detect} detected during {f['cat']} and was submitted for {method_text} under sample ID {f['etx']}, where the {noun_id} identified as {f['id']}"
            detail_sentences.append(f"Specifically, {base_sentence}." if i==0 else f"Additionally, {base_sentence}." if i==1 else f"Furthermore, {base_sentence}." if i==2 else f"Also, {base_sentence}.")
        det = f"{fail_intro} {' '.join(detail_sentences)}"
    return narr, det

def create_table_pdf(data):
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=9, leading=11, alignment=TA_CENTER)
    header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=9, leading=11, alignment=TA_CENTER, fontName='Helvetica-Bold')
    def p(text, is_header=False): return Paragraph(str(text), header_style if is_header else cell_style)
    elements = []
    elements.append(Paragraph(f"Appendix: Supplemental Tables for {data['sample_id']}", styles['Heading1']))
    elements.append(Spacer(1, 15))
    elements.append(Paragraph(f"Table 1: Information for {data['sample_id']} under investigation", styles['Heading2']))
    elements.append(Spacer(1, 5))
    t1_headers = [p("Processing Analyst", True), p("Reading Analyst", True), p("Sample ID", True), p("Events", True), p("Confirmed Microbial Events", True), p("Morphology Description", True)]
    t1_row = [p(data['analyst_name']), p(data['reader_name']), p(data['sample_id']), p(data['event_number']), p(data['confirm_number']), p(f"{data['organism_morphology']}-shaped Morphology")]
    t1 = Table([t1_headers, t1_row], colWidths=[110, 110, 130, 60, 120, 180])
    t1.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey), ('GRID', (0, 0), (-1, -1), 0.5, colors.black), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('TOPPADDING', (0, 0), (-1, -1), 5), ('BOTTOMPADDING', (0, 0), (-1, -1), 5)]))
    elements.append(t1); elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"Table 2: Environmental Monitoring from Testing Performed on {data['test_date']}", styles['Heading2']))
    elements.append(Spacer(1, 5))
    t2_headers = [p(h, True) for h in ["Sampling Site", "Freq", "Date", "Analyst", "Observation", "Plate ETX ID", "Microbial ID", "Notes"]]
    rows = []
    rows.append([p("Personnel EM Bracketing", True)] + [""]*7)
    rows.append([p("Personal (Left/Right)"), p("Daily"), p(data['test_date']), p(data['analyst_initial']), p(data['obs_pers_dur']), p(data['etx_pers_dur']), p(data['id_pers_dur']), p("None")])
    rows.append([p(f"BSC EM Bracketing ({data['bsc_id']})", True)] + [""]*7)
    rows.append([p("Surface Sampling (ISO 5)"), p("Daily"), p(data['test_date']), p(data['analyst_initial']), p(data['obs_surf_dur']), p(data['etx_surf_dur']), p(data['id_surf_dur']), p("None")])
    rows.append([p("Settling Sampling (ISO 5)"), p("Daily"), p(data['test_date']), p(data['analyst_initial']), p(data['obs_sett_dur']), p(data['etx_sett_dur']), p(data['id_sett_dur']), p("None")])
    rows.append([p(f"Weekly Bracketing (CR {data['cr_id']})", True)] + [""]*7)
    rows.append([p("Active Air Sampling"), p("Weekly"), p(data['date_of_weekly']), p(data['weekly_initial']), p(data['obs_air_wk_of']), p(data['etx_air_wk_of']), p(data['id_air_wk_of']), p("None")])
    rows.append([p("Surface Sampling"), p("Weekly"), p(data['date_of_weekly']), p(data['weekly_initial']), p(data['obs_room_wk_of']), p(data['etx_room_wk_of']), p(data['id_room_wk_of']), p("None")])
    t2 = Table([t2_headers] + rows, colWidths=[140, 50, 60, 45, 120, 100, 140, 55])
    t2.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.5, colors.black), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey), ('BACKGROUND', (0, 1), (-1, 1), colors.whitesmoke), ('SPAN', (0, 1), (-1, 1)), ('BACKGROUND', (0, 3), (-1, 3), colors.whitesmoke), ('SPAN', (0, 3), (-1, 3)), ('BACKGROUND', (0, 6), (-1, 6), colors.whitesmoke), ('SPAN', (0, 6), (-1, 6))]))
    elements.append(t2)
    doc.build(elements); buffer.seek(0)
    return buffer

def parse_email_text(text):
    # 1. JSON LOAD
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            for k, v in data.items():
                if k in FIELD_KEYS or k == "include_phase2": st.session_state[k] = v
            st.success("✅ Magic Import Successful! Reloading..."); time.sleep(1); st.rerun(); return
    except json.JSONDecodeError: pass
    # 2. REGEX PARSING
    if m := re.search(r"OOS-(\d+)", text): st.session_state.oos_id = m.group(1)
    if m := re.search(r"^(?:.*\n)?(.*\bE\d{5}\b.*)$", text, re.MULTILINE): st.session_state.client_name = m.group(1).strip()
    if m := re.search(r"(ETX-\d{6}-\d{4})", text): st.session_state.sample_id = m.group(1).strip()
    if m := re.search(r"Sample\s*Name:\s*(.*)", text, re.I): st.session_state.sample_name = m.group(1).strip()
    if m := re.search(r"(?:Lot|Batch)\s*[:\.]?\s*([^\n\r]+)", text, re.I): st.session_state.lot_number = m.group(1).strip()
    if m := re.search(r"testing\s*on\s*(\d{2}\s*\w{3}\s*\d{4})", text, re.I):
        try: st.session_state.test_date = datetime.strptime(m.group(1).strip(), "%d %b %Y").strftime("%d%b%y")
        except: pass
    if m := re.search(r"\(\s*([A-Z]{2,3})\s*\d+[a-z]{2}\s*Sample\)", text): 
        initial = m.group(1).strip()
        st.session_state.analyst_initial = initial
        found_name = get_full_name(initial)
        st.session_state.analyst_name = found_name
    if m := re.search(r"(\w+)-shaped", text, re.I):
        found_shape = m.group(1).lower()
        if "cocci" in found_shape: st.session_state.org_choice = "cocci"
        elif "rod" in found_shape: st.session_state.org_choice = "rod"
        else: st.session_state.org_choice = "Other"; st.session_state.manual_org = found_shape
