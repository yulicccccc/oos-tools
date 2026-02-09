import streamlit as st
import os
import re
import json
import io
import sys
import subprocess
import time
from datetime import datetime, timedelta

# --- SAFE UTILS IMPORT ---
try:
    from utils import apply_eagle_style, get_room_logic
except ImportError:
    def apply_eagle_style(): pass
    def get_room_logic(i): return "Unknown", "000", "", "Unknown"

# --- PAGE CONFIG ---
st.set_page_config(page_title="ScanRDI Investigation", layout="wide")
apply_eagle_style()

# --- CUSTOM STYLING ---
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #f0f2f6; }
    .main { background-color: #ffffff; }
    .stTextArea textarea { background-color: #ffffff; color: #31333F; border: 1px solid #d6d6d6; }
    div[data-testid="stNotification"] { border: 2px solid #ff4b4b; background-color: #ffe8e8; }
    div[data-testid="stAlert"] { border-radius: 5px; }
    </style>
    """, unsafe_allow_html=True)

# --- HELPER: LAZY INSTALLER ---
def ensure_dependencies():
    """Checks for required libraries and installs them if missing."""
    required = ["docxtpl", "pypdf", "reportlab"]
    missing = []
    for lib in required:
        try:
            __import__(lib)
        except ImportError:
            missing.append(lib)
    
    if missing:
        placeholder = st.empty()
        placeholder.warning(f"⚙️ Installing missing libraries: {', '.join(missing)}... (App will reload)")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)
            placeholder.success("Libraries installed! Reloading...")
            time.sleep(1)
            st.rerun()
        except Exception as e:
            placeholder.error(f"Installation failed: {e}")

# --- HELPER: INITIAL TO FULL NAME MAPPING ---
def get_full_name(initial):
    """Auto-converts initials to full names based on lab personnel."""
    if not initial: return ""
    mapping = {
        "KA": "Kathleen Aruta", "DH": "Domiasha Harrison",
        "GL": "Guanchen Li", "DS": "Devanshi Shah",
        "KC": "Kira C", "QC": "Qiyue Chen",
        "JD": "John Doe"
    }
    return mapping.get(initial.strip().upper(), initial)

# --- HELPER: VALIDATION ---
def validate_inputs():
    """Checks for empty fields and formatting errors."""
    errors = []
    warnings = []
    
    reqs = {
        "OOS Number": "oos_id", 
        "Client Name": "client_name", 
        "Sample ID": "sample_id", 
        "Test Date": "test_date",
        "Sample Name": "sample_name", 
        "Lot Number": "lot_number",
        "Prepper Name": "prepper_name",
        "Processor Name": "analyst_name",
        "Reader Name": "reader_name",
        "Changeover Name": "changeover_name",
        "BSC ID": "bsc_id",
        "ScanRDI ID": "scan_id",
        "Control Lot": "control_lot",
        "Control Exp": "control_exp",
        "Events Number": "event_number",
        "Confirmed #": "confirm_number"
    }
    
    for label, key in reqs.items():
        val = st.session_state.get(key, "").strip()
        if not val:
            warnings.append(label)
            
    date_val = st.session_state.get("test_date", "").strip()
    if date_val:
        try:
            datetime.strptime(date_val, "%d%b%y")
        except ValueError:
            errors.append(f"❌ Date Format Error: '{date_val}' is invalid. Please use format like '07Jan26' (DDMMMYY).")
            
    return errors, warnings

# --- HELPER: TABLE PDF GENERATOR (P1) ---
def create_table_pdf(data):
    """Generates the Tables PDF using ReportLab."""
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

    def p(text, is_header=False):
        return Paragraph(str(text), header_style if is_header else cell_style)

    elements = []
    elements.append(Paragraph(f"Appendix: Supplemental Tables for {data['sample_id']}", styles['Heading1']))
    elements.append(Spacer(1, 15))

    # TABLE 1
    elements.append(Paragraph(f"Table 1: Information for {data['sample_id']} under investigation", styles['Heading2']))
    elements.append(Spacer(1, 5))
    t1_headers = [p("Processing Analyst", True), p("Reading Analyst", True), p("Sample ID", True), p("Events", True), p("Confirmed Microbial Events", True), p("Morphology Description", True)]
    t1_row = [
        p(data['analyst_name']), p(data['reader_name']), p(data['sample_id']), 
        p(data['event_number']), p(data['confirm_number']), p(f"{data['organism_morphology']}-shaped Morphology")
    ]
    t1 = Table([t1_headers, t1_row], colWidths=[110, 110, 130, 60, 120, 180])
    t1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    elements.append(t1)
    elements.append(Spacer(1, 20))

    # TABLE 2
    elements.append(Paragraph(f"Table 2: Environmental Monitoring from Testing Performed on {data['test_date']}", styles['Heading2']))
    elements.append(Spacer(1, 5))
    
    h_labels = ["Sampling Site", "Freq", "Date", "Analyst", "Observation", "Plate ETX ID", "Microbial ID", "Notes"]
    t2_headers = [p(h, True) for h in h_labels]
    
    rows = []
    rows.append([p("Personnel EM Bracketing", True), "", "", "", "", "", "", ""])
    rows.append([p("Personal (Left/Right)"), p("Daily"), p(data['test_date']), p(data['analyst_initial']), p(data['obs_pers_dur']), p(data['etx_pers_dur']), p(data['id_pers_dur']), p("None")])
    
    rows.append([p(f"BSC EM Bracketing ({data['bsc_id']})", True), "", "", "", "", "", "", ""])
    rows.append([p("Surface Sampling (ISO 5)"), p("Daily"), p(data['test_date']), p(data['analyst_initial']), p(data['obs_surf_dur']), p(data['etx_surf_dur']), p(data['id_surf_dur']), p("None")])
    rows.append([p("Settling Sampling (ISO 5)"), p("Daily"), p(data['test_date']), p(data['analyst_initial']), p(data['obs_sett_dur']), p(data['etx_sett_dur']), p(data['id_sett_dur']), p("None")])
    
    rows.append([p(f"Weekly Bracketing (CR {data['cr_id']})", True), "", "", "", "", "", "", ""])
    rows.append([p("Active Air Sampling"), p("Weekly"), p(data['date_of_weekly']), p(data['weekly_initial']), p(data['obs_air_wk_of']), p(data['etx_air_wk_of']), p(data['id_air_wk_of']), p("None")])
    rows.append([p("Surface Sampling"), p("Weekly"), p(data['date_of_weekly']), p(data['weekly_initial']), p(data['obs_room_wk_of']), p(data['etx_room_wk_of']), p(data['id_room_wk_of']), p("None")])

    t2 = Table([t2_headers] + rows, colWidths=[140, 50, 60, 45, 120, 100, 140, 55])
    t2.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.whitesmoke),
        ('SPAN', (0, 1), (-1, 1)),
        ('BACKGROUND', (0, 3), (-1, 3), colors.whitesmoke),
        ('SPAN', (0, 3), (-1, 3)),
        ('BACKGROUND', (0, 6), (-1, 6), colors.whitesmoke),
        ('SPAN', (0, 6), (-1, 6)),
    ]))
    elements.append(t2)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# --- FILE PERSISTENCE ---
STATE_FILE = "investigation_state.json"
field_keys = [
    "oos_id", "client_name", "sample_id", "test_date", "sample_name", "lot_number", 
    "dosage_form", "monthly_cleaning_date", 
    "prepper_initial", "prepper_name", 
    "analyst_initial", "analyst_name",
    "changeover_initial", "changeover_name",
    "reader_initial", "reader_name",
    "writer_name", 
    "bsc_id", "chgbsc_id", "scan_id", 
    "shift_number", "active_platform",
    "org_choice", "manual_org", "test_record", "control_pos", "control_lot", 
    "control_exp", 
    "weekly_init", "date_weekly", 
    "equipment_summary", "narrative_summary", "em_details", 
    "sample_history_paragraph", "incidence_count", "oos_refs",
    "other_positives", "cross_contamination_summary",
    "total_pos_count_num", "current_pos_order",
    "diff_changeover_bsc", "has_prior_failures",
    "em_growth_observed", "diff_changeover_analyst",
    "diff_reader_analyst", "em_growth_count",
    "event_number", "confirm_number",
    "obs_pers", "etx_pers", "id_pers", 
    "obs_surf", "etx_surf", "id_surf", 
    "obs_sett", "etx_sett", "id_sett", 
    "obs_air", "etx_air_weekly", "id_air_weekly", 
    "obs_room", "etx_room_weekly", "id_room_wk_of",
    
    # --- PHASE 2 KEYS ---
    "include_phase2",
    "retest_date", "retest_sample_id", "retest_result", "retest_scan_id",
    "retest_prepper_name", "retest_prepper_initial",
    "retest_analyst_name", "retest_analyst_initial",
    "retest_reader_name", "retest_reader_initial",
    "retest_changeover_name", "retest_changeover_initial",
    "retest_bsc_id", "retest_chgbsc_id",
    "diff_retest_reader", "diff_retest_changeover", "diff_retest_bsc"
]
for i in range(20):
    field_keys.extend([f"other_id_{i}", f"other_order_{i}", f"prior_oos_{i}", f"em_cat_{i}", f"em_obs_{i}", f"em_etx_{i}", f"em_id_{i}"])

def load_saved_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r") as f:
                saved_data = json.load(f)
            for key, value in saved_data.items():
                if key in st.session_state: st.session_state[key] = value
        except: pass

def save_current_state():
    try:
        data = {k: v for k, v in st.session_state.items() if k in field_keys}
        with open(STATE_FILE, "w") as f: json.dump(data, f)
    except: pass

# --- HELPERS ---
def clean_filename(text): return re.sub(r'[\\/*?:"<>|]', '_', str(text)).strip() if text else ""

def ordinal(n):
    try:
        n = int(n)
        if 11 <= (n % 100) <= 13: suffix = 'th'
        else:
            r = n % 10
            if r == 1: suffix = 'st'
            elif r == 2: suffix = 'nd'
            elif r == 3: suffix = 'rd'
            else: suffix = 'th'
        return f"{n}{suffix}"
    except: return str(n)

def num_to_words(n):
    mapping = {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six", 7: "seven", 8: "eight", 9: "nine", 10: "ten"}
    return mapping.get(n, str(n))

# --- GENERATORS (P1 Specific) ---
def get_room_logic(bsc_id):
    try:
        from utils import get_room_logic as utils_get_room
        return utils_get_room(bsc_id)
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
    if st.session_state.incidence_count == 0 or st.session_state.has_prior_failures == "No":
        phrase = "no prior failures"
    else:
        pids = [st.session_state.get(f"prior_oos_{i}","").strip() for i in range(st.session_state.incidence_count) if st.session_state.get(f"prior_oos_{i}")]
        if not pids: refs_str = "..."
        elif len(pids) == 1: refs_str = pids[0]
        else: refs_str = ", ".join(pids[:-1]) + " and " + pids[-1]
        if len(pids) == 1: phrase = f"1 incident ({refs_str})"
        else: phrase = f"{len(pids)} incidents ({refs_str})"
    
    return f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte \"{st.session_state.sample_name}\" has had {phrase} using the Scan RDI method during this period."

def generate_cross_contam_text():
    if st.session_state.other_positives == "No": 
        return "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
    num = st.session_state.total_pos_count_num - 1
    other_list_ids = []
    detail_sentences = []
    for i in range(num):
        oid = st.session_state.get(f"other_id_{i}", "")
        oord_num = st.session_state.get(f"other_order_{i}", 1)
        oord_text = ordinal(oord_num)
        if oid:
            other_list_ids.append(oid)
            detail_sentences.append(f"{oid} was the {oord_text} sample processed")
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

# --- LOGIC: SYNC DYNAMIC UI -> FIXED FIELDS ---
def sync_dynamic_to_fixed():
    # 1. Reset defaults
    default_obs, default_etx, default_id = "No Growth", "N/A", "N/A"
    fixed_map = {
        "Personnel Obs": ("obs_pers", "etx_pers", "id_pers"),
        "Surface Obs": ("obs_surf", "etx_surf", "id_surf"),
        "Settling Obs": ("obs_sett", "etx_sett", "id_sett"),
        "Weekly Air Obs": ("obs_air", "etx_air_weekly", "id_air_weekly"),
        "Weekly Surf Obs": ("obs_room", "etx_room_weekly", "id_room_wk_of")
    }
    for cat, (k_obs, k_etx, k_id) in fixed_map.items():
        st.session_state[k_obs] = default_obs
        st.session_state[k_etx] = default_etx
        st.session_state[k_id] = default_id

    # 2. Update with user inputs
    if st.session_state.get("em_growth_observed") == "Yes":
        count = st.session_state.get("em_growth_count", 1)
        for i in range(count):
            cat = st.session_state.get(f"em_cat_{i}")
            obs = st.session_state.get(f"em_obs_{i}", "")
            etx = st.session_state.get(f"em_etx_{i}", "")
            mid = st.session_state.get(f"em_id_{i}", "")
            if cat in fixed_map:
                k_obs, k_etx, k_id = fixed_map[cat]
                st.session_state[k_obs] = obs
                st.session_state[k_etx] = etx
                st.session_state[k_id] = mid

def generate_narrative_and_details():
    # Force Sync before generating text
    sync_dynamic_to_fixed()
    
    failures = []
    def is_fail(val): return val.strip() and val.strip().lower() != "no growth"
    
    if is_fail(st.session_state.obs_pers):
        failures.append({"cat": "personnel sampling", "obs": st.session_state.obs_pers, "etx": st.session_state.etx_pers, "id": st.session_state.id_pers, "time": "daily"})
    if is_fail(st.session_state.obs_surf):
        failures.append({"cat": "surface sampling", "obs": st.session_state.obs_surf, "etx": st.session_state.etx_surf, "id": st.session_state.id_surf, "time": "daily"})
    if is_fail(st.session_state.obs_sett):
        failures.append({"cat": "settling plates", "obs": st.session_state.obs_sett, "etx": st.session_state.etx_sett, "id": st.session_state.id_sett, "time": "daily"})
    if is_fail(st.session_state.obs_air):
        failures.append({"cat": "weekly active air sampling", "obs": st.session_state.obs_air, "etx": st.session_state.etx_air_weekly, "id": st.session_state.id_air_weekly, "time": "weekly"})
    if is_fail(st.session_state.obs_room):
        failures.append({"cat": "weekly surface sampling", "obs": st.session_state.obs_room, "etx": st.session_state.etx_room_weekly, "id": st.session_state.id_room_wk_of, "time": "weekly"})

    pass_daily_clean = []
    if not is_fail(st.session_state.obs_pers): pass_daily_clean.append("personal sampling (left touch and right touch)")
    if not is_fail(st.session_state.obs_surf): pass_daily_clean.append("surface sampling")
    if not is_fail(st.session_state.obs_sett): pass_daily_clean.append("settling plates")
    
    pass_wk_clean = []
    if not is_fail(st.session_state.obs_air): pass_wk_clean.append("weekly active air sampling")
    if not is_fail(st.session_state.obs_room): pass_wk_clean.append("weekly surface sampling")

    narr_parts = []
    if pass_daily_clean:
        if len(pass_daily_clean) == 1: d_str = pass_daily_clean[0]
        elif len(pass_daily_clean) == 2: d_str = f"{pass_daily_clean[0]} and {pass_daily_clean[1]}"
        else: d_str = f"{pass_daily_clean[0]}, {pass_daily_clean[1]}, and {pass_daily_clean[2]}"
        narr_parts.append(f"no microbial growth was observed in {d_str}")
        
    if pass_wk_clean:
        if len(pass_wk_clean) == 1: w_str = pass_wk_clean[0]
        elif len(pass_wk_clean) == 2: w_str = f"{pass_wk_clean[0]} and {pass_wk_clean[1]}"
        else: w_str = ", ".join(pass_wk_clean)
        if narr_parts: narr_parts.append(f"Additionally, {w_str} showed no microbial growth")
        else: narr_parts.append(f"no microbial growth was observed in {w_str}")

    narr = "Upon analyzing the environmental monitoring results, " + ". ".join(narr_parts) + "." if narr_parts else "Upon analyzing the environmental monitoring results, microbial growth was observed in all sampled areas."

    det = ""
    if failures:
        daily_fails = [f["cat"] for f in failures if f['time'] == 'daily']
        weekly_fails = [f["cat"] for f in failures if f['time'] == 'weekly']
        intro_parts = []
        if daily_fails:
            if len(daily_fails) == 1: d_str = daily_fails[0]
            elif len(daily_fails) == 2: d_str = f"{daily_fails[0]} and {daily_fails[1]}"
            else: d_str = ", ".join(daily_fails[:-1]) + f", and {daily_fails[-1]}"
            intro_parts.append(f"{d_str} on the date")
        if weekly_fails:
            if len(weekly_fails) == 1: w_str = weekly_fails[0]
            elif len(weekly_fails) == 2: w_str = f"{weekly_fails[0]} and {weekly_fails[1]}"
            else: w_str = ", ".join(weekly_fails[:-1]) + f", and {weekly_fails[-1]}"
            intro_parts.append(f"{w_str} from week of testing")
        fail_intro = f"However, microbial growth was observed during { ' and '.join(intro_parts) }." if len(intro_parts) > 0 else ""
        
        detail_sentences = []
        for i, f in enumerate(failures):
            id_text = f['id']
            obs_val = f['obs']
            is_plural = False
            m = re.search(r'\d+', obs_val)
            if m and int(m.group()) > 1: is_plural = True
            
            verb_detect = "were" if is_plural else "was"
            noun_id = "organisms were" if is_plural else "organism was"
            method_text = "differential staining" if "gram" in id_text.lower() else "microbial identification"
            
            base_sentence = f"{obs_val} {verb_detect} detected during {f['cat']} and was submitted for {method_text} under sample ID {f['etx']}, where the {noun_id} identified as {id_text}"
            
            if i == 0: full_sent = f"Specifically, {base_sentence}."
            elif i == 1: full_sent = f"Additionally, {base_sentence}."
            elif i == 2: full_sent = f"Furthermore, {base_sentence}."
            else: full_sent = f"Also, {base_sentence}."
            detail_sentences.append(full_sent)
            
        det = f"{fail_intro} {' '.join(detail_sentences)}"
    return narr, det

# --- INIT STATE LOOP (Restored) ---
def init_state(key, default=""): 
    if key not in st.session_state: st.session_state[key] = default

# Loop through all keys to prevent AttributeError
for k in field_keys:
    if k in ["incidence_count","total_pos_count_num","current_pos_order","em_growth_count"]: init_state(k, 1)
    elif k == "include_phase2": init_state(k, False)
    elif "etx" in k or "id" in k: init_state(k, "N/A")
    elif k in ["event_number", "confirm_number"]: init_state(k, "1")
    else: init_state(k, "No" if "diff" in k or "has" in k or "growth" in k or "other" in k else "")

if "data_loaded" not in st.session_state:
    load_saved_state(); st.session_state.data_loaded = True
if "report_generated" not in st.session_state:
    st.session_state.report_generated = False
if "submission_warnings" not in st.session_state:
    st.session_state.submission_warnings = []
if "p2_generated" not in st.session_state:
    st.session_state.p2_generated = False

# --- PARSER (FIXED EMAIL LOGIC) ---
def parse_email_text(text):
    if m := re.search(r"OOS-(\d+)", text): st.session_state.oos_id = m.group(1)
    if m := re.search(r"^(?:.*\n)?(.*\bE\d{5}\b.*)$", text, re.MULTILINE): 
        st.session_state.client_name = m.group(1).strip()
    if m := re.search(r"(ETX-\d{6}-\d{4})", text): st.session_state.sample_id = m.group(1).strip()
    if m := re.search(r"Sample\s*Name:\s*(.*)", text, re.I): st.session_state.sample_name = m.group(1).strip()
    if m := re.search(r"(?:Lot|Batch)\s*[:\.]?\s*([^\n\r]+)", text, re.I): st.session_state.lot_number = m.group(1).strip()
    if m := re.search(r"testing\s*on\s*(\d{2}\s*\w{3}\s*\d{4})", text, re.I):
        try: st.session_state.test_date = datetime.strptime(m.group(1).strip(), "%d %b %Y").strftime("%d%b%y")
        except: pass
    
    # FIXED: Capture initial, THEN generate full name immediately
    if m := re.search(r"\(\s*([A-Z]{2,3})\s*\d+[a-z]{2}\s*Sample\)", text): 
        initial = m.group(1).strip()
        st.session_state.analyst_initial = initial
        # Force full name generation here
        st.session_state.analyst_name = get_full_name(initial)

    if m := re.search(r"(\w+)-shaped", text, re.I):
        found_shape = m.group(1).lower()
        if "cocci" in found_shape: st.session_state.org_choice = "cocci"
        elif "rod" in found_shape: st.session_state.org_choice = "rod"
        else: st.session_state.org_choice = "Other"; st.session_state.manual_org = found_shape
    save_current_state()

st.title("🦠 ScanRDI Investigation")

# --- UI ---
st.header("📧 Smart Email Import")
email_input = st.text_area("Paste OOS Notification email:", height=150)
if st.button("🪄 Parse Email"): parse_email_text(email_input); st.success("Updated!"); st.rerun()

st.header("1. General Test Details")
c1, c2, c3 = st.columns(3)
with c1: 
    st.text_input("OOS Number", key="oos_id", help="Required")
    st.text_input("Client Name", key="client_name", help="Required")
    st.text_input("Sample ID (ETX)", key="sample_id", help="Required")
with c2: 
    st.text_input("Test Date (07Jan26)", key="test_date", help="Required")
    st.text_input("Sample Name", key="sample_name", help="Required")
    st.text_input("Lot Number", key="lot_number", help="Required")
with c3: 
    st.selectbox("Dosage Form", ["Injectable","Aqueous Solution","Liquid","Solution"], key="dosage_form")
    st.text_input("Monthly Cleaning Date", key="monthly_cleaning_date", help="Required")

st.header("2. Personnel")
p1, p2 = st.columns(2)
with p1: 
    st.text_input("Prepper Initials", key="prepper_initial")
    if not st.session_state.prepper_name and st.session_state.prepper_initial: st.session_state.prepper_name = get_full_name(st.session_state.prepper_initial)
    st.text_input("Prepper Name", key="prepper_name", help="Required")
with p2: 
    st.text_input("Processor Initials", key="analyst_initial")
    if not st.session_state.analyst_name and st.session_state.analyst_initial: st.session_state.analyst_name = get_full_name(st.session_state.analyst_initial)
    st.text_input("Processor Name", key="analyst_name", help="Required")

st.radio("Different Reader?", ["No","Yes"], key="diff_reader_analyst", horizontal=True)
if st.session_state.diff_reader_analyst == "Yes":
    c1, c2 = st.columns(2)
    with c1: st.text_input("Reader Initials", key="reader_initial")
    with c2: st.text_input("Reader Name", key="reader_name")
else: st.session_state.reader_name = st.session_state.analyst_name; st.session_state.reader_initial = st.session_state.analyst_initial

st.radio("Different Changeover Analyst?", ["No","Yes"], key="diff_changeover_analyst", horizontal=True)
if st.session_state.diff_changeover_analyst == "Yes":
    c1, c2 = st.columns(2)
    with c1: st.text_input("Changeover Initials", key="changeover_initial")
    with c2: st.text_input("Changeover Name", key="changeover_name")
else: st.session_state.changeover_name = st.session_state.analyst_name; st.session_state.changeover_initial = st.session_state.analyst_initial

st.divider()
e1, e2 = st.columns(2)
bsc_list = ["1310","1309","1311","1312","1314","1313","1316","1798","Other"]
with e1: st.selectbox("Processing BSC ID", bsc_list, key="bsc_id")
with e2: 
    st.radio("Different Changeover BSC?", ["No","Yes"], key="diff_changeover_bsc", horizontal=True)
    if st.session_state.diff_changeover_bsc == "Yes": st.selectbox("Changeover BSC ID", bsc_list, key="chgbsc_id")
    else: st.session_state.chgbsc_id = st.session_state.bsc_id

st.header("3. Findings")
f1, f2 = st.columns(2)
with f1:
    st.selectbox("ScanRDI ID", ["1230","2017","1040","1877","2225","2132"], key="scan_id")
    st.selectbox("Shift Number", ["1", "2", "3"], key="shift_number")
    st.selectbox("Org Shape", ["rod","cocci","Other"], key="org_choice")
    if st.session_state.org_choice == "Other": st.text_input("Manual Shape", key="manual_org")
    c_a, c_b = st.columns(2)
    with c_a: st.text_input("Events Number", key="event_number", help="e.g. <1 event")
    with c_b: st.text_input("Confirmed #", key="confirm_number", help="e.g. 1")

with f2:
    st.selectbox("Positive Control", ["A. brasiliensis","B. subtilis","C. albicans","C. sporogenes","P. aeruginosa","S. aureus"], key="control_pos")
    st.text_input("Control Lot", key="control_lot", help="Required")
    st.text_input("Control Exp", key="control_exp", help="Required")

st.header("4. EM Observations")
st.radio("Microbial Growth Observed?", ["No","Yes"], key="em_growth_observed", horizontal=True)

if st.session_state.em_growth_observed == "Yes":
    count = st.number_input("Count of Growth Events", 1, 10, key="em_growth_count")
    for i in range(count):
        st.subheader(f"Growth Event #{i+1}")
        c1, c2, c3, c4 = st.columns(4)
        with c1: 
            st.selectbox(f"Category", ["Personnel Obs", "Surface Obs", "Settling Obs", "Weekly Air Obs", "Weekly Surf Obs"], key=f"em_cat_{i}")
        with c2: st.text_input(f"Obs (e.g. 1 CFU)", key=f"em_obs_{i}")
        with c3: st.text_input(f"ETX ID", key=f"em_etx_{i}")
        with c4: st.text_input(f"Microbial ID", key=f"em_id_{i}")

st.divider()
st.caption("Weekly Bracketing Meta")
m1, m2 = st.columns(2)
with m1: st.text_input("Weekly Monitor Initials", key="weekly_init", help="Required")
with m2: st.text_input("Date of Weekly Monitoring", key="date_weekly", help="Required")

st.header("5. Investigation Details")
st.subheader("Sample History")
st.radio("Prior failures in last 6 months?", ["No", "Yes"], key="has_prior_failures", horizontal=True)
if st.session_state.has_prior_failures == "Yes":
    count = st.number_input("Number of Prior Failures", 1, 10, key="incidence_count")
    for i in range(count):
        st.text_input(f"Prior Failure #{i+1} OOS ID", key=f"prior_oos_{i}")

st.subheader("Cross Contamination")
st.radio("Other samples tested positive same day?", ["No", "Yes"], key="other_positives", horizontal=True)
if st.session_state.other_positives == "Yes":
    st.number_input("Total Positive Samples that day", 2, 20, key="total_pos_count_num")
    st.number_input(f"Order of THIS Sample ({st.session_state.sample_id})", 1, 20, key="current_pos_order")
    num_others = st.session_state.total_pos_count_num - 1
    st.caption(f"Details for {num_others} other positive(s):")
    for i in range(num_others):
        c1, c2 = st.columns(2)
        with c1: st.text_input(f"Other Sample #{i+1} ID", key=f"other_id_{i}")
        with c2: st.number_input(f"Other Sample #{i+1} Order", 1, 20, key=f"other_order_{i}")

if st.button("🔄 Update Summaries Preview"):
    fresh_narr, fresh_det = generate_narrative_and_details()
    st.session_state.narrative_summary = fresh_narr
    st.session_state.em_details = fresh_det
    st.rerun()

st.text_area("Narrative Preview", key="narrative_summary", height=100)
if st.session_state.em_growth_observed == "Yes": st.text_area("Details Preview", key="em_details", height=150)

save_current_state()

st.divider()

# --- VALIDATION & GENERATION LOGIC ---
if st.button("🚀 GENERATE REPORT"):
    import time
    ensure_dependencies()
    errors, warnings = validate_inputs()
    
    if errors:
        for e in errors: st.error(e)
    elif warnings:
        st.session_state.submission_warnings = warnings
        st.rerun() 
    else:
        st.session_state.report_generated = True
        st.session_state.submission_warnings = [] 

# --- CONFIRMATION UI ---
if st.session_state.submission_warnings:
    st.warning(f"⚠️ The following fields are empty: {', '.join(st.session_state.submission_warnings)}")
    st.write("**Do you want to proceed?**")
    
    col_yes, col_no = st.columns([1, 5])
    
    if col_yes.button("✅ Yes, Proceed Anyway"):
        st.session_state.report_generated = True
        st.session_state.submission_warnings = []
        st.rerun()
        
    if col_no.button("❌ No, Let me Fix"):
        st.session_state.submission_warnings = []
        st.rerun()

# --- GENERATION & DOWNLOAD (P1) ---
if st.session_state.report_generated:
    
    # 1. GENERATE CONTENT
    fresh_narr, fresh_det = generate_narrative_and_details()
    fresh_equip = generate_equipment_text()
    fresh_history = generate_history_text()
    fresh_cross = generate_cross_contam_text()
    
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
    
    try: 
        d_obj = datetime.strptime(st.session_state.test_date, "%d%b%y")
        tr_id = f"{d_obj.strftime('%m%d%y')}-{st.session_state.scan_id}-{st.session_state.shift_number}"
        pdf_date_str = d_obj.strftime("%d-%b-%Y") 
    except: 
        tr_id = "N/A"; pdf_date_str = st.session_state.test_date

    suffix = "microorganism" if str(st.session_state.confirm_number).strip() == "1" else "microorganisms"
    raw_org = st.session_state.get('org_choice','') + " " + st.session_state.get('manual_org','')
    org_title = raw_org.strip().title()

    base_name = f"OOS-{st.session_state.oos_id} {st.session_state.client_name} - ScanRDI"
    safe_filename = clean_filename(base_name)

    final_data_docx = {k: v for k, v in st.session_state.items()}
    final_data_docx.update({
        "equipment_summary": fresh_equip,
        "sample_history_paragraph": fresh_history,
        "cross_contamination_summary": fresh_cross,
        "test_record": tr_id,
        "organism_morphology": org_title, 
        "control_positive": st.session_state.control_pos,
        "control_data": st.session_state.control_exp,
        "cr_id": t_room, "cr_suit": t_suite, "suit": t_suffix, "bsc_location": t_loc,
        "date_of_weekly": st.session_state.get("date_weekly", ""),
        "weekly_initial": st.session_state.get("weekly_init", ""),
        "obs_pers_dur": st.session_state.obs_pers, "etx_pers_dur": st.session_state.etx_pers, "id_pers_dur": st.session_state.id_pers,
        "obs_surf_dur": st.session_state.obs_surf, "etx_surf_dur": st.session_state.etx_surf, "id_surf_dur": st.session_state.id_surf,
        "obs_sett_dur": st.session_state.obs_sett, "etx_sett_dur": st.session_state.etx_sett, "id_sett_dur": st.session_state.id_sett,
        "obs_air_wk_of": st.session_state.obs_air, "etx_air_wk_of": st.session_state.etx_air_weekly, "id_air_wk_of": st.session_state.id_air_weekly,
        "obs_room_wk_of": st.session_state.obs_room, "etx_room_wk_of": st.session_state.etx_room_weekly, "id_room_wk_of": st.session_state.id_room_wk_of,
        "notes": "None" 
    })

    p7 = f"On {st.session_state.test_date}, a rapid sterility test was conducted on the sample using the ScanRDI method. The sample was initially prepared by Analyst {st.session_state.prepper_name}, processed by {st.session_state.analyst_name}, and subsequently read by {st.session_state.reader_name}. The test revealed {st.session_state.confirm_number} {org_title}-shaped viable {suffix}, see table 1."
    
    p8 = f"Table 1 (see attached tables) presents the environmental monitoring results for {st.session_state.sample_id}. The environmental monitoring (EM) plates were incubated for no less than 48 hours at 30-35°C and no less than an additional five days at 20-25°C as per SOP 2.600.002 (Environmental Monitoring of the Clean-room Facility)."
    p9 = fresh_narr
    if fresh_det: p9 += "\n\n" + fresh_det
    p10 = f"Monthly cleaning and disinfection, using H₂O₂, of the cleanroom (ISO 7) and its containing Biosafety Cabinets (BSCs, ISO 5) were performed on {st.session_state.monthly_cleaning_date}, as per SOP 2.600.018 Cleaning and Disinfection Procedure. It was documented that all H₂O₂ indicators passed."
    p11 = fresh_history
    p12 = f"To assess the potential for sample-to-sample contamination contributing to the positive results, a comprehensive review was conducted of all samples processed on the same day. {fresh_cross}"
    p13 = "Based on the observations outlined above, it is unlikely that the failing results were due to reagents, supplies, the cleanroom environment, the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is minimal and the original result is deemed to be valid."
    smart_phase1_part2 = "\n\n".join([p7, p8, p9, p10, p11, p12, p13])
    final_data_docx['Text Field50'] = smart_phase1_part2 

    # Save summary for P2 usage
    st.session_state.phase1_full_text = smart_phase1_part2

    # 2. GENERATE FILES
    docx_buf = None
    tables_docx_buf = None
    tables_pdf_buf = None
    pdf_form_buf = None

    # Main DOCX
    if os.path.exists("ScanRDI OOS template 0.docx"):
        try:
            from docxtpl import DocxTemplate
            doc = DocxTemplate("ScanRDI OOS template 0.docx")
            doc.render(final_data_docx)
            docx_buf = io.BytesIO(); doc.save(docx_buf); docx_buf.seek(0)
        except Exception as e: st.error(f"DOCX Error: {e}")

    # Tables DOCX
    if os.path.exists("tables for scan.docx"):
        try:
            from docxtpl import DocxTemplate
            doc_tbl = DocxTemplate("tables for scan.docx")
            doc_tbl.render(final_data_docx)
            tables_docx_buf = io.BytesIO(); doc_tbl.save(tables_docx_buf); tables_docx_buf.seek(0)
        except Exception as e: st.error(f"Tables DOCX Error: {e}")

    # Tables PDF
    try:
        tables_pdf_buf = create_table_pdf(final_data_docx)
    except Exception as e:
        st.warning(f"Tables PDF generation failed: {e}")

    # Main PDF Form
    try:
        from pypdf import PdfWriter, PdfReader
        analyst_sig_text = f"{st.session_state.analyst_name} (Written by: Qiyue Chen)"
        smart_personnel_block = (f"Prepper: \n{st.session_state.prepper_name} ({st.session_state.prepper_initial})\n\n"
                                 f"Processor:\n{st.session_state.analyst_name} ({st.session_state.analyst_initial})\n\n"
                                 f"Changeover\nProcessor:\n{st.session_state.changeover_name} ({st.session_state.changeover_initial})\n\n"
                                 f"Reader:\n{st.session_state.reader_name} ({st.session_state.reader_initial})")
        smart_incident_opening = f"On {st.session_state.test_date}, sample\n{st.session_state.sample_id} was found positive for viable microorganisms after ScanRDI\ntesting."
        
        # --- FIXED EMPTY LIST ERROR ---
        unique_analysts = []
        if st.session_state.prepper_name: unique_analysts.append(st.session_state.prepper_name)
        if st.session_state.analyst_name and st.session_state.analyst_name not in unique_analysts: unique_analysts.append(st.session_state.analyst_name)
        if st.session_state.reader_name and st.session_state.reader_name not in unique_analysts: unique_analysts.append(st.session_state.reader_name)
        
        if not unique_analysts: names_str = "N/A"
        elif len(unique_analysts) == 1: names_str = unique_analysts[0]
        elif len(unique_analysts) == 2: names_str = f"{unique_analysts[0]} and {unique_analysts[1]}"
        else: names_str = ", ".join(unique_analysts[:-1]) + " and " + unique_analysts[-1]

        smart_comment_interview = f"Yes, analysts {names_str} were interviewed comprehensively."
        smart_comment_samples = f"Yes, {st.session_state.sample_id}"
        smart_comment_records = f"Yes, See {tr_id} for more information."
        smart_comment_storage = f"Yes, Information is available in Eagle Trax Sample Location History under {st.session_state.sample_id}"
        
        p1 = f"All analysts involved in the prepping, processing, and reading of the samples – {names_str} – were interviewed and their answers are recorded throughout this document."
        p2 = f"The sample was stored upon arrival according to the Client’s instructions. Analysts {st.session_state.prepper_name} and {st.session_state.analyst_name} confirmed the integrity of the samples throughout both the preparation and processing stages. No leaks or turbidity were observed at any point, verifying the integrity of the sample."
        p3 = "All reagents and supplies mentioned in the material section above were stored according to the suppliers’ recommendations, and their integrity was visually verified before utilization. Moreover, each reagent and supply had valid expiration dates."
        p4 = f"During the preparation phase, {st.session_state.prepper_name} disinfected the samples using acidified bleach and placed them into a pre-disinfected storage bin. On {st.session_state.test_date}, prior to sample processing, {st.session_state.analyst_name} performed a second disinfection with acidified bleach, allowing a minimum contact time of 10 minutes before transferring the samples into the cleanroom suites. A final disinfection step was completed immediately before the samples were introduced into the ISO 5 Biological Safety Cabinet (BSC), E00{st.session_state.bsc_id}, located within the {t_loc}, (Suite {t_suite}{t_suffix}), All activities were performed in accordance with SOP 2.600.023, Rapid Scan RDI® Test Using FIFU Method."
        p5 = fresh_equip
        p6 = f"The analyst, {st.session_state.reader_name}, confirmed that the equipment was set up as per SOP 2.700.004 (Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup and Maintenance), and the negative control and the positive control for the analyst, {st.session_state.reader_name}, yielded expected results."
        smart_phase1_part1 = "\n\n".join([p1, p2, p3, p4, p5, p6])
        
        pdf_map = {
            'Text Field57': st.session_state.oos_id, 
            'Date Field0': pdf_date_str, 'Date Field1': pdf_date_str, 'Date Field2': pdf_date_str, 'Date Field3': pdf_date_str,
            'Text Field2': f"{st.session_state.sample_id}\n\n{st.session_state.client_name}", 
            'Text Field6': st.session_state.lot_number,
            'Text Field4': st.session_state.sample_name, 
            'Text Field5': st.session_state.dosage_form,
            'Text Field0': analyst_sig_text,
            'Text Field3': smart_personnel_block,
            'Text Field7': smart_incident_opening,
            'Text Field13': smart_comment_interview,
            'Text Field14': smart_comment_samples,
            'Text Field17': smart_comment_records,
            'Text Field21': smart_comment_storage,
            'Text Field30': f"E00{st.session_state.scan_id}",
            'Text Field32': f"E00{t_room} (CR{t_suite})",
            'Text Field34': f"E00{st.session_state.scan_id}",
            'Text Field24': st.session_state.control_pos,
            'Text Field25': st.session_state.control_lot,
            'Text Field26': st.session_state.control_exp,
            'Text Field49': smart_phase1_part1, 
            'Text Field50': smart_phase1_part2
        }

        if os.path.exists("ScanRDI OOS template.pdf"):
            writer = PdfWriter(clone_from="ScanRDI OOS template.pdf")
            for p in writer.pages: writer.update_page_form_field_values(p, pdf_map)
            pdf_form_buf = io.BytesIO(); writer.write(pdf_form_buf); pdf_form_buf.seek(0)
    
    except ImportError:
        pass
    except Exception as e:
        st.error(f"PDF Form Error: {e}")

    # --- 3. DISPLAY DOWNLOAD BUTTONS ---
    st.success("✅ Reports Generated Successfully!")
    st.markdown("### 📂 Download Reports")
    c1, c2 = st.columns(2)
    
    with c1:
        st.subheader("Word Documents")
        if docx_buf:
            st.download_button("📄 OOS Report (doc)", docx_buf, f"{safe_filename}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        if tables_docx_buf:
            st.download_button("📄 Tables (doc)", tables_docx_buf, f"Tables {safe_filename}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            
    with c2:
        st.subheader("PDF Documents")
        if pdf_form_buf:
            st.download_button("🔴 OOS Report (pdf)", pdf_form_buf, f"{safe_filename}.pdf", "application/pdf")
        if tables_pdf_buf:
            st.download_button("🔴 Tables (pdf)", tables_pdf_buf, f"Tables {safe_filename}.pdf", "application/pdf")

# =================================================================================================
# === PHASE 2 EXTENSION STARTS HERE ===
# =================================================================================================

st.markdown("---")
st.subheader("🚦 Phase 2 Investigation (Retest)")
st.checkbox("Include Phase 2 Investigation?", key="include_phase2")

# --- P2 GENERATOR FUNCTION ---
def generate_p2_docs():
    from docxtpl import DocxTemplate
    from pypdf import PdfWriter
    
    # 1. GENERATE RETEST EQUIPMENT TEXT (P2 SPECIFIC)
    def generate_retest_equipment_text(bsc_main, bsc_chg, analyst_main, analyst_chg, date_val):
        t_room, t_suite, t_suffix, t_loc = get_room_logic(bsc_main)
        c_room, c_suite, c_suffix, c_loc = get_room_logic(bsc_chg)
        
        subj = "Retest sample"
        
        # Scenario 1: Same BSC
        if bsc_main == bsc_chg:
            part1 = f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
            part2 = f"The ISO 5 BSC E00{bsc_main}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both testing and changeover steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{bsc_main} was certified and approved by both the Engineering and Quality Assurance teams. {subj} processing and changeover were conducted in the ISO 5 BSC E00{bsc_main} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {analyst_main} on {date_val}."
            return f"{part1}\n\n{part2}"
        else:
            if t_suite == c_suite:
                 part1 = f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
            else:
                 p1a = f"The cleanroom used for testing (E00{t_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which opens into the middle ISO 7 buffer room ({t_suite}A), and then into the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
                 p1b = f"The cleanroom used for changeover (E00{c_room}) consists of three interconnected sections: the innermost ISO 7 cleanroom ({c_suite}B), which opens into the middle ISO 7 buffer room ({c_suite}A), and then into the outermost ISO 8 anteroom ({c_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {c_suite}B through {c_suite}A and into {c_suite}."
                 part1 = f"{p1a}\n\n{p1b}"

            intro = f"The ISO 5 BSC E00{bsc_main}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), and ISO 5 BSC E00{bsc_chg}, located in the {c_loc}, (Suite {c_suite}{c_suffix}), were thoroughly cleaned and disinfected prior to their respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Furthermore, the BSCs used throughout testing, E00{bsc_main} for sample processing and E00{bsc_chg} for the changeover step, were certified and approved by both the Engineering and Quality Assurance teams."
            
            if analyst_main == analyst_chg:
                usage_sent = f"{subj} processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{bsc_main}) and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{bsc_chg}) by {analyst_main} on {date_val}."
            else:
                usage_sent = f"{subj} processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{bsc_main}) by {analyst_main} and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{bsc_chg}) by {analyst_chg} on {date_val}."

            return f"{part1}\n\n{intro} {usage_sent}"

    # 2. AUTO-RESOLVE NAMES (Trust user input primarily, fallback to initial map)
    p_name = st.session_state.retest_prepper_name or get_full_name(st.session_state.retest_prepper_initial)
    a_name = st.session_state.retest_analyst_name or get_full_name(st.session_state.retest_analyst_initial)
    r_name = st.session_state.retest_reader_name or get_full_name(st.session_state.retest_reader_initial)
    c_name = st.session_state.retest_changeover_name or get_full_name(st.session_state.retest_changeover_initial)

    # 3. PREPARE DATA
    retest_equip_sum = generate_retest_equipment_text(
        st.session_state.retest_bsc_id,
        st.session_state.retest_chgbsc_id,
        a_name,
        c_name,
        st.session_state.retest_date
    )
    
    r_room, r_suite, r_suffix, r_loc = get_room_logic(st.session_state.retest_bsc_id)
    rc_room, rc_suite, rc_suffix, rc_loc = get_room_logic(st.session_state.retest_chgbsc_id)

    # Smart Blocks
    smart_pers = (
        f"Prepper: \n{p_name} ({st.session_state.retest_prepper_initial})\n\n"
        f"Processors: \n{a_name} ({st.session_state.retest_analyst_initial})\n\n"
        f"Reader: \n{r_name} ({st.session_state.retest_reader_initial})"
    )
    
    smart_ids = f"Original Test:\n{st.session_state.sample_id}\n\nRetest:\n{st.session_state.retest_sample_id}"
    smart_orig_res = f"{st.session_state.sample_id} - Fail"
    smart_retest_res = f"{st.session_state.retest_sample_id} - {st.session_state.retest_result}"
    smart_bsc_list = f"{st.session_state.retest_bsc_id} and {st.session_state.retest_chgbsc_id}"
    smart_suite_list = f"Suite {r_suite}{r_suffix}, Suite {rc_suite}{rc_suffix}"
    smart_p1_block = f"INITIAL TEST UNDER {st.session_state.sample_id}\n\n{st.session_state.get('phase1_full_text', 'See Phase 1 Report')}"
    
    smart_p2_narrative = (
        f"RETEST UNDER SUBMISSION {st.session_state.retest_sample_id}\n\n"
        f"Analogous to original testing, the analysts involved in prepping, processing and reading the retest samples under "
        f"{st.session_state.retest_sample_id}, {p_name}, {a_name} and {r_name} "
        f"confirmed no deviations from standard procedures.\n\n"
        f"The retest sample was stored upon arrival according to the Client’s instructions. Analysts {p_name} and {a_name} "
        f"confirmed the integrity of the samples throughout both the preparation and processing stages. No leaks or turbidity were observed at any point, verifying that the samples remained intact.\n\n"
        f"All reagents and supplies mentioned in the material section above were stored according to the suppliers’ recommendations, and their integrity was visually verified before utilization. "
        f"Moreover, each reagent and supply had valid expiration dates.\n\n"
        f"During the preparation phase, {p_name} disinfected the samples using acidified bleach and placed them into a pre-disinfected storage bin. "
        f"On {st.session_state.retest_date}, prior to sample processing, {a_name} performed a second disinfection with acidified bleach, "
        f"allowing a minimum contact time of 10 minutes before transferring the samples into the cleanroom suites.\n\n"
        f"A final disinfection step was completed immediately before the samples were introduced into the ISO 5 Biological Safety Cabinet (BSC), E00{st.session_state.retest_bsc_id}, "
        f"located within the {r_loc}, (Suite {r_suite}{r_suffix}), All activities were performed in accordance with SOP 2.600.023, Rapid Scan RDI® Test Using FIFU Method.\n\n"
        f"{retest_equip_sum}\n\n"
        f"The analyst, {r_name}, confirmed that the Scan RDI equipment E00{st.session_state.retest_scan_id} was set up as per SOP 2.700.004 "
        f"(Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup and Maintenance), and the negative control and the positive control for the analyst, "
        f"{r_name}, yielded expected results.\n\n"
        f"On {st.session_state.retest_date}, a rapid sterility test was conducted on the retest sample using the ScanRDI method. The retest sample was initially prepared by Analyst "
        f"{p_name}, processed by {a_name} and subsequently read by {r_name}. "
        f"The retest sample under {st.session_state.retest_sample_id} {st.session_state.retest_result} the sterility test by ScanRDI method.\n\n"
        f"All reagents and supplies utilized during the testing process were within the expiration dates. Daily verifications (Control Beads), negative and positive controls, "
        f"were conducted to confirm the reliability of the testing process. All verification tests met the set forth criteria per SOP 2.600.023 (Rapid Scan RDI Test using FIFU Method), "
        f"and SOP 2.700.004 (Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup and Maintenance).\n\n"
        f"Following a detailed review of the available data, the conflicting results between the original test ({st.session_state.sample_id}) and the retest ({st.session_state.retest_sample_id}) "
        f"may be attributed to the non-uniform distribution of microorganisms within the sample, particularly if present at low concentrations.\n\n"
        f"Based on the observations outlined above, laboratory error cannot be conclusively confirmed for either the original test or the retest. Therefore, both the failing result "
        f"for {st.session_state.sample_id} and the passing result for {st.session_state.retest_sample_id} are considered valid.\n\n"
        f"The final disposition of the lot remains at the discretion of the client."
    )

    data = {k: v for k, v in st.session_state.items()}
    data.update({
        "retest_prepper_name": p_name, "retest_analyst_name": a_name, "retest_reader_name": r_name,
        "retest_equipment_summary": retest_equip_sum,
        "retest_bsc_location": r_loc, "retest_cr_suit": r_suite, "retest_suit": r_suffix,
        "retest_chgcr_suit": rc_suite, "retest_chgsuit": rc_suffix,
        "retest_chgbsc_id": st.session_state.retest_chgbsc_id,
        "retest_scan_id": st.session_state.retest_scan_id,
        "smart_retest_personnel_block": smart_pers,
        "smart_sample_id_block": smart_ids,
        "smart_original_result_str": smart_orig_res,
        "smart_retest_result_str": smart_retest_res,
        "smart_retest_bsc_list": smart_bsc_list,
        "smart_retest_suite_list": smart_suite_list,
        "smart_retest_scan_id": f"E00{st.session_state.retest_scan_id}",
        "smart_phase1_summary_block": smart_p1_block,
        "smart_phase2_narrative_block": smart_p2_narrative
    })

    # GENERATE FILES
    p2_docx_buf = None
    p2_pdf_buf = None

    if os.path.exists("ScanRDI OOS P2 template 0.docx"):
        try:
            doc0 = DocxTemplate("ScanRDI OOS P2 template 0.docx")
            doc0.render(data)
            p2_docx_buf = io.BytesIO(); doc0.save(p2_docx_buf); p2_docx_buf.seek(0)
        except Exception as e: st.error(f"P2 Main DOCX Error: {e}")

    if os.path.exists("ScanRDI OOS P2 template.pdf"):
        try:
            pdf_map = {
                "Text Field0": data["sample_name"],
                "Text Field1": smart_pers,
                "Text Field2": smart_ids,
                "Text Field3": smart_retest_res,
                "Text Field4": smart_orig_res,
                "Text Field30": data["oos_id"],
                "Date Field0": data["retest_date"],
                "Text Field8": data["smart_retest_scan_id"],
                "Text Field9": smart_bsc_list,
                "Text Field10": smart_suite_list,
                "Text Field22": smart_p1_block,
                "Text Field23": smart_p2_narrative
            }
            writer = PdfWriter(clone_from="ScanRDI OOS P2 template.pdf")
            for p in writer.pages: writer.update_page_form_field_values(p, pdf_map)
            p2_pdf_buf = io.BytesIO(); writer.write(p2_pdf_buf); p2_pdf_buf.seek(0)
        except Exception as e: st.error(f"P2 PDF Error: {e}")

    return p2_docx_buf, p2_pdf_buf

# --- P2 UI LOGIC ---
if st.session_state.include_phase2:
    st.info("💡 Tip: Enter Initials (e.g. DS) and press Enter. The system will auto-fill the full name if known.")
    
    r1, r2, r3 = st.columns(3)
    with r1: st.text_input("Retest Date (DDMMMYY)", key="retest_date"); st.text_input("Retest Sample ID", key="retest_sample_id")
    with r2: st.text_input("Retest Result", value="Pass", key="retest_result"); st.text_input("Retest Scan ID (e.g. 2017)", key="retest_scan_id")
    
    st.markdown("##### Retest Personnel")
    
    # Auto-fill Logic via Rerun on Initial Change
    p2_1, p2_2 = st.columns(2)
    with p2_1: 
        st.text_input("Retest Prepper", key="retest_prepper_name")
        if st.text_input("Initials", key="retest_prepper_initial"):
            if not st.session_state.retest_prepper_name:
                st.session_state.retest_prepper_name = get_full_name(st.session_state.retest_prepper_initial)
                st.rerun()

    with p2_2: 
        st.text_input("Retest Processor", key="retest_analyst_name")
        if st.text_input("Initials", key="retest_analyst_initial"):
            if not st.session_state.retest_analyst_name:
                st.session_state.retest_analyst_name = get_full_name(st.session_state.retest_analyst_initial)
                st.rerun()
    
    st.radio("Different Retest Reader?", ["No","Yes"], key="diff_retest_reader", horizontal=True)
    if st.session_state.diff_retest_reader == "Yes": 
        c1, c2 = st.columns(2)
        with c1: st.text_input("Retest Reader", key="retest_reader_name")
        with c2: 
            if st.text_input("Reader Initials", key="retest_reader_initial"):
                if not st.session_state.retest_reader_name:
                    st.session_state.retest_reader_name = get_full_name(st.session_state.retest_reader_initial)
                    st.rerun()
    else: 
        st.session_state.retest_reader_name = st.session_state.retest_analyst_name; 
        st.session_state.retest_reader_initial = st.session_state.retest_analyst_initial
    
    st.radio("Different Retest Changeover?", ["No","Yes"], key="diff_retest_changeover", horizontal=True)
    if st.session_state.diff_retest_changeover == "Yes": 
        st.text_input("Retest Changeover Name", key="retest_changeover_name")
        st.text_input("Chg Initials", key="retest_changeover_initial")
    else: 
        st.session_state.retest_changeover_name = st.session_state.retest_analyst_name; 
        st.session_state.retest_changeover_initial = st.session_state.retest_analyst_initial

    st.markdown("##### Retest Equipment")
    e2_1, e2_2 = st.columns(2)
    with e2_1: st.selectbox("Retest Proc. BSC", bsc_list, key="retest_bsc_id")
    with e2_2: 
        st.radio("Diff Retest Chg BSC?", ["No","Yes"], key="diff_retest_bsc", horizontal=True)
        if st.session_state.diff_retest_bsc == "Yes": st.selectbox("Retest Chg BSC", bsc_list, key="retest_chgbsc_id")
        else: st.session_state.retest_chgbsc_id = st.session_state.retest_bsc_id

    if st.button("🚀 GENERATE PHASE 2 REPORTS"):
        st.session_state.p2_generated = True

    if st.session_state.p2_generated:
        p2_doc, p2_pdf = generate_p2_docs()
        st.success("Phase 2 Reports Ready!")
        c1, c2 = st.columns(2)
        with c1:
            if p2_doc: st.download_button("📄 P2 Main Report", p2_doc, "P2_Report.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with c2:
            if p2_pdf: st.download_button("✅ P2 Final PDF", p2_pdf, "P2_Form.pdf", "application/pdf")
