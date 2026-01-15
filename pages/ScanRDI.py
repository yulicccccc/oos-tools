import streamlit as st
import os
import re
import json
import io
import sys
import subprocess
from datetime import datetime, timedelta

# --- UTILS IMPORT (Safe Fallback) ---
try:
    from utils import apply_eagle_style, get_full_name, get_room_logic
except ImportError:
    def apply_eagle_style(): pass
    def get_full_name(i): return i
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
    </style>
    """, unsafe_allow_html=True)

# --- HELPER: LAZY INSTALLER ---
def ensure_dependencies():
    """Installs missing libraries on the fly to prevent crashes."""
    required = ["docxtpl", "pypdf", "reportlab"]
    missing = []
    for lib in required:
        try:
            __import__(lib)
        except ImportError:
            missing.append(lib)
    
    if missing:
        placeholder = st.empty()
        placeholder.warning(f"⚙️ Installing missing libraries: {', '.join(missing)}... (This happens once)")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)
            placeholder.success("Libraries installed! Proceeding...")
            time.sleep(1)
            placeholder.empty()
        except Exception as e:
            placeholder.error(f"Installation failed: {e}")
            st.stop()

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
    "diff_reader_analyst", 
    # Fixed EM Keys
    "obs_pers", "etx_pers", "id_pers", 
    "obs_surf", "etx_surf", "id_surf", 
    "obs_sett", "etx_sett", "id_sett", 
    "obs_air", "etx_air_weekly", "id_air_weekly", 
    "obs_room", "etx_room_weekly", "id_room_wk_of"
]
# Dynamic keys
for i in range(20):
    field_keys.extend([f"other_id_{i}", f"other_order_{i}", f"prior_oos_{i}"])

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

# --- GENERATORS ---
def get_room_logic(bsc_id):
    try:
        from utils import get_room_logic as utils_get_room
        return utils_get_room(bsc_id)
    except:
        return "Unknown Room", "000", "", "Unknown Loc"

def generate_equipment_text():
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
    
    # CASE A: Same BSC
    if st.session_state.bsc_id == st.session_state.chgbsc_id:
        part1 = f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        part2 = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both testing and changeover steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{st.session_state.bsc_id} was certified and approved by both the Engineering and Quality Assurance teams. Sample processing and changeover were conducted in the ISO 5 BSC E00{st.session_state.bsc_id} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {st.session_state.analyst_name} on {st.session_state.test_date}."
        return f"{part1}\n\n{part2}"
    
    # CASE B/C: Different BSCs
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
        
    return f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte “{st.session_state.sample_name}” has had {phrase} using the Scan RDI method during this period."

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

def generate_narrative_and_details():
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

    # Pass Narrative Split
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

    if not narr_parts: narr = "Upon analyzing the environmental monitoring results, microbial growth was observed in all sampled areas."
    else: narr = "Upon analyzing the environmental monitoring results, " + ". ".join(narr_parts) + "."

    # Fail Narrative
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
            
        if len(intro_parts) == 2: fail_intro = f"However, microbial growth was observed during both {intro_parts[0]} and {intro_parts[1]}."
        else: fail_intro = f"However, microbial growth was observed during {intro_parts[0]}."
        
        detail_sentences = []
        for i, f in enumerate(failures):
            id_text = f['id']
            if "gram" in id_text.lower(): method_text = "differential staining"
            else: method_text = "microbial identification"
                
            base_sentence = f"{f['obs']} was detected during {f['cat']} and was submitted for {method_text} under sample ID {f['etx']}, where the organism was identified as {id_text}"
            
            if i == 0: full_sent = f"Specifically, {base_sentence}."
            elif i == 1: full_sent = f"Additionally, {base_sentence}."
            elif i == 2: full_sent = f"Furthermore, {base_sentence}."
            else: full_sent = f"Also, {base_sentence}."
            detail_sentences.append(full_sent)
            
        det = f"{fail_intro} {' '.join(detail_sentences)}"

    return narr, det

# --- INIT STATE ---
def init_state(key, default=""): 
    if key not in st.session_state: st.session_state[key] = default
for k in field_keys:
    if k in ["incidence_count","total_pos_count_num","current_pos_order","em_growth_count"]: init_state(k, 1)
    elif "etx" in k or "id" in k: init_state(k, "N/A")
    else: init_state(k, "No" if "diff" in k or "has" in k or "growth" in k or "other" in k else "")

if "data_loaded" not in st.session_state:
    load_saved_state(); st.session_state.data_loaded = True

# --- PARSER ---
def parse_email_text(text):
    if m := re.search(r"OOS-(\d+)", text): st.session_state.oos_id = m.group(1)
    if m := re.search(r"^.*\(E\d{5}\).*$", text, re.MULTILINE): st.session_state.client_name = m.group(0).strip()
    if m := re.search(r"(ETX-\d{6}-\d{4})", text): st.session_state.sample_id = m.group(1).strip()
    if m := re.search(r"Sample\s*Name:\s*(.*)", text, re.I): st.session_state.sample_name = m.group(1).strip()
    if m := re.search(r"(?:Lot|Batch)\s*[:\.]?\s*([^\n\r]+)", text, re.I): st.session_state.lot_number = m.group(1).strip()
    if m := re.search(r"testing\s*on\s*(\d{2}\s*\w{3}\s*\d{4})", text, re.I):
        try: st.session_state.test_date = datetime.strptime(m.group(1).strip(), "%d %b %Y").strftime("%d%b%y")
        except: pass
    
    if m := re.search(r"\(\s*([A-Z]{2,3})\s*\d+[a-z]{2}\s*Sample\)", text): 
        initial = m.group(1).strip()
        st.session_state.analyst_initial = initial
        if initial == "DS": st.session_state.analyst_name = "Devanshi Shah"
        else: st.session_state.analyst_name = get_full_name(initial)

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
with f2:
    st.selectbox("Positive Control", ["A. brasiliensis","B. subtilis","C. albicans","C. sporogenes","P. aeruginosa","S. aureus"], key="control_pos")
    st.text_input("Control Lot", key="control_lot", help="Required")
    st.text_input("Control Exp", key="control_exp", help="Required")

st.header("4. EM Observations")
st.radio("Microbial Growth Observed?", ["No","Yes"], key="em_growth_observed", horizontal=True)

if st.session_state.em_growth_observed == "No":
    st.session_state.obs_pers = "No Growth"
    st.session_state.obs_surf = "No Growth"
    st.session_state.obs_sett = "No Growth"
    st.session_state.obs_air = "No Growth"
    st.session_state.obs_room = "No Growth"
else:
    c1, c2, c3 = st.columns(3)
    with c1: 
        st.text_input("Personnel Obs", key="obs_pers", placeholder="e.g. 1 CFU or No Growth")
        st.text_input("Pers ETX", key="etx_pers")
        st.text_input("Pers ID", key="id_pers")
    with c2: 
        st.text_input("Surface Obs", key="obs_surf", placeholder="e.g. 1 CFU or No Growth")
        st.text_input("Surf ETX", key="etx_surf")
        st.text_input("Surf ID", key="id_surf")
    with c3: 
        st.text_input("Settling Obs", key="obs_sett", placeholder="e.g. 1 CFU or No Growth")
        st.text_input("Sett ETX", key="etx_sett")
        st.text_input("Sett ID", key="id_sett")
        
    st.divider()
    st.caption("Weekly Bracketing")
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("Weekly Air Obs", key="obs_air", placeholder="e.g. 1 CFU or No Growth")
        st.text_input("Air ETX", key="etx_air_weekly")
        st.text_input("Air ID", key="id_air_weekly")
    with c2:
        st.text_input("Weekly Surf Obs", key="obs_room", placeholder="e.g. 1 CFU or No Growth")
        st.text_input("Surf ETX", key="etx_room_weekly")
        st.text_input("Surf ID", key="id_room_wk_of")

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

if st.button("🚀 GENERATE FINAL REPORT"):
    # 1. Validation
    missing = []
    reqs = {"OOS #":"oos_id", "Client":"client_name", "Sample ID":"sample_id", "Date":"test_date", "Sample Name":"sample_name", "Lot":"lot_number", "Analyst":"analyst_name", "BSC":"bsc_id", "Scan ID":"scan_id"}
    for l,k in reqs.items():
        if not st.session_state.get(k,"").strip(): missing.append(l)
    if missing: st.error(f"Missing: {', '.join(missing)}"); st.stop()

    # 2. Dependency Check (LAZY LOAD)
    import time
    ensure_dependencies()

    # 3. Now Safe to Import
    from docxtpl import DocxTemplate
    from pypdf import PdfReader, PdfWriter, PdfMerger
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.styles import getSampleStyleSheet

    # 4. Update Generators
    fresh_narr, fresh_det = generate_narrative_and_details()
    fresh_equip = generate_equipment_text()
    fresh_history = generate_history_text()
    fresh_cross = generate_cross_contam_text()
    
    # 5. Data Prep & Date Conversion (PDF Specific)
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
    
    try: 
        d_obj = datetime.strptime(st.session_state.test_date, "%d%b%y")
        tr_id = f"{d_obj.strftime('%m%d%y')}-{st.session_state.scan_id}-{st.session_state.shift_number}"
        pdf_date_str = d_obj.strftime("%d-%b-%Y") 
    except: 
        tr_id = "N/A"
        pdf_date_str = st.session_state.test_date

    # Sanitize Table Data (No Growth = N/A)
    em_map = [("obs_pers", "etx_pers", "id_pers"), ("obs_surf", "etx_surf", "id_surf"), ("obs_sett", "etx_sett", "id_sett"), ("obs_air", "etx_air_weekly", "id_air_weekly"), ("obs_room", "etx_room_weekly", "id_room_wk_of")]
    for obs_k, etx_k, id_k in em_map:
        val = st.session_state[obs_k].strip()
        if not val or val.lower() == "no growth":
            st.session_state[obs_k] = "No Growth"
            st.session_state[etx_k] = "N/A"
            st.session_state[id_k] = "N/A"

    # --- WORD ---
    final_data_docx = {k: v for k, v in st.session_state.items()}
    final_data_docx.update({
        "equipment_summary": fresh_equip,
        "sample_history_paragraph": fresh_history,
        "cross_contamination_summary": fresh_cross,
        "test_record": tr_id,
        "organism_morphology": st.session_state.get('org_choice','') + " " + st.session_state.get('manual_org',''),
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

    # --- PDF ---
    analyst_sig_text = f"{st.session_state.analyst_name} (Written by: Qiyue Chen)"
    
    smart_personnel_block = (
        f"Prepper: \n{st.session_state.prepper_name} ({st.session_state.prepper_initial})\n\n"
        f"Processor:\n{st.session_state.analyst_name} ({st.session_state.analyst_initial})\n\n"
        f"Changeover\nProcessor:\n{st.session_state.changeover_name} ({st.session_state.changeover_initial})\n\n"
        f"Reader:\n{st.session_state.reader_name} ({st.session_state.reader_initial})"
    )

    smart_incident_opening = f"On {st.session_state.test_date}, sample\n{st.session_state.sample_id} was found positive for viable microorganisms after ScanRDI\ntesting."
    
    unique_analysts = []
    if st.session_state.prepper_name: unique_analysts.append(st.session_state.prepper_name)
    if st.session_state.analyst_name and st.session_state.analyst_name not in unique_analysts: unique_analysts.append(st.session_state.analyst_name)
    if st.session_state.reader_name and st.session_state.reader_name not in unique_analysts: unique_analysts.append(st.session_state.reader_name)
    
    if len(unique_analysts) == 2: names_str = f"{unique_analysts[0]} and {unique_analysts[1]}"
    elif len(unique_analysts) == 3: names_str = f"{unique_analysts[0]}, {unique_analysts[1]} and {unique_analysts[2]}"
    else: names_str = unique_analysts[0]
        
    smart_comment_interview = f"Yes, analysts {names_str} were interviewed comprehensively."
    smart_comment_samples = f"Yes, {st.session_state.sample_id}"
    smart_comment_records = f"Yes, See {tr_id} for more information."
    smart_comment_storage = f"Yes, Information is available in Eagle Trax Sample Location History under {st.session_state.sample_id}"

    # Phase I Summary (Part 1)
    p1 = f"All analysts involved in the prepping, processing, and reading of the samples – {names_str} – were interviewed and their answers are recorded throughout this document."
    p2 = f"The sample was stored upon arrival according to the Client’s instructions. Analysts {st.session_state.prepper_name} and {st.session_state.analyst_name} confirmed the integrity of the samples throughout both the preparation and processing stages. No leaks or turbidity were observed at any point, verifying the integrity of the sample."
    p3 = "All reagents and supplies mentioned in the material section above were stored according to the suppliers’ recommendations, and their integrity was visually verified before utilization. Moreover, each reagent and supply had valid expiration dates."
    p4 = f"During the preparation phase, {st.session_state.prepper_name} disinfected the samples using acidified bleach and placed them into a pre-disinfected storage bin. On {st.session_state.test_date}, prior to sample processing, {st.session_state.analyst_name} performed a second disinfection with acidified bleach, allowing a minimum contact time of 10 minutes before transferring the samples into the cleanroom suites. A final disinfection step was completed immediately before the samples were introduced into the ISO 5 Biological Safety Cabinet (BSC), E00{st.session_state.bsc_id}, located within the {t_loc}, (Suite {t_suite}{t_suffix}), All activities were performed in accordance with SOP 2.600.023, Rapid Scan RDI® Test Using FIFU Method."
    p5 = fresh_equip
    p6 = f"The analyst, {st.session_state.reader_name}, confirmed that the equipment was set up as per SOP 2.700.004 (Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup and Maintenance), and the negative control and the positive control for the analyst, {st.session_state.reader_name}, yielded expected results."
    
    smart_phase1_part1 = "\n\n".join([p1, p2, p3, p4, p5, p6])

    # Phase I Summary (Part 2)
    p7 = f"On {st.session_state.test_date}, a rapid sterility test was conducted on the sample using the ScanRDI method. The sample was initially prepared by Analyst {st.session_state.prepper_name}, processed by {st.session_state.analyst_name}, and subsequently read by {st.session_state.reader_name}. The test revealed {st.session_state.get('org_choice','')} {st.session_state.get('manual_org','')}-shaped viable microorganisms."
    p8 = f"Table 1 (see attached tables) presents the environmental monitoring results for {st.session_state.sample_id}. The environmental monitoring (EM) plates were incubated for no less than 48 hours at 30-35°C and no less than an additional five days at 20-25°C as per SOP 2.600.002 (Environmental Monitoring of the Clean-room Facility)."
    
    p9 = fresh_narr
    if fresh_det: p9 += "\n\n" + fresh_det

    p10 = f"Monthly cleaning and disinfection, using H₂O₂, of the cleanroom (ISO 7) and its containing Biosafety Cabinets (BSCs, ISO 5) were performed on {st.session_state.monthly_cleaning_date}, as per SOP 2.600.018 Cleaning and Disinfection Procedure. It was documented that all H₂O₂ indicators passed."
    p11 = fresh_history
    p12 = f"To assess the potential for sample-to-sample contamination contributing to the positive results, a comprehensive review was conducted of all samples processed on the same day. {fresh_cross}"
    p13 = "Based on the observations outlined above, it is unlikely that the failing results were due to reagents, supplies, the cleanroom environment, the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is minimal and the original result is deemed to be valid."

    smart_phase1_part2 = "\n\n".join([p7, p8, p9, p10, p11, p12, p13])

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

    # --- TABLE PDF GENERATOR (Internal) ---
    def create_appendix_pdf(data):
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(letter))
        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph(f"Appendix: Supplemental Tables for {data['sample_id']}", styles['Heading1']))
        elements.append(Spacer(1, 20))

        # TABLE 1
        elements.append(Paragraph(f"Table 1: Information for {data['sample_id']} under investigation", styles['Heading2']))
        t1_data = [
            ["Processing Analyst", "Reading Analyst", "Sample ID", "Events", "Confirmed Microbial Events", "Morphology Description"],
            [data['analyst_name'], data['reader_name'], data['sample_id'], "<1 event", "1", f"{data['organism_morphology']}-shaped Morphology"]
        ]
        t1 = Table(t1_data, colWidths=[120, 120, 120, 60, 100, 150])
        t1.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
        ]))
        elements.append(t1)
        elements.append(Spacer(1, 20))

        # TABLE 2
        elements.append(Paragraph(f"Table 2: Environmental Monitoring from Testing Performed on {data['test_date']}", styles['Heading2']))
        headers = ["Sampling Site", "Freq", "Date", "Analyst", "Observation", "Plate ETX ID", "Microbial ID", "Notes"]
        rows = []
        rows.append(["Personnel EM Bracketing", "", "", "", "", "", "", ""])
        rows.append(["Personal (Left/Right)", "Daily", data['test_date'], data['analyst_initial'], data['obs_pers_dur'], data['etx_pers_dur'], data['id_pers_dur'], "None"])
        rows.append([f"BSC EM Bracketing ({data['bsc_id']})", "", "", "", "", "", "", ""])
        rows.append(["Surface Sampling (ISO 5)", "Daily", data['test_date'], data['analyst_initial'], data['obs_surf_dur'], data['etx_surf_dur'], data['id_surf_dur'], "None"])
        rows.append(["Settling Sampling (ISO 5)", "Daily", data['test_date'], data['analyst_initial'], data['obs_sett_dur'], data['etx_sett_dur'], data['id_sett_dur'], "None"])
        rows.append([f"Weekly Bracketing (CR {data['cr_id']})", "", "", "", "", "", "", ""])
        rows.append(["Active Air Sampling", "Weekly", data['date_of_weekly'], data['weekly_initial'], data['obs_air_wk_of'], data['etx_air_wk_of'], data['id_air_wk_of'], "None"])
        rows.append(["Surface Sampling", "Weekly", data['date_of_weekly'], data['weekly_initial'], data['obs_room_wk_of'], data['etx_room_wk_of'], data['id_room_wk_of'], "None"])

        t2 = Table([headers] + rows, colWidths=[150, 60, 80, 50, 80, 100, 120, 60])
        t2.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('BACKGROUND', (0, 1), (-1, 1), colors.whitesmoke),
            ('BACKGROUND', (0, 3), (-1, 3), colors.whitesmoke),
            ('BACKGROUND', (0, 6), (-1, 6), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('SPAN', (0, 1), (-1, 1)),
            ('SPAN', (0, 3), (-1, 3)),
            ('SPAN', (0, 6), (-1, 6)),
        ]))
        elements.append(t2)
        doc.build(elements)
        buffer.seek(0)
        return buffer

    # 4. Generate DOCX
    docx_template = "ScanRDI OOS template 0.docx" 
    if os.path.exists(docx_template):
        try:
            doc = DocxTemplate(docx_template)
            doc.render(final_data_docx)
            buf = io.BytesIO(); doc.save(buf); buf.seek(0)
            st.download_button("📂 Download DOCX", buf, f"OOS-{clean_filename(st.session_state.oos_id)}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e: st.error(f"DOCX Error: {e}")
    else: st.warning(f"⚠️ Template file '{docx_template}' not found.")

    # 5. Generate PDF (MERGED)
    if os.path.exists("ScanRDI OOS template.pdf"):
        try:
            # Fill Form
            writer = PdfWriter(clone_from="ScanRDI OOS template.pdf")
            for p in writer.pages: writer.update_page_form_field_values(p, pdf_map)
            form_buf = io.BytesIO(); writer.write(form_buf); form_buf.seek(0)
            
            # Generate Tables
            tables_buf = create_appendix_pdf(final_data_docx)
            
            # Merge
            merger = PdfMerger()
            merger.append(form_buf)
            merger.append(tables_buf)
            
            final_buf = io.BytesIO()
            merger.write(final_buf)
            final_buf.seek(0)
            
            st.download_button("📂 Download PDF (Merged with Tables)", final_buf, f"OOS-{clean_filename(st.session_state.oos_id)}.pdf", "application/pdf")
        except Exception as e: st.error(f"PDF Error: {e}")
