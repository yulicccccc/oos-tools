import streamlit as st
from docxtpl import DocxTemplate
from pypdf import PdfReader, PdfWriter
import os
import re
import json
import io
from datetime import datetime, timedelta
from utils import apply_eagle_style, get_full_name, get_room_logic

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

# --- FILE PERSISTENCE (MEMORY) ---
STATE_FILE = "investigation_state.json"

# ALL keys that need to be saved/loaded
field_keys = [
    "oos_id", "client_name", "sample_id", "test_date", "sample_name", "lot_number", 
    "dosage_form", "monthly_cleaning_date", 
    "prepper_initial", "prepper_name", 
    "analyst_initial", "analyst_name",
    "changeover_initial", "changeover_name",
    "reader_initial", "reader_name",
    "bsc_id", "chgbsc_id", "scan_id", 
    "shift_number", "active_platform",
    "org_choice", "manual_org", "test_record", "control_pos", "control_lot", 
    "control_exp", "obs_pers", "etx_pers", "id_pers", "obs_surf", "etx_surf", 
    "id_surf", "obs_sett", "etx_sett", "id_sett", "obs_air", "etx_air_weekly", 
    "id_air_weekly", "obs_room", "etx_room_weekly", "id_room_wk_of", "weekly_init", 
    "date_weekly", "equipment_summary", "narrative_summary", "em_details", 
    "sample_history_paragraph", "incidence_count", "oos_refs",
    "other_positives", "cross_contamination_summary",
    "total_pos_count_num", "current_pos_order",
    "diff_changeover_bsc", "has_prior_failures",
    "em_growth_observed", "diff_changeover_analyst",
    "diff_reader_analyst",
    "em_growth_count" 
]
# Add dynamic keys for lists
for i in range(20):
    field_keys.append(f"other_id_{i}")
    field_keys.append(f"other_order_{i}")
    field_keys.append(f"prior_oos_{i}")
    field_keys.append(f"em_cat_{i}")
    field_keys.append(f"em_obs_{i}")
    field_keys.append(f"em_etx_{i}")
    field_keys.append(f"em_id_{i}")

def load_saved_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r") as f:
                saved_data = json.load(f)
            for key, value in saved_data.items():
                if key in st.session_state:
                    st.session_state[key] = value
        except Exception as e:
            st.error(f"Could not load saved state: {e}")

def save_current_state():
    data_to_save = {k: v for k, v in st.session_state.items() if k in field_keys}
    try:
        with open(STATE_FILE, "w") as f:
            json.dump(data_to_save, f)
    except Exception as e:
        st.error(f"Could not save state: {e}")

# --- HELPER FUNCTIONS ---
def clean_filename(text):
    if not text: return ""
    clean = re.sub(r'[\\/*?:"<>|]', '_', str(text))
    return clean.strip()

def num_to_words(n):
    mapping = {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six", 7: "seven", 8: "eight", 9: "nine", 10: "ten"}
    return mapping.get(n, str(n))

def ordinal(n):
    try: n = int(n)
    except: return str(n)
    if 11 <= (n % 100) <= 13: suffix = 'th'
    else: suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"

# --- GENERATE LIVE TEXTS ---
def generate_equipment_text():
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
    
    if st.session_state.bsc_id == st.session_state.chgbsc_id:
        part1 = f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}."
        part2 = f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both testing and changeover steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{st.session_state.bsc_id} was certified and approved by both the Engineering and Quality Assurance teams. Sample processing and changeover were conducted in the ISO 5 BSC E00{st.session_state.bsc_id} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {st.session_state.analyst_name} on {st.session_state.test_date}."
        return f"{part1}\n\n{part2}"
    else:
        part1 = f"The cleanroom used for testing (E00{t_room}) consists of three interconnected sections..."
        part2 = f"The ISO 5 BSC E00{st.session_state.bsc_id} and ISO 5 BSC E00{st.session_state.chgbsc_id} were thoroughly cleaned..."
        return f"{part1}\n\n{part2}"

def generate_history_text():
    if st.session_state.incidence_count == 0: 
        hist_phrase = "no prior failures"
    else:
        prior_ids = []
        for i in range(st.session_state.incidence_count):
            pid = st.session_state.get(f"prior_oos_{i}", "").strip()
            if pid: prior_ids.append(pid)
        
        if not prior_ids: refs_str = "..."
        elif len(prior_ids) == 1: refs_str = prior_ids[0]
        else: refs_str = ", ".join(prior_ids[:-1]) + f", and {prior_ids[-1]}"
        
        hist_phrase = f"{st.session_state.incidence_count} incident(s) ({refs_str})"
            
    return f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte “{st.session_state.sample_name}” has had {hist_phrase} using the Scan RDI method during this period."

def generate_cross_contam_text():
    if st.session_state.other_positives == "No":
        return "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
    else:
        num_others = st.session_state.total_pos_count_num - 1
        return f"{num_others} other samples tested positive. The analyst verified that gloves were thoroughly disinfected between samples."

def generate_narrative_and_details():
    failures = []
    count = st.session_state.get("em_growth_count", 1)
    
    cat_map = { "Personnel Obs": "personnel sampling", "Surface Obs": "surface sampling", "Settling Obs": "settling plates", "Weekly Air Obs": "weekly active air sampling", "Weekly Surf Obs": "weekly surface sampling" }
    
    for i in range(count):
        cat_friendly = st.session_state.get(f"em_cat_{i}", "Personnel Obs")
        obs_val = st.session_state.get(f"em_obs_{i}", "")
        etx_val = st.session_state.get(f"em_etx_{i}", "")
        id_val = st.session_state.get(f"em_id_{i}", "")
        category = cat_map.get(cat_friendly, "personnel sampling")
        if "weekly" in category: time_ctx = "weekly"
        else: time_ctx = "daily"
        if obs_val.strip():
            failures.append({"cat": category, "obs": obs_val, "etx": etx_val, "id": id_val, "time": time_ctx})

    if not failures:
        narr = "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, and settling plates. Additionally, weekly active air sampling and weekly surface sampling showed no microbial growth."
        det = ""
    else:
        narr = "Upon analyzing the environmental monitoring results, microbial growth was observed."
        det = "Microbial growth was observed. " + " ".join([f"{f['obs']} was detected during {f['cat']} (ID: {f['id']})." for f in failures])

    return narr, det, failures

# --- INIT STATE ---
def init_state(key, default_value=""):
    if key not in st.session_state: st.session_state[key] = default_value

for k in field_keys:
    if k == "incidence_count": init_state(k, 0)
    elif k == "shift_number": init_state(k, "1")
    elif "etx" in k or "id" in k: init_state(k, "N/A")
    elif k == "active_platform": init_state(k, "ScanRDI")
    elif k == "other_positives": init_state(k, "No")
    elif k == "total_pos_count_num": init_state(k, 2)
    elif k == "current_pos_order": init_state(k, 1) 
    elif k == "diff_changeover_bsc": init_state(k, "No")
    elif k == "has_prior_failures": init_state(k, "No")
    elif k == "em_growth_observed": init_state(k, "No")
    elif k == "diff_changeover_analyst": init_state(k, "No")
    elif k == "diff_reader_analyst": init_state(k, "No") 
    elif k == "em_growth_count": init_state(k, 1) 
    elif k.startswith("other_order_"): init_state(k, 1)
    else: init_state(k, "")

if "data_loaded" not in st.session_state:
    load_saved_state()
    st.session_state.data_loaded = True

# --- EMAIL PARSER ---
def parse_email_text(text):
    oos_match = re.search(r"OOS-(\d+)", text)
    if oos_match: st.session_state.oos_id = oos_match.group(1)
    client_match = re.search(r"([A-Za-z\s]+\(E\d+\))", text)
    if client_match: st.session_state.client_name = client_match.group(1).strip()
    etx_id_match = re.search(r"(ETX-\d{6}-\d{4})", text)
    if etx_id_match: st.session_state.sample_id = etx_id_match.group(1).strip()
    sample_match = re.search(r"Sample\s*Name:\s*(.*)", text, re.IGNORECASE)
    if sample_match: st.session_state.sample_name = sample_match.group(1).strip()
    lot_match = re.search(r"(?:Lot|Batch)\s*[:\.]?\s*([^\n\r]+)", text, re.IGNORECASE)
    if lot_match: st.session_state.lot_number = lot_match.group(1).strip()
    
    date_match = re.search(r"testing\s*on\s*(\d{2}\s*\w{3}\s*\d{4})", text, re.IGNORECASE)
    if date_match:
        try: d_obj = datetime.strptime(date_match.group(1).strip(), "%d %b %Y"); st.session_state.test_date = d_obj.strftime("%d%b%y")
        except: pass
    analyst_match = re.search(r"\(\s*([A-Z]{2,3})\s*\d+[a-z]{2}\s*Sample\)", text)
    if analyst_match: 
        st.session_state.analyst_initial = analyst_match.group(1).strip()
        st.session_state.analyst_name = get_full_name(st.session_state.analyst_initial)
    save_current_state()

st.title("🦠 ScanRDI Investigation")

# --- SMART PARSER ---
st.header("📧 Smart Email Import")
email_input = st.text_area("Paste the OOS Notification email here to auto-fill fields:", height=150)
if st.button("🪄 Parse Email & Auto-Fill"):
    if email_input: parse_email_text(email_input); st.success("Fields updated!"); st.rerun()

# --- FORM INPUTS ---
st.header("1. General Test Details")
col1, col2, col3 = st.columns(3)
with col1:
    st.text_input("OOS Number (Numbers only)", key="oos_id", help="Required")
    st.text_input("Client Name", key="client_name", help="Required")
    st.text_input("Sample ID (ETX Format)", key="sample_id", help="Required")
with col2:
    st.text_input("Test Date (e.g., 07Jan26)", key="test_date", help="Required")
    st.text_input("Sample / Active Name", key="sample_name", help="Required")
    st.text_input("Lot Number", key="lot_number", help="Required")
with col3:
    dosage_options = ["Injectable", "Aqueous Solution", "Liquid", "Solution"]
    st.selectbox("Dosage Form", dosage_options, key="dosage_form", index=0 if st.session_state.dosage_form not in dosage_options else dosage_options.index(st.session_state.dosage_form))
    st.text_input("Monthly Cleaning Date", key="monthly_cleaning_date", help="Required")

st.header("2. Personnel")
p1, p2 = st.columns(2)
with p1:
    st.text_input("Prepper Initials", key="prepper_initial")
    if st.session_state.prepper_initial and not st.session_state.prepper_name:
        st.session_state.prepper_name = get_full_name(st.session_state.prepper_initial)
    st.text_input("Prepper Full Name", key="prepper_name", help="Required")
with p2:
    st.text_input("Processor Initials", key="analyst_initial")
    if st.session_state.analyst_initial and not st.session_state.analyst_name:
        st.session_state.analyst_name = get_full_name(st.session_state.analyst_initial)
    st.text_input("Processor Full Name", key="analyst_name", help="Required")

st.session_state.diff_reader_analyst = st.radio("Was the Reading performed by a different analyst?", ["No", "Yes"], index=0 if st.session_state.diff_reader_analyst == "No" else 1, horizontal=True)
if st.session_state.diff_reader_analyst == "Yes":
    c1, c2 = st.columns(2)
    with c1: st.text_input("Reader Initials", key="reader_initial")
    with c2: 
        if st.session_state.reader_initial and not st.session_state.reader_name:
            st.session_state.reader_name = get_full_name(st.session_state.reader_initial)
        st.text_input("Reader Full Name", key="reader_name")
else:
    st.session_state.reader_initial = st.session_state.analyst_initial
    st.session_state.reader_name = st.session_state.analyst_name

st.session_state.diff_changeover_analyst = st.radio("Was the Changeover performed by a different analyst?", ["No", "Yes"], index=0 if st.session_state.diff_changeover_analyst == "No" else 1, horizontal=True)
if st.session_state.diff_changeover_analyst == "Yes":
    c1, c2 = st.columns(2)
    with c1: st.text_input("Changeover Initials", key="changeover_initial")
    with c2: 
        if st.session_state.changeover_initial and not st.session_state.changeover_name:
            st.session_state.changeover_name = get_full_name(st.session_state.changeover_initial)
        st.text_input("Changeover Full Name", key="changeover_name")
else:
    st.session_state.changeover_initial = st.session_state.analyst_initial
    st.session_state.changeover_name = st.session_state.analyst_name

st.divider()
e1, e2 = st.columns(2)
bsc_list = ["1310", "1309", "1311", "1312", "1314", "1313", "1316", "1798", "Other"]
with e1:
    st.selectbox("Select Processing BSC ID", bsc_list, key="bsc_id", index=0 if st.session_state.bsc_id not in bsc_list else bsc_list.index(st.session_state.bsc_id))
with e2:
    st.radio("Was the Changeover performed in a different BSC?", ["No", "Yes"], key="diff_changeover_bsc", horizontal=True)
    if st.session_state.diff_changeover_bsc == "Yes":
        st.selectbox("Select Changeover BSC ID", bsc_list, key="chgbsc_id", index=0 if st.session_state.chgbsc_id not in bsc_list else bsc_list.index(st.session_state.chgbsc_id))
    else:
        st.session_state.chgbsc_id = st.session_state.bsc_id

st.header("3. Findings & EM Data")
f1, f2 = st.columns(2)
with f1:
    scan_ids = ["1230", "2017", "1040", "1877", "2225", "2132"]
    st.selectbox("ScanRDI ID", scan_ids, key="scan_id", index=0 if st.session_state.scan_id not in scan_ids else scan_ids.index(st.session_state.scan_id))
    st.text_input("Shift Number", key="shift_number", help="Required")
    shape_opts = ["rod", "cocci", "Other"]
    st.selectbox("Org Shape", shape_opts, key="org_choice", index=0 if st.session_state.org_choice not in shape_opts else shape_opts.index(st.session_state.org_choice))
    if st.session_state.org_choice == "Other": st.text_input("Enter Manual Org Shape", key="manual_org")
    try: d_obj = datetime.strptime(st.session_state.test_date, "%d%b%y").strftime("%m%d%y"); st.session_state.test_record = f"{d_obj}-{st.session_state.scan_id}-{st.session_state.shift_number}"
    except: pass
    st.text_input("Record Ref", st.session_state.test_record, disabled=True)
with f2:
    ctrl_opts = ["A. brasiliensis", "B. subtilis", "C. albicans", "C. sporogenes", "P. aeruginosa", "S. aureus"]
    st.selectbox("Positive Control", ctrl_opts, key="control_pos", index=0 if st.session_state.control_pos not in ctrl_opts else ctrl_opts.index(st.session_state.control_pos))
    st.text_input("Control Lot", key="control_lot", help="Required")
    st.text_input("Control Exp Date", key="control_exp", help="Required")

st.header("4. EM Observations")
st.radio("Was microbial growth observed in Environmental Monitoring?", ["No", "Yes"], key="em_growth_observed", horizontal=True)

if st.session_state.em_growth_observed == "Yes":
    count = st.number_input("Number of EM Failures", min_value=1, step=1, key="em_growth_count")
    cat_options = ["Personnel Obs", "Surface Obs", "Settling Obs", "Weekly Air Obs", "Weekly Surf Obs"]
    for i in range(count):
        st.subheader(f"Growth #{i+1}")
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.selectbox(f"Category", cat_options, key=f"em_cat_{i}")
        with c2: st.text_input(f"Observation", key=f"em_obs_{i}")
        with c3: st.text_input(f"ETX #", key=f"em_etx_{i}")
        with c4: st.text_input(f"Microbial ID", key=f"em_id_{i}")

st.divider()
st.caption("Weekly Bracketing")
m1, m2 = st.columns(2)
with m1: st.text_input("Weekly Monitor Initials", key="weekly_init", help="Required")
with m2: st.text_input("Date of Weekly Monitoring", key="date_weekly", help="Required")

if st.session_state.em_growth_observed == "Yes":
    if st.button("🔄 Generate Narrative & Details"):
        n, d, failures = generate_narrative_and_details()
        st.session_state.narrative_summary = n
        st.session_state.em_details = d
        st.rerun()
    st.text_area("Narrative Content", key="narrative_summary", height=120)
    st.text_area("Details Content", key="em_details", height=200)
else:
    st.session_state.narrative_summary = "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, and settling plates. Additionally, weekly active air sampling and weekly surface sampling showed no microbial growth."
    st.session_state.em_details = ""

st.header("5. Automated Summaries & Analysis")
st.subheader("Sample History")
st.radio("Were there any prior failures in the last 6 months?", ["No", "Yes"], key="has_prior_failures", horizontal=True)
if st.session_state.has_prior_failures == "Yes":
    count = st.number_input("Number of Prior Failures", min_value=1, step=1, key="incidence_count")
    for i in range(count):
        if f"prior_oos_{i}" not in st.session_state: st.session_state[f"prior_oos_{i}"] = ""
        st.text_input(f"Prior Failure #{i+1} OOS ID", key=f"prior_oos_{i}")
    if st.button("🔄 Generate History Text"):
        st.session_state.sample_history_paragraph = generate_history_text()
        st.rerun()
    st.text_area("History Text", key="sample_history_paragraph", height=120)
else:
    st.session_state.sample_history_paragraph = f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte “{st.session_state.sample_name}” has had no prior failures using the Scan RDI method during this period."

st.divider()
st.subheader("Cross-Contamination Analysis")
st.radio("Did other samples test positive on the same day?", ["No", "Yes"], key="other_positives", horizontal=True)
if st.session_state.other_positives == "Yes":
    st.number_input("Total # of Positive Samples that day", min_value=2, step=1, key="total_pos_count_num")
    st.number_input(f"Order of THIS Sample", min_value=1, step=1, key="current_pos_order")
    num_others = st.session_state.total_pos_count_num - 1
    for i in range(num_others):
        sub_c1, sub_c2 = st.columns(2)
        with sub_c1: st.text_input(f"Other Sample #{i+1} ID", key=f"other_id_{i}")
        with sub_c2: st.number_input(f"Order", min_value=1, step=1, key=f"other_order_{i}")
    if st.button("🔄 Generate Cross-Contam Text"):
        st.session_state.cross_contamination_summary = generate_cross_contam_text()
        st.rerun()
    st.text_area("Cross-Contam Text", key="cross_contamination_summary", height=250)
else:
    st.session_state.cross_contamination_summary = "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."

save_current_state()

# --- FINAL GENERATION ---
st.divider()
if st.button("🚀 GENERATE FINAL REPORT"):
    # ----------------------------------------------------
    # 1. ROBUST VALIDATION (Stops execution if failed)
    # ----------------------------------------------------
    req_map = {
        "OOS Number": "oos_id", "Client Name": "client_name", "Sample ID": "sample_id",
        "Test Date": "test_date", "Sample Name": "sample_name", "Lot Number": "lot_number",
        "Monthly Cleaning": "monthly_cleaning_date", "Prepper Name": "prepper_name",
        "Analyst Name": "analyst_name", "BSC ID": "bsc_id", "ScanRDI ID": "scan_id",
        "Shift Number": "shift_number", "Control Lot": "control_lot", "Control Exp": "control_exp",
        "Weekly Initials": "weekly_init", "Weekly Date": "date_weekly"
    }
    missing = []
    for lbl, k in req_map.items():
        if not str(st.session_state.get(k, "")).strip() or st.session_state.get(k) == "N/A":
            missing.append(lbl)

    if missing:
        st.error(f"🛑 STOP! You are missing required fields: {', '.join(missing)}")
        st.stop() # This halts the script here.

    # ----------------------------------------------------
    # 2. GENERATE SMART VARIABLES (For Word & PDF)
    # ----------------------------------------------------
    
    # Calculate derived values
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
    
    st.session_state.equipment_summary = generate_equipment_text()
    
    # Create the "Smart" text blocks expected by the Word Template
    smart_personnel_block = f"{st.session_state.analyst_name}"
    smart_incident_opening = f"An OOS result was obtained for {st.session_state.sample_name} (Lot: {st.session_state.lot_number}) on {st.session_state.test_date}."
    smart_comment_interview = f"Yes, {st.session_state.analyst_name} was interviewed regarding the testing event. No anomalies were noted."
    smart_comment_samples = "Yes, the sample identification and lot number were verified against the test request."
    smart_comment_records = "Yes, all raw data and calculations have been verified."
    smart_comment_storage = "Yes, the sample was stored at ambient temperature prior to testing."
    smart_phase1_summary = st.session_state.narrative_summary
    smart_phase1_continued = st.session_state.em_details if st.session_state.em_growth_observed == "Yes" else ""
    
    # Combine everything into one dictionary for the Template Render
    final_data = {k: v for k, v in st.session_state.items()}
    final_data.update({
        "analyst_signature": st.session_state.analyst_name,
        "report_header": st.session_state.sample_id,
        "smart_personnel_block": smart_personnel_block,
        "smart_incident_opening": smart_incident_opening,
        "smart_comment_interview": smart_comment_interview,
        "smart_comment_samples": smart_comment_samples,
        "smart_comment_records": smart_comment_records,
        "smart_comment_storage": smart_comment_storage,
        "smart_phase1_summary": smart_phase1_summary,
        "smart_phase1_continued": smart_phase1_continued,
        # Room logic fields
        "cr_id": t_room, "cr_suit": t_suite, "suit": t_suffix, "bsc_location": t_loc
    })

    # ----------------------------------------------------
    # 3. GENERATE PDF
    # ----------------------------------------------------
    pdf_template = "ScanRDI OOS template.pdf"
    if os.path.exists(pdf_template):
        try:
            writer = PdfWriter(clone_from=pdf_template)
            
            # Map Python keys to PDF Field Names
            # We map the "Smart" variables here too so PDF matches DOCX
            pdf_data = {
                'Text Field57': st.session_state.oos_id,
                'Date Field0':  st.session_state.test_date,
                'Text Field2':  st.session_state.sample_id,
                'Text Field6':  st.session_state.lot_number,
                'Text Field3':  smart_personnel_block,     # Use smart var
                'Text Field5':  st.session_state.dosage_form,
                'Text Field4':  st.session_state.sample_name,
                
                # Big Text Blocks
                'Text Field49': smart_phase1_summary,
                'Text Field50': smart_phase1_continued,
                
                # Personnel
                'Text Field26': st.session_state.prepper_name,   
                'Text Field27': st.session_state.reader_name,    
                
                # Equipment 
                'Text Field30': st.session_state.scan_id,        
                'Text Field32': st.session_state.bsc_id,
                
                # Investigation Questions (Mapping text to PDF fields)
                'Text Field10': smart_comment_interview, 
                'Text Field11': smart_comment_samples,
                'Text Field12': "Yes, as per SOP 2.600.023",
                'Text Field13': "Yes, as per SOP 2.600.023",
                'Text Field14': smart_comment_records
            }

            for page in writer.pages:
                writer.update_page_form_field_values(page, pdf_data)
            
            out_pdf = f"OOS-{clean_filename(st.session_state.oos_id)} {clean_filename(st.session_state.client_name)} - ScanRDI.pdf"
            with open(out_pdf, "wb") as f: writer.write(f)
            with open(out_pdf, "rb") as f:
                st.download_button(label="📂 Download PDF Report", data=f, file_name=out_pdf, mime="application/pdf")
        except Exception as e:
            st.warning(f"PDF Gen Error: {e}")

    # ----------------------------------------------------
    # 4. GENERATE DOCX
    # ----------------------------------------------------
    docx_template = "ScanRDI OOS template.docx"
    if os.path.exists(docx_template):
        try:
            doc = DocxTemplate(docx_template)
            doc.render(final_data) # Uses the dictionary with smart variables
            out_docx = f"OOS-{clean_filename(st.session_state.oos_id)} {clean_filename(st.session_state.client_name)} - ScanRDI.docx"
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button(label="📂 Download DOCX Report", data=buf, file_name=out_docx, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"DOCX Gen Error: {e}")
