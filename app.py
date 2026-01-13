import streamlit as st
from docxtpl import DocxTemplate
from pypdf import PdfReader, PdfWriter
import os
import re
import json
import io
from datetime import datetime, timedelta

# --- PAGE CONFIG ---
st.set_page_config(page_title="LabOps Smart Report Tool", layout="wide")

# --- CUSTOM STYLING ---
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #f0f2f6; }
    .main { background-color: #ffffff; }
    .stTextArea textarea { background-color: #ffffff; color: #31333F; border: 1px solid #d6d6d6; }
    </style>
    """, unsafe_allow_html=True)

# --- FILE PERSISTENCE ---
STATE_FILE = "investigation_state.json"

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
    "date_weekly", "incident_description", "narrative_summary", "sample_history_paragraph",
    "cross_contamination_summary", "equipment_summary", "em_details", "em_growth_observed",
    "em_growth_count", "has_prior_failures", "incidence_count", "other_positives",
    "total_pos_count_num", "current_pos_order"
]
# Dynamic keys for extra rows
for i in range(20):
    field_keys.extend([f"other_id_{i}", f"other_order_{i}", f"prior_oos_{i}", 
                       f"em_cat_{i}", f"em_obs_{i}", f"em_etx_{i}", f"em_id_{i}"])

def load_saved_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r") as f:
                saved_data = json.load(f)
            for key, value in saved_data.items():
                if key in st.session_state: st.session_state[key] = value
        except: pass

def save_current_state():
    data_to_save = {k: v for k, v in st.session_state.items() if k in field_keys}
    try:
        with open(STATE_FILE, "w") as f: json.dump(data_to_save, f)
    except: pass

# --- HELPERS ---
def get_full_name(initials):
    lookup = {
        "HS": "Halaina Smith", "DS": "Devanshi Shah", "GS": "Gabbie Surber",
        "MRB": "Muralidhar Bythatagari", "KSM": "Karla Silva", "DT": "Debrework Tassew",
        "PG": "Pagan Gary", "GA": "Gerald Anyangwe", "DH": "Domiasha Harrison",
        "TK": "Tamiru Kotisso", "AO": "Ayomide Odugbesi", "CCD": "Cuong Du",
        "ES": "Alex Saravia", "MJ": "Mukyang Jang", "KA": "Kathleen Aruta",
        "SMO": "Simin Mohammad", "VV": "Varsha Subramanian", "CSG": "Clea S. Garza",
        "GL": "Guanchen Li", "QYC": "Qiyue Chen"
    }
    return lookup.get(initials.upper().strip(), "")

def num_to_words(n):
    mapping = {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six", 7: "seven", 8: "eight", 9: "nine", 10: "ten"}
    return mapping.get(n, str(n))

def ordinal(n):
    try: n = int(n)
    except: return str(n)
    if 11 <= (n % 100) <= 13: suffix = 'th'
    else: suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"

def get_room_logic(bsc_id):
    try:
        num = int(bsc_id)
        suffix = "B" if num % 2 == 0 else "A"
        location = "innermost ISO 7 room" if suffix == "B" else "middle ISO 7 buffer room"
    except: suffix, location = "B", "innermost ISO 7 room"
    
    if bsc_id in ["1310", "1309"]: suite = "117"
    elif bsc_id in ["1311", "1312"]: suite = "116"
    elif bsc_id in ["1314", "1313"]: suite = "115"
    elif bsc_id in ["1316", "1798"]: suite = "114"
    else: suite = "Unknown"
    room_map = {"117": "1739", "116": "1738", "115": "1737", "114": "1736"}
    return room_map.get(suite, "Unknown"), suite, suffix, location

# --- DATA GENERATORS ---
def generate_history_text():
    if st.session_state.incidence_count == 0: 
        hist_phrase = "no prior failures"
    else:
        prior_ids = [st.session_state.get(f"prior_oos_{i}", "").strip() for i in range(st.session_state.incidence_count)]
        prior_ids = [pid for pid in prior_ids if pid]
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
        details = []
        ids = [st.session_state.get(f"other_id_{i}", "") for i in range(num_others)] + [st.session_state.sample_id]
        ids = [i for i in ids if i]
        ids_str = ", ".join(ids[:-1]) + f" and {ids[-1]}" if len(ids) > 1 else ids[0]
        return f"{ids_str} were the {num_to_words(st.session_state.total_pos_count_num)} samples tested positive for microbial growth. The analyst confirmed that these samples were not processed concurrently, sequentially, or within the manifold run. Glove disinfection was verified. These findings suggest that cross-contamination is highly unlikely."

def generate_narrative_and_details():
    if st.session_state.em_growth_observed == "No":
        return "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, and settling plates. Additionally, weekly active air sampling and weekly surface sampling showed no microbial growth.", ""
    # Simplified logic for the EM summary
    return "Microbial growth was observed during environmental monitoring. (See specific details below).", "Specific EM Failure Details: " + ", ".join([st.session_state.get(f"em_obs_{i}", "") for i in range(st.session_state.em_growth_count)])

# --- INITIALIZE ---
for k in field_keys:
    if k not in st.session_state:
        if "count" in k: st.session_state[k] = 1
        elif "observed" in k or "failures" in k or "positives" in k: st.session_state[k] = "No"
        else: st.session_state[k] = ""
load_saved_state()

# --- UI ---
st.title(f"LabOps Smart Tool: {st.session_state.active_platform}")

st.header("1. General Details")
c1, c2, c3 = st.columns(3)
with c1:
    st.text_input("OOS Number", key="oos_id")
    st.text_input("Client Name", key="client_name")
    st.text_input("Sample ID", key="sample_id")
    st.text_input("Incident Description", key="incident_description")
with c2:
    st.text_input("Test Date (e.g. 07Jan26)", key="test_date")
    st.text_input("Sample Name", key="sample_name")
    st.text_input("Lot Number", key="lot_number")
with c3:
    st.text_input("Monthly Cleaning Date", key="monthly_cleaning_date")
    st.selectbox("Dosage Form", ["Injectable", "Aqueous Solution", "Liquid"], key="dosage_form")

st.header("2. Personnel")
p1, p2, p3, p4 = st.columns(4)
with p1: 
    st.text_input("Prepper Initials", key="prepper_initial")
    st.session_state.prepper_name = get_full_name(st.session_state.prepper_initial)
with p2: 
    st.text_input("Processor Initials", key="analyst_initial")
    st.session_state.analyst_name = get_full_name(st.session_state.analyst_initial)
with p3: 
    st.text_input("Changeover Initials", key="changeover_initial")
    st.session_state.changeover_name = get_full_name(st.session_state.changeover_initial)
with p4: 
    st.text_input("Reader Initials", key="reader_initial")
    st.session_state.reader_name = get_full_name(st.session_state.reader_initial)

st.header("3. Equipment & Controls")
e1, e2, e3 = st.columns(3)
with e1: st.selectbox("ScanRDI ID", ["1230", "2017", "1040", "1877"], key="scan_id")
with e2: st.selectbox("BSC ID", ["1310", "1309", "1311", "1312", "1314", "1313"], key="bsc_id")
with e3: 
    st.text_input("Positive Control", key="control_pos")
    st.text_input("Control Lot", key="control_lot")
    st.text_input("Control Exp", key="control_exp")

st.header("4. EM Observations")
st.radio("Microbial growth in EM?", ["No", "Yes"], key="em_growth_observed", horizontal=True)
if st.session_state.em_growth_observed == "Yes":
    st.number_input("How many failures?", min_value=1, key="em_growth_count")
    for i in range(st.session_state.em_growth_count):
        ec1, ec2, ec3 = st.columns(3)
        with ec1: st.selectbox(f"Category #{i+1}", ["Personnel", "Surface", "Air"], key=f"em_cat_{i}")
        with ec2: st.text_input(f"Observation #{i+1}", key=f"em_obs_{i}")
        with ec3: st.text_input(f"Microbial ID #{i+1}", key=f"em_id_{i}")

st.header("5. History & Analysis")
h1, h2 = st.columns(2)
with h1:
    st.radio("Prior failures (6 mo)?", ["No", "Yes"], key="has_prior_failures", horizontal=True)
    if st.session_state.has_prior_failures == "Yes":
        st.number_input("Incidence count", min_value=1, key="incidence_count")
        for i in range(st.session_state.incidence_count):
            st.text_input(f"Prior OOS #{i+1}", key=f"prior_oos_{i}")
with h2:
    st.radio("Other positives same day?", ["No", "Yes"], key="other_positives", horizontal=True)
    if st.session_state.other_positives == "Yes":
        st.number_input("Total positives", min_value=2, key="total_pos_count_num")
        for i in range(st.session_state.total_pos_count_num - 1):
            st.text_input(f"Other Sample ID #{i+1}", key=f"other_id_{i}")

# --- GENERATE ---
st.divider()
if st.button("🚀 GENERATE SMART REPORT"):
    save_current_state()
    cr_id, cr_suit, suit, bsc_loc = get_room_logic(st.session_state.bsc_id)
    
    # Run narrative engines
    st.session_state.narrative_summary, st.session_state.em_details = generate_narrative_and_details()
    st.session_state.sample_history_paragraph = generate_history_text()
    st.session_state.cross_contamination_summary = generate_cross_contam_text()
    
    try:
        d_obj = datetime.strptime(st.session_state.test_date, "%d%b%y")
        st.session_state.test_record = f"{d_obj.strftime('%m%d%y')}-{st.session_state.scan_id}-1"
    except: st.session_state.test_record = "N/A"

    # BUILD SMART VARIABLES
    unique_names = list(dict.fromkeys([n for n in [st.session_state.prepper_name, st.session_state.analyst_name, st.session_state.reader_name] if n]))
    analyst_str = " and ".join([", ".join(unique_names[:-1]), unique_names[-1]]) if len(unique_names) > 1 else unique_names[0]
    
    smart_personnel = f"Prepper: {st.session_state.prepper_name}\nProcessor: {st.session_state.analyst_name}\nChangeover: {st.session_state.changeover_name}\nReader: {st.session_state.reader_name}"
    smart_opening = f"On {st.session_state.test_date}, sample {st.session_state.sample_id} was found positive for viable microorganisms after ScanRDI testing."
    
    # FINAL PHASE 1 SUMMARY
    smart_phase1_summary = f"""All analysts involved – {analyst_str} – were interviewed. Integrity of {st.session_state.sample_id} was confirmed by {st.session_state.prepper_name} and {st.session_state.analyst_name}. Samples were disinfected per SOP on {st.session_state.test_date}. Processing occurred in BSC E00{st.session_state.bsc_id} in {bsc_loc} (Suite {cr_suit}{suit})."""

    # FINAL PHASE 1 CONTINUED
    smart_phase1_continued = f"""Testing on {st.session_state.test_date} revealed growth. {st.session_state.narrative_summary}\n\n{st.session_state.sample_history_paragraph}\n\n{st.session_state.cross_contamination_summary}"""

    # BUNDLE
    word_data = {
        "smart_personnel_block": smart_personnel, "smart_incident_opening": smart_opening,
        "smart_comment_interview": f"Yes, analysts {analyst_str} were interviewed.",
        "smart_comment_samples": f"Yes, {st.session_state.sample_id}",
        "smart_comment_sop": "Yes, per SOP 2.600.023", "smart_comment_records": f"See {st.session_state.test_record}",
        "smart_phase1_summary": smart_phase1_summary, "smart_phase1_continued": smart_phase1_continued,
        "smart_footer_na": f"N/A QYC {datetime.today().strftime('%d%b%y')}"
    }

    pdf_data = {
        'Text Field8': st.session_state.oos_id, 'Text Field3': smart_personnel, 
        'Text Field2': st.session_state.sample_id, 'Text Field7': smart_opening,
        'Text Field49': smart_phase1_summary, 'Text Field50': smart_phase1_continued,
        'Text Field13': word_data["smart_comment_interview"], 'Text Field14': word_data["smart_comment_samples"],
        'Text Field30': f"E00{st.session_state.scan_id}", 'Text Field32': f"E00{cr_id}"
    }

    # EXPORT
    if os.path.exists("ScanRDI OOS template.pdf"):
        writer = PdfWriter(clone_from="ScanRDI OOS template.pdf")
        for page in writer.pages: writer.update_page_form_field_values(page, pdf_data)
        out_pdf = f"OOS-{st.session_state.oos_id}_Report.pdf"
        with open(out_pdf, "wb") as f: writer.write(f)
        st.download_button("📂 Download PDF", open(out_pdf, "rb"), out_pdf)

    if os.path.exists("ScanRDI OOS template.docx"):
        doc = DocxTemplate("ScanRDI OOS template.docx")
        doc.render(word_data)
        buf = io.BytesIO()
        doc.save(buf)
        st.download_button("📂 Download Word", buf.getvalue(), f"OOS-{st.session_state.oos_id}_Report.docx")
