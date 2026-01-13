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
    "cross_contamination_summary", "equipment_summary"
]

def load_saved_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r") as f:
                saved_data = json.load(f)
            for key, value in saved_data.items():
                if key in st.session_state:
                    st.session_state[key] = value
        except: pass

def save_current_state():
    data_to_save = {k: v for k, v in st.session_state.items() if k in field_keys}
    try:
        with open(STATE_FILE, "w") as f:
            json.dump(data_to_save, f)
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

# --- INITIALIZE STATE ---
for k in field_keys:
    if k not in st.session_state: st.session_state[k] = ""
if "active_platform" not in st.session_state: st.session_state.active_platform = "ScanRDI"

load_saved_state()

# --- UI LAYOUT ---
st.title(f"LabOps Smart Tool: {st.session_state.active_platform}")

st.header("1. General Details")
c1, c2, c3 = st.columns(3)
with c1:
    st.text_input("OOS Number", key="oos_id")
    st.text_input("Client Name", key="client_name")
    st.text_input("Sample ID", key="sample_id")
    st.text_input("Incident Brief Description", key="incident_description")
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
    st.caption(f"Name: {st.session_state.prepper_name}")
with p2: 
    st.text_input("Processor Initials", key="analyst_initial")
    st.session_state.analyst_name = get_full_name(st.session_state.analyst_initial)
    st.caption(f"Name: {st.session_state.analyst_name}")
with p3: 
    st.text_input("Changeover Initials", key="changeover_initial")
    st.session_state.changeover_name = get_full_name(st.session_state.changeover_initial)
    st.caption(f"Name: {st.session_state.changeover_name}")
with p4: 
    st.text_input("Reader Initials", key="reader_initial")
    st.session_state.reader_name = get_full_name(st.session_state.reader_initial)
    st.caption(f"Name: {st.session_state.reader_name}")

st.header("3. Equipment & Controls")
e1, e2, e3 = st.columns(3)
with e1: st.selectbox("ScanRDI ID", ["1230", "2017", "1040", "1877"], key="scan_id")
with e2: st.selectbox("BSC ID", ["1310", "1309", "1311", "1312", "1314", "1313"], key="bsc_id")
with e3: 
    st.text_input("Positive Control", key="control_pos")
    st.text_input("Control Lot", key="control_lot")
    st.text_input("Control Exp", key="control_exp")

# Auto-generate Test Record ID
try:
    d_part = datetime.strptime(st.session_state.test_date, "%d%b%y").strftime("%m%d%y")
    st.session_state.test_record = f"{d_part}-{st.session_state.scan_id}-1"
except: st.session_state.test_record = "PENDING"

# --- GENERATION LOGIC ---
st.divider()
if st.button("🚀 GENERATE SMART REPORT"):
    save_current_state()
    cr_id, cr_suit, suit, bsc_loc = get_room_logic(st.session_state.bsc_id)
    org_morph = st.session_state.manual_org if st.session_state.org_choice == "Other" else st.session_state.org_choice

    # 1. SMART PERSONNEL BLOCK (Word & PDF Text Field 3)
    p_list = [
        f"Prepper: {st.session_state.prepper_name} ({st.session_state.prepper_initial})",
        f"Processor: {st.session_state.analyst_name} ({st.session_state.analyst_initial})",
        f"Changeover Processor: {st.session_state.changeover_name} ({st.session_state.changeover_initial})",
        f"Reader: {st.session_state.reader_name} ({st.session_state.reader_initial})"
    ]
    smart_personnel = "\n".join(p_list)

    # 2. SMART INTERVIEW COMMENT (Grammar Aware for PDF Text Field 13)
    unique_names = list(dict.fromkeys([st.session_state.prepper_name, st.session_state.analyst_name, st.session_state.reader_name]))
    if len(unique_names) == 3: interview_names = f"{unique_names[0]}, {unique_names[1]} and {unique_names[2]}"
    elif len(unique_names) == 2: interview_names = f"{unique_names[0]} and {unique_names[1]}"
    else: interview_names = unique_names[0]
    smart_comment_interview = f"Yes, analysts {interview_names} were interviewed comprehensively."

    # 3. SMART INCIDENT OPENING (PDF Text Field 7)
    smart_opening = f"On {st.session_state.test_date}, sample {st.session_state.sample_id} was found positive for viable microorganisms after ScanRDI testing."

    # 4. EQUIPMENT SUMMARY BUNDLE
    equip_summary = f"Sample processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {cr_suit}{suit}, BSC E00{st.session_state.bsc_id}) by {st.session_state.analyst_name} on {st.session_state.test_date}."

    # 5. SMART PHASE 1 SUMMARY (PDF Text Field 49)
    smart_phase1_summary = f"""All analysts involved in the prepping, processing, and reading of the samples – {interview_names} – were interviewed and their answers are recorded throughout this document. 

The sample was stored upon arrival according to the Client’s instructions. Analysts {st.session_state.prepper_name} and {st.session_state.analyst_name} confirmed the integrity of the samples throughout both the preparation and processing stages. No leaks or turbidity were observed at any point, verifying the integrity of the sample.

All reagents and supplies mentioned in the material section above were stored according to the suppliers’ recommendations, and their integrity was visually verified before utilization. Moreover, each reagent and supply had valid expiration dates. 

During the preparation phase, {st.session_state.prepper_name} disinfected the samples using acidified bleach and placed them into a pre-disinfected storage bin. On {st.session_state.test_date}, prior to sample processing, {st.session_state.analyst_name} performed a second disinfection with acidified bleach, allowing a minimum contact time of 10 minutes before transferring the samples into the cleanroom suites. A final disinfection step was completed immediately before the samples were introduced into the ISO 5 Biological Safety Cabinet (BSC), E00{st.session_state.bsc_id}, located within the {bsc_loc}, (Suite {cr_suit}{suit}), All activities were performed in accordance with SOP 2.600.023, Rapid Scan RDI® Test Using FIFU Method.

{equip_summary}

The analyst, {st.session_state.reader_name}, confirmed that the equipment was set up as per SOP 2.700.004 (Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup and Maintenance), and the negative control and the positive control for the analyst, {st.session_state.reader_name}, yielded expected results."""

    # 6. SMART PHASE 1 CONTINUED (PDF Text Field 50)
    smart_phase1_continued = f"""On {st.session_state.test_date}, a rapid sterility test was conducted on the sample using the ScanRDI method. The sample was initially prepared by Analyst {st.session_state.prepper_name}, processed by {st.session_state.analyst_name}, and subsequently read by {st.session_state.analyst_name}. The test revealed {org_morph}-shaped viable microorganisms.

Table 1 (see attached tables) presents the environmental monitoring results for {st.session_state.sample_id}. The environmental monitoring (EM) plates were incubated for no less than 48 hours at 30-35°C and no less than an additional five days at 20-25°C as per SOP 2.600.002 (Environmental Monitoring of the Clean-room Facility).

{st.session_state.narrative_summary}

Monthly cleaning and disinfection, using H₂O₂, of the cleanroom (ISO 7) and its containing Biosafety Cabinets (BSCs, ISO 5) were performed on {st.session_state.monthly_cleaning_date}, as per SOP 2.600.018 Cleaning and Disinfection Procedure. It was documented that all H₂O₂ indicators passed. 

{st.session_state.sample_history_paragraph}

To assess the potential for sample-to-sample contamination contributing to the positive results, a comprehensive review was conducted of all samples processed on the same day. {st.session_state.cross_contamination_summary}

Based on the observations outlined above, it is unlikely that the failing results were due to reagents, supplies, the cleanroom environment, the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is minimal. Therefore, the original test result is deemed valid."""

    # --- FINAL BUNDLE FOR WORD ---
    word_data = {
        "full_case_ref": f"{st.session_state.sample_id} - {st.session_state.client_name}",
        "smart_personnel_block": smart_personnel,
        "smart_incident_opening": smart_opening,
        "smart_comment_interview": smart_comment_interview,
        "smart_comment_samples": f"Yes, {st.session_state.sample_id}",
        "smart_comment_sop": "Yes, as per SOP 2.600.023, 2.700.004",
        "smart_comment_records": f"Yes, See {st.session_state.test_record} for more information.",
        "smart_comment_storage": f"Yes, Information is available in Eagle Trax Sample Location History under {st.session_state.sample_id}",
        "smart_phase1_summary": smart_phase1_summary,
        "smart_phase1_continued": smart_phase1_continued,
        "smart_footer_na": f"N/A QYC {datetime.today().strftime('%d%b%y')}"
    }

    # --- FINAL BUNDLE FOR PDF (PLATFORM -> PDF MAPPING) ---
    pdf_data = {
        'Text Field8':  st.session_state.oos_id,
        'Text Field0':  st.session_state.analyst_name,
        'Date Field0':  st.session_state.test_date,
        'Date Field1':  st.session_state.test_date,
        'Date Field2':  st.session_state.test_date,
        'Text Field1':  "Scan RDI Sterility Test",
        'Text Field2':  st.session_state.sample_id,
        'Text Field3':  smart_personnel,             # Personnel Block
        'Text Field4':  st.session_state.sample_name,
        'Text Field5':  st.session_state.dosage_form,
        'Text Field6':  st.session_state.lot_number,
        'Text Field7':  smart_opening,               # Description of Incident
        'Text Field12': "Kathan Parikh",
        'Date Field3':  st.session_state.test_date,
        'Text Field13': smart_comment_interview,
        'Text Field14': word_data["smart_comment_samples"],
        'Text Field15': word_data["smart_comment_sop"],
        'Text Field16': word_data["smart_comment_sop"],
        'Text Field17': word_data["smart_comment_records"],
        'Text Field18': "Yes, all analysts are trained and qualified by quality to perform the test.",
        'Text Field19': "Not Applicable",
        'Text Field20': "Not Applicable",
        'Text Field21': word_data["smart_comment_storage"],
        'Text Field57': st.session_state.oos_id,
        'Text Field24': st.session_state.control_pos,
        'Text Field25': st.session_state.control_lot,
        'Text Field26': st.session_state.control_exp,
        'Text Field28': "Not Applicable",
        'Text Field30': f"E00{st.session_state.scan_id}",
        'Text Field31': "Oct26",
        'Text Field32': f"E00{cr_id} (CR{cr_suit})",
        'Text Field33': "Dec25",
        'Text Field34': f"E00{st.session_state.scan_id}",
        'Text Field35': "Oct26",
        'Text Field49': smart_phase1_summary,        # Page 3 Full Summary
        'Text Field50': smart_phase1_continued,      # Page 4 Full Continued
        'Text Field51': word_data["smart_footer_na"] # Page 6 Footer
    }

    # --- OUTPUT GENERATION ---
    # PDF
    pdf_template = "ScanRDI OOS template.pdf"
    if os.path.exists(pdf_template):
        writer = PdfWriter(clone_from=pdf_template)
        for page in writer.pages:
            writer.update_page_form_field_values(page, pdf_data)
        out_pdf = f"OOS-{st.session_state.oos_id}_Report.pdf"
        with open(out_pdf, "wb") as f: writer.write(f)
        with open(out_pdf, "rb") as f: st.download_button("📂 Download PDF", f, out_pdf)

    # WORD
    doc_template = "ScanRDI OOS template.docx"
    if os.path.exists(doc_template):
        doc = DocxTemplate(doc_template)
        doc.render(word_data)
        buf = io.BytesIO()
        doc.save(buf)
        st.download_button("📂 Download Word Doc", buf.getvalue(), f"OOS-{st.session_state.oos_id}_Report.docx")
