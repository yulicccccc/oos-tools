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
    "date_weekly", "equipment_summary", "narrative_summary", "em_details", 
    "sample_history_paragraph", "incidence_count", "oos_refs",
    "other_positives", "cross_contamination_summary",
    "total_pos_count_num", "current_pos_order",
    "diff_changeover_bsc", "has_prior_failures",
    "em_growth_observed", "diff_changeover_analyst",
    "diff_reader_analyst",
    "em_growth_count" 
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
def num_to_words(n): return {1:"one",2:"two",3:"three",4:"four",5:"five"}.get(n, str(n))
def ordinal(n):
    try: n = int(n)
    except: return str(n)
    if 11 <= (n % 100) <= 13: return f"{n}th"
    return f"{n}{{1:'st',2:'nd',3:'rd'}.get(n%10,'th')}"

# --- GENERATORS (Strict Old Template Logic) ---
def generate_equipment_text():
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    return f"The ISO 5 BSC E00{st.session_state.bsc_id}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for testing."

def generate_history_text():
    if st.session_state.incidence_count == 0: phrase = "no prior failures"
    else:
        pids = [st.session_state.get(f"prior_oos_{i}","") for i in range(st.session_state.incidence_count) if st.session_state.get(f"prior_oos_{i}")]
        if not pids: refs = "..."
        elif len(pids)==1: refs = pids[0]
        else: refs = ", ".join(pids[:-1]) + f", and {pids[-1]}"
        phrase = f"{st.session_state.incidence_count} incident(s) ({refs})"
    return f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte “{st.session_state.sample_name}” has had {phrase} using the Scan RDI method during this period."

def generate_cross_contam_text():
    if st.session_state.other_positives == "No": 
        return "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
    num = st.session_state.total_pos_count_num - 1
    return f"{num} other samples tested positive. The analyst verified that gloves were thoroughly disinfected between samples."

def generate_narrative_and_details():
    failures = []
    count = st.session_state.get("em_growth_count", 1)
    cat_map = {"Personnel Obs": "personnel sampling", "Surface Obs": "surface sampling", "Settling Obs": "settling plates", "Weekly Air Obs": "weekly active air sampling", "Weekly Surf Obs": "weekly surface sampling"}
    for i in range(count):
        cat = cat_map.get(st.session_state.get(f"em_cat_{i}"), "personnel sampling")
        obs = st.session_state.get(f"em_obs_{i}",""); etx = st.session_state.get(f"em_etx_{i}",""); mid = st.session_state.get(f"em_id_{i}","")
        if obs.strip(): failures.append({"cat": cat, "obs": obs, "etx": etx, "id": mid})
    
    if not failures:
        narr = "Upon analyzing the environmental monitoring results, no microbial growth was observed in personal sampling (left touch and right touch), surface sampling, and settling plates. Additionally, weekly active air sampling and weekly surface sampling showed no microbial growth."
        det = ""
    else:
        narr = "Upon analyzing the environmental monitoring results, microbial growth was observed."
        det = "Microbial growth was observed. " + " ".join([f"{f['obs']} was detected during {f['cat']} (ID: {f['id']})." for f in failures])
    return narr, det, failures

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
    if m := re.search(r"([A-Za-z\s]+\(E\d+\))", text): st.session_state.client_name = m.group(1).strip()
    if m := re.search(r"(ETX-\d{6}-\d{4})", text): st.session_state.sample_id = m.group(1).strip()
    if m := re.search(r"Sample\s*Name:\s*(.*)", text, re.I): st.session_state.sample_name = m.group(1).strip()
    if m := re.search(r"(?:Lot|Batch)\s*[:\.]?\s*([^\n\r]+)", text, re.I): st.session_state.lot_number = m.group(1).strip()
    if m := re.search(r"testing\s*on\s*(\d{2}\s*\w{3}\s*\d{4})", text, re.I):
        try: st.session_state.test_date = datetime.strptime(m.group(1).strip(), "%d %b %Y").strftime("%d%b%y")
        except: pass
    if m := re.search(r"\(\s*([A-Z]{2,3})\s*\d+[a-z]{2}\s*Sample\)", text): 
        st.session_state.analyst_initial = m.group(1).strip()
        st.session_state.analyst_name = get_full_name(st.session_state.analyst_initial)
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
    st.text_input("Shift Number", key="shift_number", help="Required")
    st.selectbox("Org Shape", ["rod","cocci","Other"], key="org_choice")
    if st.session_state.org_choice == "Other": st.text_input("Manual Shape", key="manual_org")
with f2:
    st.selectbox("Positive Control", ["A. brasiliensis","B. subtilis","C. albicans","C. sporogenes","P. aeruginosa","S. aureus"], key="control_pos")
    st.text_input("Control Lot", key="control_lot", help="Required")
    st.text_input("Control Exp", key="control_exp", help="Required")

st.header("4. EM Observations")
st.radio("Microbial Growth Observed?", ["No","Yes"], key="em_growth_observed", horizontal=True)
if st.session_state.em_growth_observed == "Yes":
    count = st.number_input("Failures Count", 1, 20, key="em_growth_count")
    for i in range(count):
        st.subheader(f"Growth #{i+1}")
        c1,c2,c3,c4 = st.columns(4)
        with c1: st.selectbox(f"Cat {i+1}", ["Personnel Obs","Surface Obs","Settling Obs","Weekly Air Obs","Weekly Surf Obs"], key=f"em_cat_{i}")
        with c2: st.text_input(f"Obs {i+1}", key=f"em_obs_{i}")
        with c3: st.text_input(f"ETX {i+1}", key=f"em_etx_{i}")
        with c4: st.text_input(f"ID {i+1}", key=f"em_id_{i}")

st.divider()
st.caption("Weekly Bracketing")
m1, m2 = st.columns(2)
with m1: st.text_input("Weekly Monitor Initials", key="weekly_init", help="Required")
with m2: st.text_input("Date of Weekly Monitoring", key="date_weekly", help="Required")

if st.button("🔄 Update Summaries"):
    n, d, fails = generate_narrative_and_details()
    st.session_state.narrative_summary = n
    st.session_state.em_details = d
    st.rerun()

st.text_area("Narrative", key="narrative_summary", height=100)
if st.session_state.em_growth_observed == "Yes": st.text_area("Details", key="em_details", height=150)

save_current_state()

st.divider()

if st.button("🚀 GENERATE FINAL REPORT"):
    # 1. Validation
    missing = []
    reqs = {"OOS #":"oos_id", "Client":"client_name", "Sample ID":"sample_id", "Date":"test_date", "Sample Name":"sample_name", "Lot":"lot_number", "Analyst":"analyst_name", "BSC":"bsc_id", "Scan ID":"scan_id"}
    for l,k in reqs.items():
        if not st.session_state.get(k,"").strip(): missing.append(l)
    if missing: st.error(f"Missing: {', '.join(missing)}"); st.stop()

    # 2. DEFINE SMART VARIABLES (Hardcoded 'Qiyue Chen')
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    c_room, c_suite, c_suffix, c_loc = get_room_logic(st.session_state.chgbsc_id)
    
    try: d_obj = datetime.strptime(st.session_state.test_date, "%d%b%y").strftime("%m%d%y"); tr_id = f"{d_obj}-{st.session_state.scan_id}-{st.session_state.shift_number}"
    except: tr_id = "N/A"

    # [FIXED] Analyst + (Written by: Qiyue Chen)
    analyst_sig_text = f"{st.session_state.analyst_name} (Written by: Qiyue Chen)"

    # [FIXED] Personnel Block
    smart_personnel_block = (
        f"Prepper: {st.session_state.prepper_name} ({st.session_state.prepper_initial}), "
        f"Processor: {st.session_state.analyst_name} ({st.session_state.analyst_initial}), "
        f"Changeover Processor: {st.session_state.changeover_name} ({st.session_state.changeover_initial}), "
        f"Reader: {st.session_state.reader_name} ({st.session_state.reader_initial})"
    )

    # [FIXED] Phase I Summary - 1:1 Match to Template 0
    p1 = f"All analysts involved in the prepping, processing, and reading of the samples – {st.session_state.prepper_name}, {st.session_state.analyst_name} and {st.session_state.reader_name} – were interviewed and their answers are recorded throughout this document."
    p2 = f"The sample was stored upon arrival according to the Client’s instructions. Analysts {st.session_state.prepper_name} and {st.session_state.analyst_name} confirmed the integrity of the samples throughout both the preparation and processing stages. No leaks or turbidity were observed at any point, verifying the integrity of the sample."
    p3 = "All reagents and supplies mentioned in the material section above were stored according to the suppliers’ recommendations, and their integrity was visually verified before utilization. Moreover, each reagent and supply had valid expiration dates."
    p4 = f"During the preparation phase, {st.session_state.prepper_name} disinfected the samples using acidified bleach and placed them into a pre-disinfected storage bin. On {st.session_state.test_date}, prior to sample processing, {st.session_state.analyst_name} performed a second disinfection with acidified bleach, allowing a minimum contact time of 10 minutes before transferring the samples into the cleanroom suites. A final disinfection step was completed immediately before the samples were introduced into the ISO 5 Biological Safety Cabinet (BSC), E00{st.session_state.bsc_id}, located within the {t_loc}, (Suite {t_suite}{t_suffix}), All activities were performed in accordance with SOP 2.600.023, Rapid Scan RDI® Test Using FIFU Method."
    p5 = generate_equipment_text()
    p6 = f"The analyst, {st.session_state.reader_name}, confirmed that the equipment was set up as per SOP 2.700.004 (Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup and Maintenance), and the negative control and the positive control for the analyst, {st.session_state.reader_name}, yielded expected results."
    p7 = f"On {st.session_state.test_date}, a rapid sterility test was conducted on the sample using the ScanRDI method. The sample was initially prepared by Analyst {st.session_state.prepper_name}, processed by {st.session_state.analyst_name}, and subsequently read by {st.session_state.analyst_name}. The test revealed {st.session_state.get('org_choice','')} {st.session_state.get('manual_org','')}-shaped viable microorganisms."
    p8 = f"Table 1 (see attached tables) presents the environmental monitoring results for {st.session_state.sample_id}. The environmental monitoring (EM) plates were incubated for no less than 48 hours at 30-35°C and no less than an additional five days at 20-25°C as per SOP 2.600.002 (Environmental Monitoring of the Clean-room Facility)."
    p9 = st.session_state.narrative_summary
    p10 = f"Monthly cleaning and disinfection, using H₂O₂, of the cleanroom (ISO 7) and its containing Biosafety Cabinets (BSCs, ISO 5) were performed on {st.session_state.monthly_cleaning_date}, as per SOP 2.600.018 Cleaning and Disinfection Procedure. It was documented that all H₂O₂ indicators passed."
    p11 = generate_history_text()
    p12 = f"To assess the potential for sample-to-sample contamination contributing to the positive results, a comprehensive review was conducted of all samples processed on the same day. {generate_cross_contam_text()}"
    p13 = "Based on the observations outlined above, it is unlikely that the failing results were due to reagents, supplies, the cleanroom environment, the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is minimal and the original result is deemed to be valid."

    smart_phase1_text = "\n\n".join([p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13])

    smart_data = {
        "analyst_signature": analyst_sig_text,
        "report_header": st.session_state.sample_id,
        "smart_personnel_block": smart_personnel_block,
        "smart_incident_opening": f"On {st.session_state.test_date}, sample {st.session_state.sample_id} was found positive for viable microorganisms after ScanRDI testing.",
        "smart_comment_interview": f"Yes, analysts {st.session_state.prepper_name}, {st.session_state.analyst_name} and {st.session_state.reader_name} were interviewed comprehensively.",
        "smart_comment_samples": f"Yes, {st.session_state.sample_id}",
        "smart_comment_records": f"Yes, See {tr_id} for more information.",
        "smart_comment_storage": f"Yes, Information is available in Eagle Trax Sample Location History under {st.session_state.sample_id}",
        "smart_phase1_summary": smart_phase1_text,
        "smart_phase1_continued": st.session_state.em_details if st.session_state.em_growth_observed == "Yes" else ""
    }

    final_data = {k: v for k, v in st.session_state.items()}
    final_data.update(smart_data)
    
    final_data.update({
        "cr_id": t_room, "cr_suit": t_suite, "suit": t_suffix, "bsc_location": t_loc,
        "obs_pers_dur": st.session_state.get("obs_pers") or "No Growth",
        "etx_pers_dur": st.session_state.get("etx_pers") or "N/A",
        "id_pers_dur": st.session_state.get("id_pers") or "N/A",
        "obs_surf_dur": st.session_state.get("obs_surf") or "No Growth",
        "etx_surf_dur": st.session_state.get("etx_surf") or "N/A",
        "id_surf_dur": st.session_state.get("id_surf") or "N/A",
        "obs_sett_dur": st.session_state.get("obs_sett") or "No Growth",
        "etx_sett_dur": st.session_state.get("etx_sett") or "N/A",
        "id_sett_dur": st.session_state.get("id_sett") or "N/A",
        "date_of_weekly": st.session_state.get("date_weekly", ""),
        "weekly_initial": st.session_state.get("weekly_init", ""),
        "obs_air_wk_of": st.session_state.get("obs_air") or "No Growth",
        "etx_air_wk_of": st.session_state.get("etx_air_weekly") or "N/A",
        "id_air_wk_of": st.session_state.get("id_air_weekly") or "N/A",
        "obs_room_wk_of": st.session_state.get("obs_room") or "No Growth",
        "etx_room_wk_of": st.session_state.get("etx_room_weekly") or "N/A",
        "id_room_wk_of": st.session_state.get("id_room_wk_of") or "N/A"
    })

    # 4. Generate DOCX
    if os.path.exists("ScanRDI OOS template.docx"):
        try:
            doc = DocxTemplate("ScanRDI OOS template.docx")
            doc.render(final_data)
            buf = io.BytesIO(); doc.save(buf); buf.seek(0)
            st.download_button("📂 Download DOCX", buf, f"OOS-{clean_filename(st.session_state.oos_id)}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e: st.error(f"DOCX Error: {e}")

    # 5. Generate PDF
    if os.path.exists("ScanRDI OOS template.pdf"):
        try:
            writer = PdfWriter(clone_from="ScanRDI OOS template.pdf")
            pdf_map = {
                'Text Field57': st.session_state.oos_id, 'Date Field0': st.session_state.test_date,
                'Text Field2': st.session_state.sample_id, 'Text Field6': st.session_state.lot_number,
                'Text Field3': smart_personnel_block,
                'Text Field5': st.session_state.dosage_form,
                'Text Field4': st.session_state.sample_name, 
                'Text Field49': smart_phase1_text, 
                'Text Field50': final_data['smart_phase1_continued'], 
                'Text Field26': st.session_state.prepper_name,
                'Text Field27': st.session_state.reader_name, 'Text Field30': st.session_state.scan_id,
                'Text Field32': st.session_state.bsc_id, 
                'Text Field10': final_data['smart_comment_interview'],
                'Text Field11': final_data['smart_comment_samples'], 
                'Text Field12': "Yes, as per SOP 2.600.023",
                'Text Field13': "Yes, as per SOP 2.600.023", 
                'Text Field14': final_data['smart_comment_records']
            }
            for p in writer.pages: writer.update_page_form_field_values(p, pdf_map)
            buf = io.BytesIO(); writer.write(buf); buf.seek(0)
            st.download_button("📂 Download PDF", buf, f"OOS-{clean_filename(st.session_state.oos_id)}.pdf", "application/pdf")
        except Exception as e: st.error(f"PDF Error: {e}")
