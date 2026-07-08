# filename: pages/Celsis.py
import streamlit as st
import os
import re
import json
import io
import sys
import subprocess
import time
from datetime import datetime, timedelta

# --- 1. SAFE UTILS & LOGIC IMPORT ---
try:
    from utils import apply_eagle_style, get_room_logic, get_full_name, get_business_day_back, clean_analyst_name, get_monthly_cleaning_date
    import celsis_logic as cl
except ImportError as e:
    st.error(f"Import Error: {e}")
    def apply_eagle_style(): pass
    def get_room_logic(i): return "Unknown", "000", "", "Unknown"
    def get_full_name(i): return i
    def get_business_day_back(d, n): return d
    def clean_analyst_name(n): return n
    def get_monthly_cleaning_date(d): return ""

# --- 2. PAGE CONFIG & STYLING ---
st.set_page_config(page_title="Celsis Investigation", layout="wide")
apply_eagle_style()

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #f0f2f6; }
    .main { background-color: #ffffff; }
    .stTextArea textarea { background-color: #ffffff; color: #31333F; border: 1px solid #d6d6d6; }
    div[data-testid="stNotification"] { border: 2px solid #ff4b4b; background-color: #ffe8e8; }
    div[data-testid="stAlert"] { border-radius: 5px; }
    </style>
    """, unsafe_allow_html=True)

def ensure_dependencies():
    required = ["docxtpl", "pypdf"]
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
            time.sleep(1); st.rerun()
        except Exception as e: placeholder.error(f"Install failed: {e}")

# --- 3. FILE PERSISTENCE & KEYS ---
STATE_FILE = "celsis_investigation_state.json"
field_keys = cl.FIELD_KEYS if hasattr(cl, 'FIELD_KEYS') else []
if "process_date" not in field_keys: field_keys.append("process_date")

def load_saved_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r") as f: saved_data = json.load(f)
            for key, value in saved_data.items():
                if key in st.session_state: st.session_state[key] = value
        except: pass

def save_current_state():
    try:
        data = {k: v for k, v in st.session_state.items() if k in field_keys}
        with open(STATE_FILE, "w") as f: json.dump(data, f)
    except: pass

def clean_filename(text): 
    return re.sub(r'[\\/*?:"<>|]', '_', str(text)).strip() if text else ""

# --- 4. INIT STATE LOOP ---
def init_state(key, default=""): 
    if key not in st.session_state: st.session_state[key] = default

for k in field_keys:
    if k in ["incidence_count", "total_pos_count_num", "current_pos_order", "em_growth_count", "pos_bottle_count"] or k.startswith("other_order_"): 
        init_state(k, 1)
    elif "etx" in k or "id" in k: init_state(k, "N/A")
    else: init_state(k, "No" if "has" in k or "growth" in k or k == "other_positives" else "")

if "data_loaded" not in st.session_state: load_saved_state(); st.session_state.data_loaded = True
if "report_generated" not in st.session_state: st.session_state.report_generated = False
if "submission_warnings" not in st.session_state: st.session_state.submission_warnings = []

# --- 5. SMART EMAIL PARSER ---
def parse_email_text(text):
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            for k, v in data.items():
                if k in field_keys: st.session_state[k] = v
            st.success("✅ Magic Restore Successful!"); time.sleep(1); st.rerun(); return
    except json.JSONDecodeError: pass

    if m := re.search(r"OOS-(\d+)", text): st.session_state.oos_id = m.group(1)
    if m := re.search(r"^(?:.*\n)?(.*\bE\d{5}\b.*)$", text, re.MULTILINE): 
        st.session_state.client_name = re.sub(r"^Client:\s*", "", m.group(1).strip(), flags=re.IGNORECASE)
    
    if m := re.search(r"\(\s*([A-Z]{2,3})\s*\d+[a-z]{2}\s*Sample\)", text): 
        initial = m.group(1).strip()
        st.session_state.analyst_initial = initial
        st.session_state.analyst_name = get_full_name(initial)

    sample_blocks = re.findall(
        r"(ETX-\d{6}-\d{4})\s*[\r\n]+Sample\s*Name:\s*([^\r\n]+)\s*[\r\n]+(?:Lot|Batch)\s*[:\.]?\s*([^\n\r]+)", 
        text, re.IGNORECASE
    )
    if sample_blocks:
        sample_ids = [b[0].strip() for b in sample_blocks]
        sample_names = [b[1].strip() for b in sample_blocks]
        lot_numbers = [b[2].strip() for b in sample_blocks]
        def join_list(lst):
            if not lst: return ""
            if len(lst) == 1: return lst[0]
            if len(lst) == 2: return f"{lst[0]} and {lst[1]}"
            return ", ".join(lst[:-1]) + " and " + lst[-1]

        st.session_state.sample_id = join_list(sample_ids)
        st.session_state.sample_name = join_list(sample_names)
        st.session_state.lot_number = join_list(lot_numbers)

    if m := re.search(r"aliquoting\s*\(\s*(\d{1,2}\s*[A-Za-z]{3}\s*\d{4})\s*\)", text, re.IGNORECASE):
        try: st.session_state.test_date = datetime.strptime(m.group(1).replace(" ", ""), "%d%b%Y").strftime("%d%b%y")
        except: pass
    if m := re.search(r"processing set up\s*\(\s*(\d{1,2}\s*[A-Za-z]{3}\s*\d{4})\s*\)", text, re.IGNORECASE):
        try: st.session_state.process_date = datetime.strptime(m.group(1).replace(" ", ""), "%d%b%Y").strftime("%d%b%y")
        except: pass

    if m := re.search(r"identification is on-going under\s*(ETX-\d{6}-\d{4})", text, re.IGNORECASE):
        st.session_state.pos_bottle_count = 1
        st.session_state.pos_id_0 = m.group(1).strip()
        st.session_state.pos_org_0 = "Pending"
        if "tsb media" in text.lower() or "in tsb" in text.lower():
            st.session_state.pos_media_0 = "TSB"
        elif "ftm media" in text.lower() or "in ftm" in text.lower():
            st.session_state.pos_media_0 = "FTM"
        else:
            st.session_state.pos_media_0 = "TSB and FTM"
    else:
        microbial_matches = re.findall(r"(ETX-\d{6}-\d{4})\s*\(for", text, re.IGNORECASE)
        if microbial_matches:
            st.session_state.pos_bottle_count = len(microbial_matches)
            for i, mid in enumerate(microbial_matches):
                st.session_state[f"pos_id_{i}"] = mid.strip()
                st.session_state[f"pos_org_{i}"] = "Pending"
                st.session_state[f"pos_media_{i}"] = "TSB and FTM"

    if st.session_state.get("process_date"):
        m_date = get_monthly_cleaning_date(st.session_state.process_date)
        if m_date:
            st.session_state.monthly_cleaning_date = m_date

    save_current_state()

# ================= UI LAYOUT =================
st.title("🧪 Celsis OOS Investigation")

st.header("📧 Smart Email Import / 💾 Restore")
email_input = st.text_area("Paste Celsis Email Content OR Save File here:", height=150)
if st.button("🪄 Parse / Restore"): parse_email_text(email_input); st.success("Updated!"); st.rerun()

st.header("1. General Test Details")
c1, c2, c3, c4 = st.columns(4)
with c1: 
    st.text_input("OOS Number", key="oos_id", help="Required")
    st.text_input("Sample Name", key="sample_name", help="Required")
with c2: 
    st.text_input("Client Name", key="client_name", help="Required")
    st.text_input("Lot Number", key="lot_number", help="Required")
with c3: 
    st.text_input("Sample ID (ETX)", key="sample_id", help="Required")
    current_dosage = st.session_state.get("dosage_form", "")
    if not current_dosage:
        current_dosage = "Injectable"
    options = ["Injectable", "Liquid", "Solution", "Aqueous Solution", "Other"]
    if "dosage_form_select" in st.session_state:
        selected_dosage = st.selectbox("Dosage Form", options, key="dosage_form_select")
    else:
        idx = options.index(current_dosage) if current_dosage in options[:-1] else 4
        selected_dosage = st.selectbox("Dosage Form", options, index=idx, key="dosage_form_select")
    if selected_dosage == "Other":
        prev_custom = st.session_state.get("dosage_form", "")
        default_custom = prev_custom if prev_custom not in options[:-1] else ""
        custom_val = st.text_input("Custom Dosage Form", value=default_custom, key="dosage_form_custom")
        st.session_state.dosage_form = custom_val
    else:
        st.session_state.dosage_form = selected_dosage
with c4: 
    st.text_input("Test Date (Aliquoting)", key="test_date", help="DDMMMYY")
    st.text_input("Process Date (Set up)", key="process_date", help="DDMMMYY")

received_date_str = "[Missing Process Date]"
if st.session_state.get("process_date"):
    try:
        p_dt = datetime.strptime(st.session_state.process_date, "%d%b%y")
        r_dt = get_business_day_back(p_dt, 1)
        received_date_str = r_dt.strftime("%d%b%y")
        st.info(f"📅 **Auto-Calculated Engine:** Received Date (T-1 Business Day): `{received_date_str}`")
    except: pass

st.text_input("Monthly Cleaning Date", key="monthly_cleaning_date", help="Required")

st.header("2. Personnel & Equipment")
p1, p2, p3 = st.columns(3)
with p1: 
    st.text_input("Prepper Initials", key="prepper_initial")
    cl.auto_fill_name("prepper_initial", "prepper_name")
    st.text_input("Prepper Name", key="prepper_name")
with p2: 
    st.text_input("Processor Initials", key="analyst_initial")
    cl.auto_fill_name("analyst_initial", "analyst_name")
    st.text_input("Processor Name", key="analyst_name")
with p3: 
    st.text_input("Aliquoting Initials", key="aliquoting_initial")
    cl.auto_fill_name("aliquoting_initial", "aliquoting_name")
    st.text_input("Aliquoting Name", key="aliquoting_name", help="Acts as Reader too")

e1, e2 = st.columns(2)
bsc_list = ["1310", "1309", "1311", "1312", "1314", "1313", "1316", "1798", "Other"]
with e1:
    current_bsc = st.session_state.get("bsc_id", "")
    idx_bsc = bsc_list.index(current_bsc) if current_bsc in bsc_list[:-1] else len(bsc_list)-1
    selected_bsc = st.selectbox("Processing BSC ID", bsc_list, index=idx_bsc, key="bsc_id_select")
    if selected_bsc == "Other":
        prev_custom_bsc = st.session_state.get("bsc_id", "")
        default_custom_bsc = prev_custom_bsc if prev_custom_bsc not in bsc_list[:-1] else ""
        custom_bsc = st.text_input("Custom BSC ID", value=default_custom_bsc, key="bsc_id_custom")
        st.session_state.bsc_id = custom_bsc
    else:
        st.session_state.bsc_id = selected_bsc

celsis_list = ["2222", "2011", "Other"]
with e2:
    current_celsis = st.session_state.get("celsis_id", "")
    idx_celsis = celsis_list.index(current_celsis) if current_celsis in celsis_list[:-1] else len(celsis_list)-1
    selected_celsis = st.selectbox("Celsis Instrument ID", celsis_list, index=idx_celsis, key="celsis_id_select")
    if selected_celsis == "Other":
        prev_custom_celsis = st.session_state.get("celsis_id", "")
        default_custom_celsis = prev_custom_celsis if prev_custom_celsis not in celsis_list[:-1] else ""
        custom_celsis = st.text_input("Custom Celsis ID", value=default_custom_celsis, key="celsis_id_custom")
        st.session_state.celsis_id = custom_celsis
    else:
        st.session_state.celsis_id = selected_celsis

st.header("3. Celsis Findings")
st.markdown("##### Media & Organism Identifications")
f1, f2, f3 = st.columns(3)
with f1:
    st.number_input("Total Positive Bottles", min_value=1, max_value=10, key="pos_bottle_count")
with f2:
    st.text_input("ATP Control Lot", key="control_lot")
with f3:
    st.text_input("ATP Control Exp", key="control_data")

st.caption("Please specify the details for EACH positive bottle below:")
for i in range(st.session_state.pos_bottle_count):
    col_a, col_b, col_c = st.columns([1, 2, 2])
    with col_a: 
        st.selectbox(f"Bottle #{i+1} Media", ["TSB", "FTM", "TSB and FTM"], key=f"pos_media_{i}")
    with col_b: 
        st.text_input(f"Bottle #{i+1} Microbial ID (ETX)", key=f"pos_id_{i}")
    with col_c: 
        st.text_input(f"Bottle #{i+1} Organism", key=f"pos_org_{i}", help="Pending or actual bug name")

st.header("4. EM Observations")
st.radio("Microbial Growth Observed?", ["No","Yes"], key="em_growth_observed", horizontal=True)

if st.session_state.em_growth_observed == "Yes":
    count = st.number_input("Count of Growth Events", 1, 10, key="em_growth_count")
    for i in range(count):
        st.subheader(f"Growth Event #{i+1}")
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.selectbox(f"Category", ["Personnel Obs", "Surface Obs", "Settling Obs", "Weekly Air Obs", "Weekly Surf Obs"], key=f"em_cat_{i}")
        with col2: st.text_input(f"Obs (e.g. 1 CFU)", key=f"em_obs_{i}")
        with col3: st.text_input(f"ETX ID", key=f"em_etx_{i}")
        with col4: st.text_input(f"Microbial ID", key=f"em_id_{i}")

st.header("5. Investigation Details")
st.subheader("Sample History")
st.radio("Prior failures in last 6 months?", ["No", "Yes"], key="has_prior_failures", horizontal=True)
if st.session_state.has_prior_failures == "Yes":
    count = st.number_input("Number of Prior Failures", 1, 10, key="incidence_count")
    for i in range(count): st.text_input(f"Prior Failure #{i+1} OOS ID", key=f"prior_oos_{i}")

st.subheader("Cross Contamination")
st.radio("Other samples tested positive same day?", ["No", "Yes"], key="other_positives", horizontal=True)
if st.session_state.other_positives == "Yes":
    st.number_input("Total Positive Samples that day", 2, 20, key="total_pos_count_num")
    st.number_input(f"Order of THIS Sample ({st.session_state.sample_id})", 1, 20, key="current_pos_order")
    num_others = st.session_state.total_pos_count_num - 1
    for i in range(num_others):
        col1, col2 = st.columns(2)
        with col1: st.text_input(f"Other Sample #{i+1} ID", key=f"other_id_{i}")
        with col2: st.number_input(f"Other Sample #{i+1} Order", 1, 20, key=f"other_order_{i}")

def create_table_pdf(data):
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8, leading=10, alignment=TA_CENTER)
    header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=8, leading=10, alignment=TA_CENTER, fontName='Helvetica-Bold')
    def p(text, is_header=False): return Paragraph(str(text), header_style if is_header else cell_style)
    elements = []
    
    elements.append(Paragraph(f"Appendix: Supplemental Tables for {data.get('sample_id', '')}", styles['Heading1']))
    elements.append(Spacer(1, 15))
    elements.append(Paragraph(f"Table 1: Information for {data.get('sample_id', '')} under investigation", styles['Heading2']))
    elements.append(Spacer(1, 5))
    
    t1_headers = [p("Processing Analyst", True), p("Aliquoting Analyst", True), p("Sample ID", True), p("Related Microbial ID", True), p("Media with microbial growth", True), p("Microbial ID", True)]
    t1_row = [p(data.get('analyst_name', '')), p(data.get('aliquoting_name', 'N/A')), p(data.get('sample_id', '')), p(data.get('positive_id', '')), p(data.get('positive_media', '')), p(data.get('positive_org', ''))]
    t1 = Table([t1_headers, t1_row], colWidths=[130, 130, 110, 110, 130, 130])
    t1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5)
    ]))
    elements.append(t1)
    elements.append(Spacer(1, 20))
    
    elements.append(Paragraph(f"Table 2: Environmental Monitoring from Processing Performed on {data.get('process_date', '')}", styles['Heading2']))
    elements.append(Spacer(1, 5))
    
    t2_headers = [p(h, True) for h in ["Sampling Site", "Freq", "Date", "Analyst", "Day/Week(s)", "Observation*", "Plate ETX ID", "Microbial ID", "Notes"]]
    
    rows = []
    # Personnel
    rows.append([p("Personnel EM Bracketing", True)] + [""]*8)
    rows.append([p("Personal (Left/Right)"), p("Daily"), p(data.get('before_test', '')), p(data.get('analyst_initial', '')), p("Date Before Testing"), p(data.get('be_obs_pers_dur_pro', '')), p(data.get('be_etx_pers_dur_pro', '')), p(data.get('be_id_pers_dur_pro', '')), p("None")])
    rows.append([p("Personal (Left/Right)"), p("Daily"), p(data.get('test_date', '')), p(data.get('analyst_initial', '')), p("Date of Testing"), p(data.get('obs_pers_dur_pro', '')), p(data.get('etx_pers_dur_pro', '')), p(data.get('id_pers_dur_pro', '')), p("None")])
    rows.append([p("Personal (Left/Right)"), p("Daily"), p(data.get('after_test', '')), p(data.get('analyst_initial', '')), p("Date After Testing"), p(data.get('af_obs_pers_dur_pro', '')), p(data.get('af_etx_pers_dur_pro', '')), p(data.get('af_id_pers_dur_pro', '')), p("None")])
    
    # BSC
    rows.append([p(f"Biological Safety Cabinet EM Bracketing ({data.get('bsc_id', '')})", True)] + [""]*8)
    rows.append([p("Surface Sampling (ISO 5)"), p("Daily"), p(data.get('before_test', '')), p(data.get('analyst_initial', '')), p("Date Before Testing"), p(data.get('be_obs_surf_dur_pro', '')), p(data.get('be_etx_surf_dur_pro', '')), p(data.get('be_id_surf_dur_pro', '')), p("None")])
    rows.append([p("Surface Sampling (ISO 5)"), p("Daily"), p(data.get('test_date', '')), p(data.get('analyst_initial', '')), p("Date of Testing"), p(data.get('obs_surf_dur_pro', '')), p(data.get('etx_surf_dur_pro', '')), p(data.get('id_surf_dur_pro', '')), p("None")])
    rows.append([p("Surface Sampling (ISO 5)"), p("Daily"), p(data.get('after_test', '')), p(data.get('analyst_initial', '')), p("Date After Testing"), p(data.get('af_obs_surf_dur_pro', '')), p(data.get('af_etx_surf_dur_pro', '')), p(data.get('af_id_surf_dur_pro', '')), p("None")])
    
    # Settling
    rows.append([p("Settling Sampling of ISO 5", True)] + [""]*8)
    rows.append([p("Settling Sampling (ISO 5)"), p("Daily"), p(data.get('before_test', '')), p(data.get('analyst_initial', '')), p("Date Before Testing"), p(data.get('be_obs_sett_dur_pro', '')), p(data.get('be_etx_sett_dur_pro', '')), p(data.get('be_id_sett_dur_pro', '')), p("None")])
    rows.append([p("Settling Sampling (ISO 5)"), p("Daily"), p(data.get('test_date', '')), p(data.get('analyst_initial', '')), p("Date of Testing"), p(data.get('obs_sett_dur_pro', '')), p(data.get('etx_sett_dur_pro', '')), p(data.get('id_sett_dur_pro', '')), p("None")])
    rows.append([p("Settling Sampling (ISO 5)"), p("Daily"), p(data.get('after_test', '')), p(data.get('analyst_initial', '')), p("Date After Testing"), p(data.get('af_obs_sett_dur_pro', '')), p(data.get('af_etx_sett_dur_pro', '')), p(data.get('af_id_sett_dur_pro', '')), p("None")])
    
    # Weekly Air
    rows.append([p("Weekly Active Air Sampling Bracketing", True)] + [""]*8)
    rows.append([p("Active Air Sampling"), p("Weekly"), p(data.get('date_of_weekly', '')), p("SMO"), p("Week (Before Testing Date)"), p(data.get('obs_air_wk_of', '')), p(data.get('etx_air_wk_of', '')), p(data.get('id_air_wk_of', '')), p("None")])
    rows.append([p("Active Air Sampling"), p("Weekly"), p(data.get('date_of_weekly', '')), p("SMO"), p("Week (On/After Testing Date)"), p(data.get('obs_air_wk_of', '')), p(data.get('etx_air_wk_of', '')), p(data.get('id_air_wk_of', '')), p("None")])
    
    # Weekly Surface
    rows.append([p("Surface Sampling of Anteroom and Cleanroom Bracketing", True)] + [""]*8)
    rows.append([p("Surface Sampling"), p("Weekly"), p(data.get('date_of_weekly', '')), p("SMO"), p("Week (Before Testing Date)"), p(data.get('obs_room_wk_of', '')), p(data.get('etx_room_wk_of', '')), p(data.get('id_room_wk_of', '')), p("None")])
    rows.append([p("Surface Sampling"), p("Weekly"), p(data.get('date_of_weekly', '')), p("SMO"), p("Week (On/After Testing Date)"), p(data.get('obs_room_wk_of', '')), p(data.get('etx_room_wk_of', '')), p(data.get('id_room_wk_of', '')), p("None")])
    
    t2 = Table([t2_headers] + rows, colWidths=[150, 40, 60, 45, 130, 80, 80, 110, 45])
    t2.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.whitesmoke), ('SPAN', (0, 1), (-1, 1)),
        ('BACKGROUND', (0, 5), (-1, 5), colors.whitesmoke), ('SPAN', (0, 5), (-1, 5)),
        ('BACKGROUND', (0, 9), (-1, 9), colors.whitesmoke), ('SPAN', (0, 9), (-1, 9)),
        ('BACKGROUND', (0, 13), (-1, 13), colors.whitesmoke), ('SPAN', (0, 13), (-1, 13)),
        ('BACKGROUND', (0, 16), (-1, 16), colors.whitesmoke), ('SPAN', (0, 16), (-1, 16)),
    ]))
    elements.append(t2)
    doc.build(elements)
    buffer.seek(0)
    return buffer

save_current_state()
st.divider()

# --- 7. GENERATION & VALIDATION ---
if st.button("🚀 GENERATE CELSIS REPORT", type="primary"):
    ensure_dependencies()
    errors, warnings = cl.validate_inputs()
    if errors:
        for e in errors: st.error(e)
    elif warnings:
        st.session_state.submission_warnings = warnings
        st.rerun() 
    else:
        st.session_state.report_generated = True
        st.session_state.submission_warnings = [] 

if st.session_state.submission_warnings:
    st.warning(f"⚠️ The following fields are empty: {', '.join(st.session_state.submission_warnings)}")
    col_yes, col_no = st.columns([1, 5])
    if col_yes.button("✅ Yes, Proceed Anyway"):
        st.session_state.report_generated = True; st.session_state.submission_warnings = []; st.rerun()
    if col_no.button("❌ No, Let me Fix"):
        st.session_state.submission_warnings = []; st.rerun()

if st.session_state.report_generated:
    with st.spinner("Compiling Celsis logic..."):
        
        pos_media_list = [st.session_state.get(f"pos_media_{i}", "") for i in range(st.session_state.pos_bottle_count)]
        pos_id_list = [st.session_state.get(f"pos_id_{i}", "") for i in range(st.session_state.pos_bottle_count)]
        pos_org_list = [st.session_state.get(f"pos_org_{i}", "") for i in range(st.session_state.pos_bottle_count)]
        
        def join_unique(lst):
            clean_lst = [str(x).strip() for x in lst if str(x).strip() and str(x).strip() != "N/A"]
            if not clean_lst: return "N/A"
            unique_lst = list(dict.fromkeys(clean_lst))
            if len(unique_lst) == 1: return unique_lst[0]
            if len(unique_lst) == 2: return f"{unique_lst[0]} and {unique_lst[1]}"
            return ", ".join(unique_lst[:-1]) + " and " + unique_lst[-1]

        # 智能降维处理 TSB/FTM
        raw_media = [str(x).strip() for x in pos_media_list if str(x).strip() and str(x).strip() != "N/A"]
        if "TSB and FTM" in raw_media or ("TSB" in raw_media and "FTM" in raw_media):
            st.session_state.positive_media = "TSB and FTM"
        elif "TSB" in raw_media:
            st.session_state.positive_media = "TSB"
        elif "FTM" in raw_media:
            st.session_state.positive_media = "FTM"
        else:
            st.session_state.positive_media = "N/A"

        st.session_state.positive_id = join_unique(pos_id_list)
        st.session_state.positive_org = join_unique(pos_org_list)

        fresh_equip = cl.generate_celsis_equipment_text()
        fresh_narr, fresh_det = cl.generate_celsis_narrative_and_details()
        fresh_history = cl.generate_celsis_history_text()
        fresh_cross = cl.generate_celsis_cross_contam_text()
        
        t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
        safe_filename = clean_filename(f"OOS-{st.session_state.oos_id} {st.session_state.client_name} - Celsis")

        # =====================================================================
        # --- 智能双擎单复数探测器 (Smart Dual-Engine Concordance) ---
        # =====================================================================
        is_plural_sample = "and" in str(st.session_state.sample_id).lower() or "," in str(st.session_state.sample_id)
        sample_noun = "samples" if is_plural_sample else "sample"
        sample_verb = "were" if is_plural_sample else "was"
        
        is_plural_bottle = st.session_state.pos_bottle_count > 1 or "and" in str(st.session_state.positive_media).lower() or "and" in str(st.session_state.sample_id).lower()
        bottle_noun = "bottles" if is_plural_bottle else "bottle"
        submit_verb = "were" if is_plural_bottle else "was"
        org_noun = "organisms were" if is_plural_bottle or "and" in str(st.session_state.positive_org).lower() else "organism was"

        # Collect and deduplicate analyst names preserving order
        analysts_raw = [
            st.session_state.get("prepper_name", ""),
            st.session_state.get("analyst_name", ""),
            st.session_state.get("aliquoting_name", ""),
        ]
        analysts_clean = [str(x).strip() for x in analysts_raw if str(x).strip() and str(x).strip() != "N/A"]
        analysts_unique = list(dict.fromkeys(analysts_clean))
        
        if not analysts_unique:
            names_only_phrase = "N/A"
            analysts_with_prefix_phrase = "the analysts"
        elif len(analysts_unique) == 1:
            names_only_phrase = analysts_unique[0]
            analysts_with_prefix_phrase = f"analyst {analysts_unique[0]}"
        elif len(analysts_unique) == 2:
            names_only_phrase = f"{analysts_unique[0]} and {analysts_unique[1]}"
            analysts_with_prefix_phrase = f"analysts {analysts_unique[0]} and {analysts_unique[1]}"
        else:
            names_only_phrase = ", ".join(analysts_unique[:-1]) + ", and " + analysts_unique[-1]
            analysts_with_prefix_phrase = f"analysts " + ", ".join(analysts_unique[:-1]) + ", and " + analysts_unique[-1]

        p1 = f"All analysts involved in the prepping, processing, aliquoting, and reading of the {sample_noun} – {names_only_phrase} were interviewed comprehensively. Their answers are recorded throughout this document."
        p2 = f"Upon arrival, the {sample_noun} {sample_verb} stored in accordance with the Client’s instructions. Analyst {st.session_state.prepper_name} verified the integrity of the {sample_noun} throughout both the preparation and processing stages. No leaks or turbidity were observed at any point, verifying the integrity of the {sample_noun}."
        p3 = "All reagents and supplies mentioned in the material section above were stored according to the suppliers’ recommendations, and their integrity was visually verified before utilization. Moreover, all reagents and supplies had valid expiration dates. The functionality of all equipment was confirmed by reviewing data generated by our comprehensive in-house continuous monitoring system."
        p4 = f"During the preparation phase, {st.session_state.prepper_name} disinfected the {sample_noun} using acidified bleach and placed them into a pre-disinfected storage bin. On {st.session_state.test_date}, prior to sample processing, {st.session_state.analyst_name} performed a second disinfection with acidified bleach, allowing a minimum contact time of 10 minutes before transferring the {sample_noun} into the cleanroom suites. A final disinfection step was completed immediately before the {sample_noun} were introduced into the ISO 5 Biological Safety Cabinet (BSC), E00{st.session_state.bsc_id}, located within the {t_loc}, (Suite {t_suite}{t_suffix}). All activities were conducted in accordance with SOP 2.600.059 for the Celsis sterility testing."
        p5 = fresh_equip
        p6 = f"On {received_date_str}, the sample vials for {st.session_state.sample_id} were received from the Sample Submissions team and brought into the Sterile Microbiology lab. Upon arrival, each sample vial was sprayed with an acidified bleach disinfectant, placed into pre-disinfected bins, and allowed a 10-minute contact time. The secondary disinfection happened in the ISO 8 anteroom (Suite {t_suite}), where the vials were again treated with acidified bleach and provided a 10‑minute contact time before processing. Subsequently, the vials were moved into the ISO 7 cleanroom Suite {t_suite}{t_suffix}. Inside this cleanroom, the processing analyst, {st.session_state.analyst_name}, performed a final disinfection step, allowing an additional 10-minute contact time. Once fully disinfected, the vials were transferred into the ISO 5 BSC E00{st.session_state.bsc_id}."
        p7 = f"Once transferred into the ISO 5 BSC, the vials were placed on the disinfected working surface of the BSC E00{st.session_state.bsc_id} and aseptically opened and tested in accordance with SOP 2.600.059 (Celsis Sterility Testing). Following testing, the media bottles were subsequently transferred into designated incubators, E001356 and E001357, to initiate incubation."
        p8 = f"Upon completion of incubation on {st.session_state.test_date}, both TSB & FTM bottles were disinfected and transferred to the middle ISO 7 buffer room (Suite 114A) for aliquoting step per SOP 2.600.059 (Celsis Sterility Testing). In Suite 114A, the media bottles were disinfected one more time before transferring them to the ISO 5 BSC E001798 located in Suite 114A. In ISO 5 BSC E001798, the {sample_noun} {sample_verb} aliquoted into assay cuvettes by analyst {st.session_state.aliquoting_name}. After aliquoting, Celsis Sterility Reading was performed in accordance with SOP 2.600.059 by analyst {st.session_state.aliquoting_name}."
        p9 = f"Following the reading, {sample_noun} {st.session_state.sample_id} {sample_verb} found to yield a positive reading in the {st.session_state.positive_media} media {bottle_noun}. The average Relative Luminescence Units (RLU) from the duplicate reading tubes, originating from the {st.session_state.positive_media} sample {bottle_noun}, exceeded the average RLU of the {st.session_state.positive_media} negative control, confirming a positive result. The %CV from the duplicate reading tubes for the positive {st.session_state.positive_media} {bottle_noun} were well within the specification (< 30%). Additionally, all Daily Controls, including the Instrument Blank, Reagent Blank, and ATP Positive Control, were within the defined specifications, each with a %CV below 30%."
        p10 = f"Following the OOS result, the positive {st.session_state.positive_media} {bottle_noun} for {st.session_state.sample_id} {submit_verb} submitted for Differential Staining and Microbial Identification under {st.session_state.positive_id}, where the {org_noun} identified as {st.session_state.positive_org}."
        p11 = "The culture media utilized were within their expiry period. The negative culture media bottles for the direct inoculation method for the original culture were handled, processed, and incubated in a manner identical to that of actual samples. No microbial growth was observed in the corresponding negative control."
        p12 = fresh_narr
        if fresh_det: p12 += "\n\n" + fresh_det
        p13 = "The analysts confirmed full compliance with cleaning procedures as outlined in SOPs 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology), 2.600.059 (Celsis Sterility Testing), and 2.600.008 (USP <71> / EP 2.6.1 Sterility Test)."
        p14 = f"Monthly cleaning and disinfection of the outermost ISO 8 Anteroom, the middle ISO 7 Buffer room, the innermost ISO 7 cleanroom, and its containing ISO 5 Biosafety Cabinets for CR {t_suite} and CR 114 was performed on {st.session_state.monthly_cleaning_date}, as per SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology) During both cleaning cycles, it was documented that all H₂O₂ indicators passed. This confirms the efficient monthly cleaning of all three parts of Cleanrooms {t_suite} and 114."
        p15 = fresh_history
        p16 = f"To assess the potential for sample-to-sample contamination contributing to the positive results, a comprehensive review was conducted of all samples processed on the same day. {fresh_cross}"
        p17 = "Based on the observations outlined above, it is unlikely that the failing results were due to reagents, supplies, the cleanroom environment, the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is minimal and the original result is deemed to be valid."

        smart_phase1_full = "\n\n".join([p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17])
        smart_phase1_part1 = "\n\n".join([p1, p2, p3, p4, p5, p6, p7])
        smart_phase1_part2 = "\n\n".join([p8, p9, p10, p11, p12, p13, p14, p15, p16, p17])

        analyst_sig_text = f"{st.session_state.analyst_name} (Written by: Qiyue Chen)"
        personnel_lines = []
        p_name = st.session_state.prepper_name.strip().lower()
        a_name = st.session_state.analyst_name.strip().lower()
        p_init = st.session_state.prepper_initial.strip().lower()
        a_init = st.session_state.analyst_initial.strip().lower()
        is_same = (p_name == a_name) or (p_init and a_init and p_init == a_init)
        
        if not is_same:
            personnel_lines.append(f"Prepper: \n{st.session_state.prepper_name} ({st.session_state.prepper_initial})")
        personnel_lines.extend([
            f"Processor:\n{st.session_state.analyst_name} ({st.session_state.analyst_initial})",
            f"Aliquoting Analyst:\n{st.session_state.aliquoting_name} ({st.session_state.aliquoting_initial})"
        ])
        smart_personnel_block = "\n\n".join(personnel_lines)
                                 
        smart_incident_opening = f"On {st.session_state.test_date}, {sample_noun} {st.session_state.sample_id} {sample_verb} found positive for viable microorganisms after Celsis sterility testing."
        
        try:
            d_obj = datetime.strptime(st.session_state.test_date, "%d%b%y")
            pdf_date_str = d_obj.strftime("%d-%b-%Y")
        except Exception:
            pdf_date_str = st.session_state.test_date

        word_data = {
            "test_date": st.session_state.test_date, "process_date": st.session_state.process_date, "received_data": received_date_str,
            "oos_id": st.session_state.oos_id, "client_name": st.session_state.client_name, "sample_id": st.session_state.sample_id,
            "sample_name": st.session_state.sample_name, "lot_number": st.session_state.lot_number, "dosage_form": st.session_state.dosage_form,
            "prepper_name": st.session_state.prepper_name, "prepper_initial": st.session_state.prepper_initial,
            "analyst_name": st.session_state.analyst_name, "analyst_initial": st.session_state.analyst_initial,
            "aliquoting_name": st.session_state.aliquoting_name, "aliquoting_initial": st.session_state.aliquoting_initial,
            "bsc_id": st.session_state.bsc_id, "smart_bsc_id": f"E00{st.session_state.bsc_id}", "cr_suit": t_suite, "suit": t_suffix, "bsc_location": t_loc,
            "positive_media": st.session_state.positive_media, "positive_id": st.session_state.positive_id, "positive_org": st.session_state.positive_org,
            "monthly_cleaning_date": st.session_state.monthly_cleaning_date,
            "equipment_summary": fresh_equip, "narrative_summary": fresh_narr, "sample_history_paragraph": fresh_history, "cross_contamination_summary": fresh_cross,
            "report_header": f"{st.session_state.sample_id}\n\n{st.session_state.client_name}", "analyst_signature": analyst_sig_text,
            "smart_personnel_block": smart_personnel_block, "smart_incident_opening": smart_incident_opening,
            "smart_comment_interview": f"Yes, {analysts_with_prefix_phrase} were interviewed comprehensively.",
            "smart_comment_samples": f"Yes, {sample_noun} ID: {st.session_state.sample_id}",
            "smart_comment_records": f"Yes, Information is available in EagleTrax under {st.session_state.sample_id}",
            "smart_comment_storage": f"Yes, the {sample_noun} {sample_verb} stored as per client's instructions. Information is available in EagleTrax Sample Location History under {st.session_state.sample_id}",
            "control_positive": "Celsis ATP Positive Control", "control_lot": st.session_state.control_lot, "control_data": st.session_state.control_data,
            "smart_scan_id": f"E00{st.session_state.celsis_id}", "smart_cr_id": f"E00{t_room} (CR{t_suite})" if t_suite == "114" else f"For Processing/Reading: E00{t_room} (CR{t_suite})\nFor Aliquoting: E001736 (CR114)",
            "smart_phase1_summary": smart_phase1_full, "smart_phase1_continued": ""
        }

        pdf_map = {
            'Text Field57': st.session_state.oos_id, 'Date Field0': pdf_date_str, 'Date Field1': pdf_date_str, 
            'Date Field2': pdf_date_str, 'Date Field3': pdf_date_str,
            'Text Field2': st.session_state.sample_id, 'Text Field6': st.session_state.lot_number, 
            'Text Field4': st.session_state.sample_name + "\n\n\n\n", 'Text Field5': st.session_state.dosage_form, 
            'Text Field0': analyst_sig_text, 'Text Field3': smart_personnel_block, 'Text Field7': smart_incident_opening + "\n\n",
            'Text Field13': word_data["smart_comment_interview"], 'Text Field14': word_data["smart_comment_samples"], 
            'Text Field17': word_data["smart_comment_records"], 'Text Field21': word_data["smart_comment_storage"],
            'Text Field30': f"E00{st.session_state.celsis_id}", 'Text Field32': word_data["smart_cr_id"], 
            'Text Field34': f"E00{st.session_state.celsis_id}", 'Text Field25': st.session_state.control_lot, 
            'Text Field26': st.session_state.control_data, 'Text Field49': smart_phase1_part1, 'Text Field50': smart_phase1_part2
        }

        docx_buf, pdf_form_buf = None, None
        tables_docx_buf, tables_pdf_buf = None, None

        # --- CALCULATE TABLE SPECIFIC DATA ---
        table_data = word_data.copy()
        process_date_str = st.session_state.get("process_date", "")
        before_test_val = ""
        after_test_val = ""
        if process_date_str:
            try:
                fmt = "%d%b%y" if len(process_date_str) <= 7 else "%d%b%Y"
                p_dt = datetime.strptime(process_date_str, fmt)
                before_dt = p_dt - timedelta(days=1)
                after_dt = p_dt + timedelta(days=1)
                before_test_val = before_dt.strftime("%d%b%y")
                after_test_val = after_dt.strftime("%d%b%y")
            except: pass
            
        table_data["before_test"] = before_test_val
        table_data["after_test"] = after_test_val
        table_data["test_date"] = process_date_str
        
        table_data["obs_pers_dur_pro"] = st.session_state.get("obs_pers", "No Growth")
        table_data["etx_pers_dur_pro"] = st.session_state.get("etx_pers", "N/A")
        table_data["id_pers_dur_pro"] = st.session_state.get("id_pers", "N/A")
        
        table_data["obs_surf_dur_pro"] = st.session_state.get("obs_surf", "No Growth")
        table_data["etx_surf_dur_pro"] = st.session_state.get("etx_surf", "N/A")
        table_data["id_surf_dur_pro"] = st.session_state.get("id_surf", "N/A")
        
        table_data["obs_sett_dur_pro"] = st.session_state.get("obs_sett", "No Growth")
        table_data["etx_sett_dur_pro"] = st.session_state.get("etx_sett", "N/A")
        table_data["id_sett_dur_pro"] = st.session_state.get("id_sett", "N/A")
        
        for prefix in ["be_", "af_"]:
            for suffix in ["pers_dur_pro", "surf_dur_pro", "sett_dur_pro"]:
                table_data[f"{prefix}obs_{suffix}"] = "No Growth"
                table_data[f"{prefix}etx_{suffix}"] = "N/A"
                table_data[f"{prefix}id_{suffix}"] = "N/A"
                
        table_data["obs_air_wk_of"] = st.session_state.get("obs_air", "No Growth")
        table_data["etx_air_wk_of"] = st.session_state.get("etx_air_weekly", "N/A")
        table_data["id_air_wk_of"] = st.session_state.get("id_air_weekly", "N/A")
        
        table_data["obs_room_wk_of"] = st.session_state.get("obs_room", "No Growth")
        table_data["etx_room_wk_of"] = st.session_state.get("etx_room_weekly", "N/A")
        table_data["id_room_wk_of"] = st.session_state.get("id_room_wk_of", "N/A")
        
        table_data["positive_id"] = st.session_state.get("positive_id", "N/A")
        table_data["positive_media"] = st.session_state.get("positive_media", "N/A")
        table_data["positive_org"] = st.session_state.get("positive_org", "N/A")

        # --- 1. RENDER MAIN DOCX ---
        target_template = "Celsis OOS P1 template 0.docx"
        if not os.path.exists(target_template): target_template = "Celsis OOS P1 template.docx"
        if os.path.exists(target_template):
            try:
                from docxtpl import DocxTemplate
                doc = DocxTemplate(target_template)
                doc.render(word_data); docx_buf = io.BytesIO(); doc.save(docx_buf); docx_buf.seek(0)
            except Exception as e: st.error(f"DOCX Error: {e}")
        else: st.warning("⚠️ Could not find either 'Celsis OOS P1 template 0.docx' or 'Celsis OOS P1 template.docx'.")

        # --- 2. RENDER TABLES DOCX ---
        target_tables_template = "tables for celsis.docx"
        if os.path.exists(target_tables_template):
            try:
                from docxtpl import DocxTemplate
                doc_tbl = DocxTemplate(target_tables_template)
                doc_tbl.render(table_data); tables_docx_buf = io.BytesIO(); doc_tbl.save(tables_docx_buf); tables_docx_buf.seek(0)
            except Exception as e: st.error(f"Tables DOCX Error: {e}")
        else: st.warning(f"⚠️ Could not find {target_tables_template}.")
            
        # --- 3. RENDER MAIN PDF ---
        if os.path.exists("Celsis OOS P1 template.pdf"):
            try:
                from pypdf import PdfWriter
                writer = PdfWriter(clone_from="Celsis OOS P1 template.pdf") 
                for p in writer.pages: writer.update_page_form_field_values(p, pdf_map)
                pdf_form_buf = io.BytesIO(); writer.write(pdf_form_buf); pdf_form_buf.seek(0)
            except Exception as e: st.error(f"PDF Form Error: {e}")

        # --- 4. RENDER TABLES PDF ---
        try:
            tables_pdf_buf = create_table_pdf(table_data)
        except Exception as e: st.error(f"Tables PDF Error: {e}")

        st.success("✅ Celsis Reports and Tables Generated!")
        st.markdown("### 📂 Download Reports")
        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            if docx_buf: st.download_button("📄 Celsis Report (doc)", docx_buf, f"{safe_filename}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            if tables_docx_buf: st.download_button("📄 Tables (doc)", tables_docx_buf, f"Tables {safe_filename}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with c_dl2:
            if pdf_form_buf: st.download_button("🔴 Celsis Report (pdf)", pdf_form_buf, f"{safe_filename}.pdf", "application/pdf")
            if tables_pdf_buf: st.download_button("🔴 Tables (pdf)", tables_pdf_buf, f"Tables {safe_filename}.pdf", "application/pdf")

        st.markdown("---")
        current_data = {k: st.session_state[k] for k in field_keys if k in st.session_state}
        st.download_button("💾 Save Session Data (.txt)", json.dumps(current_data, indent=2), f"SAVE_{safe_filename}.txt", "text/plain")
