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
    from utils import apply_eagle_style, get_room_logic, get_full_name, get_business_day_back
    import usp71_logic as ul
except ImportError as e:
    st.error(f"Import Error: {e}")
    def apply_eagle_style(): pass
    def get_room_logic(i): return "Unknown", "000", "", "Unknown"
    def get_full_name(i): return i
    def get_business_day_back(d, n): return d

# --- 2. PAGE CONFIG & STYLING ---
st.set_page_config(page_title="USP 71 Investigation", layout="wide")
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
STATE_FILE = "usp71_investigation_state.json"
field_keys = ul.FIELD_KEYS if hasattr(ul, 'FIELD_KEYS') else []

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


# --- HELPER: USERNAME TO INITIALS ---
def username_to_initials(username):
    if not username:
        return ""
    username = username.strip()
    known = {
        "gsurber": "GS", "GSurber": "GS",
        "enioupin": "EN", "ENioupin": "EN",
        "acarrillo": "AC", "ACarrillo": "AC",
        "rseymour": "RS", "RSeymour": "RS",
        "jowens": "JO", "JOwens": "JO"
    }
    if username in known:
        return known[username]
    if username.lower() in known:
        return known[username.lower()]
        
    uppers = "".join([c for c in username if c.isupper()])
    if len(uppers) >= 2:
        return uppers[:3]
    return username[0].upper() + (username[1].upper() if len(username) > 1 else "")


# --- STATE FILE MERGING HELPERS ---
def load_state_from_file():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r") as f:
                return json.load(f)
        except: pass
    return {}

def save_state_to_file(data):
    try:
        with open(STATE_FILE, "w") as f:
            json.dump(data, f)
    except: pass

# --- 5. SMART COMBINED PARSER ---
def parse_combined_text(text):
    # 1. TRY JSON LOAD (RESTORE FUNCTION)
    try:
        data = json.loads(text)
        if isinstance(data, dict):
            for k, v in data.items():
                if k in field_keys: st.session_state[k] = v
            save_current_state()
            st.success("✅ Magic Restore Successful!"); time.sleep(1); st.rerun(); return
    except json.JSONDecodeError: pass

    # Load existing persisted state to prevent losing fields
    persisted = load_state_from_file()
    parsed = {}

    is_email = any(x in text.lower() for x in ["oos-", "day of testing", "day of incubation", "positive for", "results have shown"])
    is_event = any(x in text.lower() for x in ["status changed", "sample prep", "sterility read", "incubation started", "sample positive"]) or "\t" in text

    if is_email:
        # Email parsing
        if m := re.search(r"OOS-(\d+)", text): parsed["oos_id"] = m.group(1)
        if m := re.search(r"^(?:.*\n)?(.*\bE\d{5}\b.*)$", text, re.MULTILINE): 
            parsed["client_name"] = re.sub(r"^Client:\s*", "", m.group(1).strip(), flags=re.IGNORECASE)
        
        if m := re.search(r"\(\s*([A-Z]{2,3})\s*\d+[a-z]{2}\s*Sample\)", text): 
            initial = m.group(1).strip()
            parsed["analyst_initial"] = initial
            parsed["analyst_name"] = get_full_name(initial)

        if m := re.search(r"(ETX-\d{6}-\d{4})", text): parsed["sample_id"] = m.group(1).strip()
        if m := re.search(r"Sample\s*Name:\s*(.*)", text, re.I): parsed["sample_name"] = m.group(1).strip()
        if m := re.search(r"(?:Lot|Batch)\s*[:\.]?\s*([^\n\r]+)", text, re.I): parsed["lot_number"] = m.group(1).strip()

        # Dates
        if m := re.search(r"day of testing\s*\(\s*([^)]+)\)", text, re.IGNORECASE):
            date_str = m.group(1).strip()
            try: parsed["process_date"] = datetime.strptime(date_str, "%d %b %Y").strftime("%d%b%y")
            except: pass
        elif m := re.search(r"testing\s*on\s*(\d{1,2}\s*[A-Za-z]{3}\s*\d{4})", text, re.IGNORECASE):
            try: parsed["process_date"] = datetime.strptime(m.group(1).strip(), "%d %b %Y").strftime("%d%b%y")
            except: pass
        elif m := re.search(r"reading\s*\(\s*(\d{1,2}\s*[A-Za-z]{3}\s*\d{4})\s*\)", text, re.IGNORECASE):
            try: parsed["process_date"] = datetime.strptime(m.group(1).replace(" ", ""), "%d%b%Y").strftime("%d%b%y")
            except: pass

        if m := re.search(r"as of\s*(\d{1,2}\s*[A-Za-z]{3}\s*\d{4})", text, re.IGNORECASE):
            date_str = m.group(1).strip()
            try: parsed["test_date"] = datetime.strptime(date_str, "%d %b %Y").strftime("%d%b%y")
            except: pass

        inc_days_email = None
        if m := re.search(r"on Day\s*(\d+)\s*of incubation", text, re.IGNORECASE):
            try: 
                inc_days_email = int(m.group(1))
                parsed["incubation_time"] = str(inc_days_email)
            except: pass
        
        process_date_val = parsed.get("process_date") or persisted.get("process_date") or st.session_state.get("process_date")
        test_date_val = parsed.get("test_date") or persisted.get("test_date") or st.session_state.get("test_date")
        
        if inc_days_email is not None:
            if test_date_val and not process_date_val:
                try:
                    t_dt = datetime.strptime(test_date_val, "%d%b%y")
                    p_dt = t_dt - timedelta(days=inc_days_email)
                    parsed["process_date"] = p_dt.strftime("%d%b%y")
                except: pass
            elif process_date_val and not test_date_val:
                try:
                    p_dt = datetime.strptime(process_date_val, "%d%b%y")
                    t_dt = p_dt + timedelta(days=inc_days_email)
                    parsed["test_date"] = t_dt.strftime("%d%b%y")
                except: pass

        if m := re.search(r"identification is on-going under\s*(ETX-\d{6}-\d{4})", text, re.IGNORECASE):
            parsed["pos_bottle_count"] = 1
            parsed["pos_id_0"] = m.group(1).strip()
            parsed["pos_org_0"] = "Pending"
            if "tsb media" in text.lower() or "in tsb" in text.lower():
                parsed["pos_media_0"] = "TSB"
            elif "ftm media" in text.lower() or "in ftm" in text.lower():
                parsed["pos_media_0"] = "FTM"
            else:
                parsed["pos_media_0"] = "TSB and FTM"
        else:
            microbial_matches = re.findall(r"(ETX-\d{6}-\d{4})\s*\(for", text, re.IGNORECASE)
            if microbial_matches:
                parsed["pos_bottle_count"] = len(microbial_matches)
                for i, mid in enumerate(microbial_matches):
                    parsed[f"pos_id_{i}"] = mid.strip()
                    parsed[f"pos_org_{i}"] = "Pending"
                    parsed[f"pos_media_{i}"] = "TSB and FTM"

        if m := re.search(r"results have shown\s+([^\s]+(?: \(\+\)| \(\-\))? [^\s]+)", text, re.I):
            parsed["organism_morphology"] = m.group(1).strip()
        elif m := re.search(r"(\w+)-shaped", text, re.I):
            parsed["organism_morphology"] = m.group(1).strip()

        # Only clear subculture if not parsing event history in the same run
        if not is_event:
            parsed["subculture_initial"] = ""
            parsed["subculture_name"] = ""

    if is_event:
        # Event history parsing
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        
        prepper_user = None
        for line in lines:
            if "sample prep" in line.lower():
                parts = line.split("\t")
                if len(parts) >= 3:
                    prepper_user = parts[2].strip()
                    break
        if prepper_user:
            initial = username_to_initials(prepper_user)
            parsed["prepper_initial"] = initial
            parsed["prepper_name"] = get_full_name(initial)
        else:
            parsed["prepper_initial"] = ""
            parsed["prepper_name"] = ""

        processor_user = None
        for line in lines:
            if "status changed" in line.lower() and "sample analysis" in line.lower():
                parts = line.split("\t")
                if len(parts) >= 3:
                    processor_user = parts[2].strip()
                    break
        if not processor_user and prepper_user:
            processor_user = prepper_user

        if processor_user:
            initial = username_to_initials(processor_user)
            parsed["analyst_initial"] = initial
            parsed["analyst_name"] = get_full_name(initial)

        reader_user = None
        inc_days_event = None
        for line in lines:
            if "sterility read" in line.lower() and "positive" in line.lower():
                parts = line.split("\t")
                if len(parts) >= 3:
                    reader_user = parts[2].strip()
                    day_match = re.search(r"Day:\s*(\d+)", parts[0], re.I)
                    if day_match:
                        inc_days_event = int(day_match.group(1))
                        parsed["incubation_time"] = str(inc_days_event)
                    break
        
        if reader_user:
            initial = username_to_initials(reader_user)
            parsed["reading_initial"] = initial
            parsed["reading_name"] = get_full_name(initial)

        subculture_user = None
        has_inconclusive = "inconclusive" in text.lower()
        if has_inconclusive:
            for line in lines:
                if "sterility read" in line.lower() and "inconclusive" in line.lower():
                    parts = line.split("\t")
                    if len(parts) >= 3:
                        subculture_user = parts[2].strip()
                        break
            if not subculture_user:
                for line in lines:
                    if "inconclusive" in line.lower():
                        parts = line.split("\t")
                        if len(parts) >= 3:
                            subculture_user = parts[2].strip()
                            break
        
        if has_inconclusive and subculture_user:
            initial = username_to_initials(subculture_user)
            parsed["subculture_initial"] = initial
            parsed["subculture_name"] = get_full_name(initial)
        else:
            parsed["subculture_initial"] = ""
            parsed["subculture_name"] = ""

        read_date_str = None
        for line in lines:
            if "sterility read" in line.lower() and "positive" in line.lower():
                parts = line.split("\t")
                if len(parts) >= 2:
                    date_part = parts[1].split()[0]
                    try:
                        dt = datetime.strptime(date_part, "%m/%d/%Y")
                        read_date_str = dt.strftime("%d%b%y")
                    except:
                        try:
                            dt = datetime.strptime(date_part, "%Y-%m-%d")
                            read_date_str = dt.strftime("%d%b%y")
                        except: pass
                    break
        
        if read_date_str:
            parsed["test_date"] = read_date_str
            if inc_days_event is not None:
                try:
                    t_dt = datetime.strptime(read_date_str, "%d%b%y")
                    p_dt = t_dt - timedelta(days=inc_days_event)
                    parsed["process_date"] = p_dt.strftime("%d%b%y")
                except: pass
        else:
            inoc_date_str = None
            for line in lines:
                if "incubation started" in line.lower():
                    parts = line.split("\t")
                    if len(parts) >= 2:
                        date_part = parts[1].split()[0]
                        try:
                            dt = datetime.strptime(date_part, "%m/%d/%Y")
                            inoc_date_str = dt.strftime("%d%b%y")
                        except:
                            try:
                                dt = datetime.strptime(date_part, "%Y-%m-%d")
                                inoc_date_str = dt.strftime("%d%b%y")
                            except: pass
                        break
            if inoc_date_str:
                parsed["process_date"] = inoc_date_str
                if inc_days_event is not None:
                    try:
                        p_dt = datetime.strptime(inoc_date_str, "%d%b%y")
                        t_dt = p_dt + timedelta(days=inc_days_event)
                        parsed["test_date"] = t_dt.strftime("%d%b%y")
                    except: pass

        for line in lines:
            if "status changed" in line.lower() and "sample positive" in line.lower():
                parts = line.split("\t")[0]
                m_id = re.search(r"Sample Positive:\s*(ETX-\d{6}-\d{4})", parts, re.I)
                if m_id:
                    parsed["sample_id"] = m_id.group(1).strip()
                    break

        for line in lines:
            if "incubation started" in line.lower():
                parts = line.split("\t")[0]
                m_media = re.search(r"Media:\s*(\w+)", parts, re.I)
                if m_media:
                    parsed["positive_media"] = m_media.group(1).strip()
                    break

    if is_email or is_event:
        persisted.update(parsed)
        save_state_to_file(persisted)
        for k, v in persisted.items():
            if k in field_keys: st.session_state[k] = v

# ================= UI LAYOUT =================
st.title("🧪 USP <71> OOS Investigation")

st.header("📥 Import & Restore Center")
combined_input = st.text_area(
    "Paste USP <71> Email Content, Event History, OR Save File Content here:", 
    height=180, 
    key="combined_import_text"
)
if st.button("🪄 Smart Parse / Restore", key="combined_parse_btn"):
    if combined_input.strip():
        parse_combined_text(combined_input)
        st.success("✅ Content Parsed & Loaded!")
        time.sleep(1)
        st.rerun()

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

st.divider()
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
    st.text_input("Inoculation Date", key="process_date", help="DDMMMYY")
    st.text_input("Reading Date (Incident Date)", key="test_date", help="DDMMMYY")

received_date_str = "[Missing Inoculation Date]"
if st.session_state.get("process_date"):
    try:
        p_dt = datetime.strptime(st.session_state.process_date, "%d%b%y")
        r_dt = get_business_day_back(p_dt, 1)
        received_date_str = r_dt.strftime("%d%b%y")
        st.info(f"📅 **Auto-Calculated Engine:** Received Date (T-1 Business Day of Inoculation Date): `{received_date_str}`")
    except: pass

st.text_input("Monthly Cleaning Date", key="monthly_cleaning_date", help="Required")

st.header("2. Personnel & Equipment")
p1, p2, p3, p4 = st.columns(4)
with p1: 
    st.text_input("Prepper Initials", key="prepper_initial")
    ul.auto_fill_name("prepper_initial", "prepper_name")
    st.text_input("Prepper Name", key="prepper_name")
with p2: 
    st.text_input("Processor Initials", key="analyst_initial")
    ul.auto_fill_name("analyst_initial", "analyst_name")
    st.text_input("Processor Name", key="analyst_name")
with p3: 
    st.text_input("Reader Initials", key="reading_initial")
    ul.auto_fill_name("reading_initial", "reading_name")
    st.text_input("Reader Name", key="reading_name")
with p4: 
    st.text_input("Subculture Initials", key="subculture_initial")
    ul.auto_fill_name("subculture_initial", "subculture_name")
    st.text_input("Subculture Name", key="subculture_name")

bsc_list = ["1310", "1309", "1311", "1312", "1314", "1313", "1316", "1798", "Other"]
st.selectbox("Processing BSC ID", bsc_list, key="bsc_id")

st.header("3. USP 71 Findings")
st.markdown("##### Media & Organism Identifications")
f1, f2, f3 = st.columns(3)
with f1:
    st.number_input("Total Positive Bottles", min_value=1, max_value=10, key="pos_bottle_count")
with f2:
    st.text_input("Incubation Time (Days)", key="incubation_time", value="14")
with f3:
    st.text_input("Organism Morphology", key="organism_morphology", help="e.g. Rod, Cocci")

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
ev1, ev2 = st.columns(2)
with ev1: st.text_input("Event / Deviation Number", key="event_number")
with ev2: st.text_input("Confirmation / CAPA Number", key="confirm_number")

st.subheader("Sample History")
st.radio("Prior failures in last 6 months?", ["No", "Yes"], key="has_prior_failures", horizontal=True)
if st.session_state.has_prior_failures == "Yes":
    count = st.number_input("Number of Prior Failures", 1, 10, key="incidence_count")
    for i in range(count):
        h1, h2, h3, h4 = st.columns(4)
        with h1: st.text_input(f"Prior #{i+1} OOS ID", key=f"prior_oos_{i}")
        with h2: st.text_input(f"Prior #{i+1} Sample ID", key=f"oos1_sample_id_{i}" if i > 0 else "oos1_sample_id")
        with h3: st.text_input(f"Prior #{i+1} Sample Name", key=f"oos1_sample_name_{i}" if i > 0 else "oos1_sample_name")
        with h4: st.text_input(f"Prior #{i+1} Analyst", key=f"oos1_analyst_name_{i}" if i > 0 else "oos1_analyst_name")

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

save_current_state()
st.divider()

# --- 7. GENERATION & VALIDATION ---
if st.button("🚀 GENERATE USP 71 REPORT", type="primary"):
    ensure_dependencies()
    errors, warnings = ul.validate_inputs()
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
    with st.spinner("Compiling USP 71 bulk insertion logic..."):
        
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

        fresh_equip = ul.generate_usp71_equipment_text()
        fresh_narr, fresh_det = ul.generate_usp71_narrative_and_details()
        fresh_history = ul.generate_usp71_history_text()
        fresh_cross = ul.generate_usp71_cross_contam_text()
        
        t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
        safe_filename = clean_filename(f"OOS-{st.session_state.oos_id} {st.session_state.client_name} - USP71")

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

        # BULK INSERTION MACRO-ASSEMBLY
        has_prepper = st.session_state.get("prepper_name", "").strip() and st.session_state.get("prepper_name", "").strip() != "N/A"
        has_subculture = st.session_state.get("subculture_name", "").strip() and st.session_state.get("subculture_name", "").strip() != "N/A"
        
        # Collect and deduplicate analyst names preserving order
        analysts_raw = []
        if has_prepper:
            analysts_raw.append(st.session_state.get("prepper_name", ""))
        analysts_raw.extend([
            st.session_state.get("analyst_name", ""),
            st.session_state.get("reading_name", ""),
        ])
        if has_subculture:
            analysts_raw.append(st.session_state.get("subculture_name", ""))
            
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

        # Dynamic role description based on prepper and subculture existence
        roles = []
        if has_prepper:
            roles.append("prepping")
        roles.append("processing")
        roles.append("reading")
        if has_subculture:
            roles.append("subculturing")
        
        if len(roles) == 1:
            roles_phrase = roles[0]
        elif len(roles) == 2:
            roles_phrase = f"{roles[0]} and {roles[1]}"
        else:
            roles_phrase = ", ".join(roles[:-1]) + ", and " + roles[-1]

        p1 = f"All analysts involved in the {roles_phrase} of the {sample_noun} – {names_only_phrase} were interviewed comprehensively. Their answers are recorded throughout this document."
        
        if has_subculture:
            p8 = f"During the periodic visual inspections by analyst {st.session_state.reading_name}, microbial growth was observed. As a result, a subculture was initiated by analyst {st.session_state.subculture_name} to confirm the presence of viable microorganisms and to proceed with identification."
        else:
            p8 = f"During the periodic visual inspections by analyst {st.session_state.reading_name}, microbial growth was observed, confirming the presence of viable microorganisms and proceeding with identification."
        
        interview_comment = f"Yes, {analysts_with_prefix_phrase} were interviewed comprehensively."

        prep_analyst = st.session_state.prepper_name if has_prepper else st.session_state.analyst_name
        p2 = f"Upon arrival, the {sample_noun} {sample_verb} stored in accordance with the Client’s instructions. Analyst {prep_analyst} verified the integrity of the {sample_noun} throughout both the preparation and processing stages. No leaks or turbidity were observed at any point, verifying the integrity of the {sample_noun}."
        p3 = "All reagents and supplies mentioned in the material section above were stored according to the suppliers’ recommendations, and their integrity was visually verified before utilization. Moreover, all reagents and supplies had valid expiration dates. The functionality of all equipment was confirmed by reviewing data generated by our comprehensive in-house continuous monitoring system."
        prepper_part = f"Analyst {st.session_state.analyst_name}" if not has_prepper else st.session_state.prepper_name
        p4 = f"During the preparation phase, {prepper_part} disinfected the {sample_noun} using acidified bleach and placed them into a pre-disinfected storage bin. On {st.session_state.process_date}, prior to sample processing, {st.session_state.analyst_name} performed a second disinfection with acidified bleach, allowing a minimum contact time of 10 minutes before transferring the {sample_noun} into the cleanroom suites. A final disinfection step was completed immediately before the {sample_noun} were introduced into the ISO 5 Biological Safety Cabinet (BSC), E00{st.session_state.bsc_id}, located within the {t_loc}, (Suite {t_suite}{t_suffix}). All activities were conducted in accordance with SOP 2.600.008 for the USP <71> / EP 2.6.1 sterility testing."
        p5 = fresh_equip
        p6 = f"On {received_date_str}, the sample vials for {st.session_state.sample_id} were received from the Sample Submissions team and brought into the Sterile Microbiology lab. Upon arrival, each sample vial was sprayed with an acidified bleach disinfectant, placed into pre-disinfected bins, and allowed a 10-minute contact time. The secondary disinfection happened in the ISO 8 anteroom (Suite {t_suite}), where the vials were again treated with acidified bleach and provided a 10‑minute contact time before processing. Subsequently, the vials were moved into the ISO 7 cleanroom Suite {t_suite}{t_suffix}. Inside this cleanroom, the processing analyst, {st.session_state.analyst_name}, performed a final disinfection step, allowing an additional 10-minute contact time. Once fully disinfected, the vials were transferred into the ISO 5 BSC E00{st.session_state.bsc_id}."
        p7 = f"Once transferred into the ISO 5 BSC, the vials were placed on the disinfected working surface of the BSC E00{st.session_state.bsc_id} and aseptically processed in accordance with SOP 2.600.008 (USP <71> / EP 2.6.1 Sterility Test). The {sample_noun} {sample_verb} inoculated into Fluid Thioglycollate Medium (FTM) and Tryptic Soy Broth (TSB). Following inoculation, the media bottles were transferred into designated incubators to initiate a 14-day continuous incubation cycle. FTM bottles were incubated at 30-35°C, while TSB bottles were incubated at 20-25°C."
        p9 = f"Following the {st.session_state.incubation_time}-day incubation and visual readings, {sample_noun} {st.session_state.sample_id} {sample_verb} found to yield a positive reading in the {st.session_state.positive_media} media {bottle_noun}, confirming a positive result for microbial growth."
        p10 = f"Following the OOS result, the positive {st.session_state.positive_media} {bottle_noun} for {st.session_state.sample_id} {submit_verb} submitted for Differential Staining and Microbial Identification under {st.session_state.positive_id}, where the {org_noun} identified as {st.session_state.positive_org}."
        p11 = "The culture media utilized were within their expiry period. The negative culture media bottles for the direct inoculation method for the original culture were handled, processed, and incubated in a manner identical to that of actual samples. No microbial growth was observed in the corresponding negative control."
        p12 = fresh_narr
        if fresh_det: p12 += "\n\n" + fresh_det
        p13 = "The analysts confirmed full compliance with cleaning procedures as outlined in SOPs 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology) and 2.600.008 (USP <71> / EP 2.6.1 Sterility Test)."
        p14 = f"Monthly cleaning and disinfection of the outermost ISO 8 Anteroom, the middle ISO 7 Buffer room, the innermost ISO 7 cleanroom, and its containing ISO 5 Biosafety Cabinets for CR {t_suite} was performed on {st.session_state.monthly_cleaning_date}, as per SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). During both cleaning cycles, it was documented that all H₂O₂ indicators passed. This confirms the efficient monthly cleaning of all three parts of Cleanroom {t_suite}."
        p15 = fresh_history
        p16 = f"To assess the potential for sample-to-sample contamination contributing to the positive results, a comprehensive review was conducted of all samples processed on the same day. {fresh_cross}"
        p17 = "Based on the observations outlined above, it is unlikely that the failing results were due to reagents, supplies, the cleanroom environment, the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is minimal and the original result is deemed to be valid."

        smart_phase1_full = "\n\n".join([p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17])
        smart_phase1_part1 = "\n\n".join([p1, p2, p3, p4, p5, p6, p7])
        smart_phase1_part2 = "\n\n".join([p8, p9, p10, p11, p12, p13, p14, p15, p16, p17])

        analyst_sig_text = f"{st.session_state.analyst_name} (Written by: Qiyue Chen)"
        
        personnel_lines = []
        if has_prepper:
            personnel_lines.append(f"Prepper: \n{st.session_state.prepper_name} ({st.session_state.prepper_initial})")
        personnel_lines.extend([
            f"Processor:\n{st.session_state.analyst_name} ({st.session_state.analyst_initial})",
            f"Reading Analyst:\n{st.session_state.reading_name} ({st.session_state.reading_initial})"
        ])
        if has_subculture:
            personnel_lines.append(f"Subculture Analyst:\n{st.session_state.subculture_name} ({st.session_state.subculture_initial})")
        smart_personnel_block = "\n\n".join(personnel_lines)
                                 
        smart_incident_opening = f"On {st.session_state.test_date}, {sample_noun} {st.session_state.sample_id} {sample_verb} found positive for viable microorganisms after USP <71> / EP 2.6.1 sterility testing."
        
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
            "reading_name": st.session_state.reading_name, "reading_initial": st.session_state.reading_initial,
            "subculture_name": st.session_state.subculture_name, "subculture_initial": st.session_state.subculture_initial,
            "bsc_id": st.session_state.bsc_id, "usp71_bsc": st.session_state.bsc_id,
            "cr_suit": t_suite, "suit": t_suffix, "bsc_location": t_loc,
            "positive_media": st.session_state.positive_media, "positive_id": st.session_state.positive_id, "positive_org": st.session_state.positive_org,
            "monthly_cleaning_date": st.session_state.monthly_cleaning_date,
            "equipment_summary": fresh_equip, "narrative_summary": fresh_narr, "sample_history_paragraph": fresh_history, "cross_contamination_summary": fresh_cross,
            "report_header": f"{st.session_state.sample_id}\n\n{st.session_state.client_name}", "analyst_signature": analyst_sig_text,
            "smart_personnel_block": smart_personnel_block, "smart_incident_opening": smart_incident_opening,
            "smart_comment_interview": interview_comment,
            "smart_comment_samples": f"Yes, {sample_noun} ID: {st.session_state.sample_id}",
            "smart_comment_records": f"Yes, Information is available in EagleTrax under {st.session_state.sample_id}",
            "smart_comment_storage": f"Yes, the {sample_noun} {sample_verb} stored as per client's instructions. Information is available in EagleTrax Sample Location History under {st.session_state.sample_id}",
            "smart_scan_id": st.session_state.usp71_id, "smart_cr_id": f"For Processing: E00{t_room} (CR{t_suite})\nFor Testing: E00{t_room} (CR{t_suite})",
            "smart_phase1_summary": smart_phase1_full, "smart_phase1_continued": "",
            "incubation_time": st.session_state.incubation_time, "usp71_id": st.session_state.usp71_id,
            "date_of_weekly": st.session_state.get("date_of_weekly", "N/A"), "weekly_initial": st.session_state.get("weekly_initial", "N/A"),
            "sample_noun": sample_noun, "sample_verb": sample_verb, "bottle_noun": bottle_noun,
            "submit_verb": submit_verb, "org_noun": org_noun,
            # --- 26 missing vars fix ---
            "reader_name": st.session_state.get("reading_name", ""),
            "cr_id": t_room,
            "changeover_initial": st.session_state.get("changeover_initial", ""),
            "chgbsc_id": st.session_state.get("chgbsc_id", ""),
            "event_number": st.session_state.get("event_number", ""),
            "confirm_number": st.session_state.get("confirm_number", ""),
            "organism_morphology": st.session_state.get("organism_morphology", ""),
            "oos1_sample_id": st.session_state.get("oos1_sample_id", ""),
            "oos1_sample_name": st.session_state.get("oos1_sample_name", ""),
            "oos1_analyst_name": st.session_state.get("oos1_analyst_name", ""),
            "oos1_organism_morphology": st.session_state.get("oos1_organism_morphology", st.session_state.get("organism_morphology", ""))
        }

        # Handle any dynamic fields from EM observations
        for k in field_keys:
            if k not in word_data and k in st.session_state:
                word_data[k] = st.session_state[k]

        pdf_map = {
            'Text Field57': st.session_state.oos_id, 'Date Field0': pdf_date_str, 'Date Field1': pdf_date_str, 
            'Date Field2': pdf_date_str, 'Date Field3': pdf_date_str,
            'Text Field2': st.session_state.sample_id, 'Text Field6': st.session_state.lot_number, 
            'Text Field4': st.session_state.sample_name + "\n\n\n\n", 'Text Field5': st.session_state.dosage_form, 
            'Text Field0': analyst_sig_text, 'Text Field3': smart_personnel_block, 'Text Field7': smart_incident_opening + "\n\n",
            'Text Field13': word_data["smart_comment_interview"], 'Text Field14': word_data["smart_comment_samples"], 
            'Text Field17': word_data["smart_comment_records"], 'Text Field21': word_data["smart_comment_storage"],
            'Text Field30': f"E00{st.session_state.bsc_id}", 'Text Field32': word_data["smart_cr_id"], 
            'Text Field34': st.session_state.usp71_id, 'Text Field49': smart_phase1_part1, 'Text Field50': smart_phase1_part2
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
        
        table_data["obs_pers_dur_pro"] = st.session_state.get("obs_pers_dur", "No Growth")
        table_data["etx_pers_dur_pro"] = st.session_state.get("etx_pers_dur", "N/A")
        table_data["id_pers_dur_pro"] = st.session_state.get("id_pers_dur", "N/A")
        
        table_data["obs_surf_dur_pro"] = st.session_state.get("obs_surf_dur", "No Growth")
        table_data["etx_surf_dur_pro"] = st.session_state.get("etx_surf_dur", "N/A")
        table_data["id_surf_dur_pro"] = st.session_state.get("id_surf_dur", "N/A")
        
        table_data["obs_sett_dur_pro"] = st.session_state.get("obs_sett_dur", "No Growth")
        table_data["etx_sett_dur_pro"] = st.session_state.get("etx_sett_dur", "N/A")
        table_data["id_sett_dur_pro"] = st.session_state.get("id_sett_dur", "N/A")
        
        for prefix in ["be_", "af_"]:
            for suffix in ["pers_dur_pro", "surf_dur_pro", "sett_dur_pro"]:
                table_data[f"{prefix}obs_{suffix}"] = "No Growth"
                table_data[f"{prefix}etx_{suffix}"] = "N/A"
                table_data[f"{prefix}id_{suffix}"] = "N/A"
                
        table_data["obs_air_wk_of"] = st.session_state.get("obs_air_wk_of", "No Growth")
        table_data["etx_air_wk_of"] = st.session_state.get("etx_air_wk_of", "N/A")
        table_data["id_air_wk_of"] = st.session_state.get("id_air_wk_of", "N/A")
        
        table_data["obs_room_wk_of"] = st.session_state.get("obs_room_wk_of", "No Growth")
        table_data["etx_room_wk_of"] = st.session_state.get("etx_room_wk_of", "N/A")
        table_data["id_room_wk_of"] = st.session_state.get("id_room_wk_of", "N/A")
        
        table_data["aliquoting_name"] = "N/A"
        table_data["positive_id"] = st.session_state.get("positive_id", "N/A")
        table_data["positive_media"] = st.session_state.get("positive_media", "N/A")
        table_data["positive_org"] = st.session_state.get("positive_org", "N/A")

        # --- 1. RENDER MAIN DOCX ---
        target_template = "USP71 OOS P1 template.docx"
        if not os.path.exists(target_template): target_template = "USP71 OOS P1 template 0.docx"
        if os.path.exists(target_template):
            try:
                from docxtpl import DocxTemplate
                doc = DocxTemplate(target_template)
                doc.render(word_data); docx_buf = io.BytesIO(); doc.save(docx_buf); docx_buf.seek(0)
            except Exception as e: st.error(f"DOCX Error: {e}")
        else: st.warning(f"⚠️ Could not find {target_template}.")

        # --- 2. RENDER TABLES DOCX ---
        target_tables_template = "tables for 71.docx"
        if not os.path.exists(target_tables_template): target_tables_template = "USP71 table.docx"
        if os.path.exists(target_tables_template):
            try:
                from docxtpl import DocxTemplate
                doc_tbl = DocxTemplate(target_tables_template)
                doc_tbl.render(table_data); tables_docx_buf = io.BytesIO(); doc_tbl.save(tables_docx_buf); tables_docx_buf.seek(0)
            except Exception as e: st.error(f"Tables DOCX Error: {e}")
        else: st.warning(f"⚠️ Could not find {target_tables_template}.")
            
        # --- 3. RENDER MAIN PDF ---
        if os.path.exists("USP71 OOS P1 template.pdf"):
            try:
                from pypdf import PdfWriter
                writer = PdfWriter(clone_from="USP71 OOS P1 template.pdf") 
                for p in writer.pages: writer.update_page_form_field_values(p, pdf_map)
                pdf_form_buf = io.BytesIO(); writer.write(pdf_form_buf); pdf_form_buf.seek(0)
            except Exception as e: st.error(f"PDF Form Error: {e}")

        # --- 4. RENDER TABLES PDF ---
        try:
            tables_pdf_buf = create_table_pdf(table_data)
        except Exception as e: st.error(f"Tables PDF Error: {e}")

        st.success("✅ USP 71 Reports and Tables Generated!")
        st.markdown("### 📂 Download Reports")
        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            if docx_buf: st.download_button("📄 USP 71 Report (doc)", docx_buf, f"{safe_filename}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            if tables_docx_buf: st.download_button("📄 Tables (doc)", tables_docx_buf, f"Tables {safe_filename}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with c_dl2:
            if pdf_form_buf: st.download_button("🔴 USP 71 Report (pdf)", pdf_form_buf, f"{safe_filename}.pdf", "application/pdf")
            if tables_pdf_buf: st.download_button("🔴 Tables (pdf)", tables_pdf_buf, f"Tables {safe_filename}.pdf", "application/pdf")

        st.markdown("---")
        current_data = {k: st.session_state[k] for k in field_keys if k in st.session_state}
        st.download_button("💾 Save Session Data (.txt)", json.dumps(current_data, indent=2), f"SAVE_{safe_filename}.txt", "text/plain")
