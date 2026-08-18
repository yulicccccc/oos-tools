# filename: em_logic.py
import streamlit as st
import os
import re
import json
import io
import sys
from datetime import datetime, timedelta

# ReportLab for dynamic Page 7 Table Generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# --- 1. Central Utilities ---
try:
    from utils import get_room_logic as u_grl, get_full_name, ordinal, num_to_words, get_cleanroom_narrative
except ImportError:
    def u_grl(i): return "Unknown", "000", "", "Unknown"
    def get_full_name(i): return i
    def ordinal(n): return str(n)
    def num_to_words(n): return str(n)
    def get_cleanroom_narrative(s, r=None, a="", v=""): return ""

# --- 2. CONFIG & KEYS ---
FIELD_KEYS = [
    "oos_id", "sample_name", "test_date", "sampling_type", "bsc_id",
    "analyst_name", "analyst_initial", "reader_name", "reader_initial",
    "action_level", "cfu_count", "event_number", "manual_org", "monthly_cleaning_date",
    "cleaner_name", "writer_name", "manager_name", "test_method",
    # Media Plate / Reagents
    "plate_media_type", "media_plate_lot", "media_plate_exp",
    # Bracketing fields
    "before_date", "after_date",
    "pers_obs_before", "pers_obs_during", "pers_obs_after",
    "bsc_surf_obs_before", "bsc_surf_obs_during", "bsc_surf_obs_after",
    "bsc_sett_obs_before", "bsc_sett_obs_during", "bsc_sett_obs_after",
    "date_of_weekly_air", "weekly_air_analyst", "air_obs", "air_etx", "air_id",
    "date_of_weekly_surf", "weekly_surf_analyst", "room_surf_obs", "room_surf_etx", "room_surf_id"
]

def auto_fill_name(initial_key, name_key):
    initial = st.session_state.get(initial_key, "")
    current_name = st.session_state.get(name_key, "")
    if initial:
        calculated_name = get_full_name(initial)
        if calculated_name and not current_name:
            st.session_state[name_key] = calculated_name

def validate_inputs():
    errors, warnings = [], []
    reqs = {
        "OOS Number": "oos_id", "Sample / Plate Name": "sample_name", 
        "Test Date": "test_date", "Setup Analyst Name": "analyst_name",
        "BSC / Cleanroom ID": "bsc_id"
    }
    for label, key in reqs.items():
        if not st.session_state.get(key, "").strip(): warnings.append(label)
    date_val = st.session_state.get("test_date", "").strip()
    if date_val:
        try:
            if len(date_val) >= 9:
                datetime.strptime(date_val, "%d%b%Y")
            else:
                datetime.strptime(date_val, "%d%b%y")
        except ValueError:
            errors.append(f"❌ Date Error: '{date_val}' invalid. Use DDMMMYY (e.g. 17Feb26).")
    return errors, warnings

def clean_filename(text): 
    return re.sub(r'[\\/*?:"<>|]', '_', str(text)).strip() if text else ""

def format_date_std(date_str):
    """Converts diverse date strings to DD-Mon-YYYY or DD Mon YYYY format"""
    if not date_str:
        return ""
    clean_d = re.sub(r'[\s\-]', '', str(date_str).strip())
    for fmt in ["%d%b%Y", "%d%b%y", "%Y%m%d", "%m/%d/%Y", "%d/%m/%Y"]:
        try:
            dt = datetime.strptime(clean_d, fmt)
            return dt.strftime("%d-%b-%Y")
        except ValueError:
            continue
    return date_str

def parse_em_text(text):
    """Smart Paste parser for EM email & notification text"""
    data = {}
    if not text or not text.strip():
        return data
        
    # 1. OOS ID
    oos_match = re.search(r"OOS[-\s]*(\d+)", text, re.IGNORECASE)
    if oos_match:
        data["oos_id"] = f"OOS-{oos_match.group(1).strip()}"
        
    # 2. ETX / Event ID
    etx_match = re.search(r"(ETX-\d{6}-\d{4})", text, re.IGNORECASE)
    if etx_match:
        data["event_number"] = etx_match.group(1).strip()
        
    # 3. Plate / Sample Name (e.g. ScanC/O HS BSC1309 S3 17FEB2026 or Sterility GS BSC1314 Sett2 05FEB2026)
    plate_match = re.search(r"((?:Scan|Sterility|EM)[^\t\r\n]+)", text)
    if plate_match:
        p_name = plate_match.group(1).strip()
        data["sample_name"] = p_name
        
        # Extract date from plate name
        date_in_plate = re.search(r"(\d{1,2}\s*[A-Za-z]{3}\s*\d{2,4})", p_name)
        if date_in_plate:
            raw_d = date_in_plate.group(1).replace(" ", "")
            try:
                if len(raw_d) >= 9:
                    d_obj = datetime.strptime(raw_d, "%d%b%Y")
                else:
                    d_obj = datetime.strptime(raw_d, "%d%b%y")
                data["test_date"] = d_obj.strftime("%d%b%y")
            except: pass
            
        # Extract BSC / Equipment ID
        bsc_match = re.search(r"(?:BSC|E00)?(\d{4})", p_name, re.IGNORECASE)
        if bsc_match:
            data["bsc_id"] = f"BSC {bsc_match.group(1)}"
            
        # Extract Setup Analyst Initial
        analyst_match = re.search(r"(?:ScanC/O|ScanCO|Scan|Sterility|EM)\s+([A-Z]{2,3})\b", p_name, re.IGNORECASE)
        if analyst_match:
            init = analyst_match.group(1).upper()
            full_n = get_full_name(init)
            if full_n and full_n != init:
                data["analyst_name"] = full_n
                data["analyst_initial"] = init

        # Infer sampling type
        p_lower = p_name.lower()
        if "sett" in p_lower:
            data["sampling_type"] = "Settling Sampling"
        elif "c/o" in p_lower or "changeover" in p_lower:
            data["sampling_type"] = "Surface Sampling (Changeover)"
        elif any(s in p_lower for s in ["s1", "s2", "s3", "s4", "surf"]):
            data["sampling_type"] = "Surface Sampling"
        elif "glove" in p_lower or "pers" in p_lower:
            data["sampling_type"] = "Personnel Sampling (Glove)"
        elif "cart" in p_lower or "floor" in p_lower or "room" in p_lower or "air" in p_lower:
            data["sampling_type"] = "Weekly Cleanroom Sampling"

    # 4. CFU Count
    cfu_match = re.search(r"(?:Total CFU Count on Plate|CFU Count|CFU)\s*[:\n\r]*\s*(\d+)", text, re.IGNORECASE)
    if cfu_match:
        data["cfu_count"] = cfu_match.group(1).strip()
        
    # 5. Colony Description & Organism Identification
    org_match = re.search(r"(?:Microbial Identification|Colony Description|Organism)\s*(?:\(Optional\))?\s*[:\n\r]*\s*([^\n\r]+)", text, re.IGNORECASE)
    if org_match and org_match.group(1).strip().upper() not in ["N/A", "NONE", ""]:
        data["manual_org"] = org_match.group(1).strip()

    # 6. Reagent / Plate Media Lot & Exp
    lot_match = re.search(r"(?:TSA Lot|Contact Plate Lot|Plate Lot|Lot\s*#?|Media Lot)\s*[:\s]*(\d{7,10})", text, re.IGNORECASE)
    if not lot_match:
        lot_match = re.search(r"\b(1011\d{6})\b", text)
    if lot_match:
        data["media_plate_lot"] = lot_match.group(1).strip()
        
    exp_match = re.search(r"(?:Exp|Expiry|Expiration)\s*[:\s]*(\d{1,2}\s*[A-Za-z]{3}\s*\d{2,4})", text, re.IGNORECASE)
    if not exp_match:
        exp_match = re.search(r"\b(\d{1,2}[A-Za-z]{3}\d{2,4})\b", text)
    if exp_match:
        raw_exp = exp_match.group(1).replace(" ", "").upper()
        data["media_plate_exp"] = raw_exp

    # Defaults
    data["reader_name"] = "Simin Mohammad & Maraya Chukwumerije"
    data["writer_name"] = "Maryam Naeem"
    data["manager_name"] = "Kathan Parikh"
    data["cleaner_name"] = "Rey Estrada"

    return data

# --- 3. NARRATIVE GENERATION LOGIC ---
def generate_em_narrative():
    s = st.session_state
    
    analyst_name = s.get("analyst_name", "Gabrielle Surber")
def compute_em_dates(test_date_str, etx_id=""):
    """
    Computes standard EM incubation milestones and OOS initiation/incident dates.
    - Test Date: Setup date (e.g., 07-May-2026)
    - 48h Read (30-35°C in E001031): +2d (Mon-Wed) or +4d (Thu-Fri)
    - 5-Day Read / Date of Incident / Date Initiated (20-25°C in E001034):
      NLT 5 days later (concluding on business day, e.g., 18-May-2026).
    """
    # 1. Parse setup test_date
    clean_d = re.sub(r'[\s\-]', '', str(test_date_str).strip())
    d_obj = None
    for fmt in ["%d%b%Y", "%d%b%y"]:
        try:
            d_obj = datetime.strptime(clean_d, fmt)
            break
        except: pass

    # 2. Check if ETX ID encodes the discovery/initiation date (e.g. ETX-260518-0254 -> 18-May-2026)
    dt_etx = None
    if etx_id:
        etx_match = re.search(r'ETX-(\d{2})(\d{2})(\d{2})-\d+', str(etx_id), re.IGNORECASE)
        if etx_match:
            yy, mm, dd = etx_match.groups()
            try:
                dt_etx = datetime.strptime(f"20{yy}{mm}{dd}", "%Y%m%d")
            except: pass

    if d_obj:
        test_d_std = d_obj.strftime("%d-%b-%Y")
        d_start = d_obj.strftime("%d %b %Y")
        w = d_obj.weekday() # 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri, 5=Sat, 6=Sun
        
        # 48 hours incubation read date
        if w in [3, 4]: # Thu -> Mon (+4d), Fri -> Tue (+4d)
            d_48h_dt = d_obj + timedelta(days=4)
        else: # Mon -> Wed (+2d), Tue -> Thu (+2d), Wed -> Fri (+2d)
            d_48h_dt = d_obj + timedelta(days=2)
        d_48h = d_48h_dt.strftime("%d %b %Y")

        # NLT 5 days incubation read / initiation date
        if dt_etx:
            d_final_dt = dt_etx
        else:
            if w == 3: # Thu setup -> final read 2nd Mon (+11d)
                d_final_dt = d_obj + timedelta(days=11)
            elif w == 4: # Fri setup -> final read 2nd Mon (+10d)
                d_final_dt = d_obj + timedelta(days=10)
            else: # Mon/Tue/Wed setup -> final read same day next week (+7d)
                d_final_dt = d_obj + timedelta(days=7)
        
        d_5d = d_final_dt.strftime("%d %b %Y")
        date_initiated = d_final_dt.strftime("%d-%b-%Y")
        date_of_incident = date_initiated
        
        before_d = (d_obj - timedelta(days=3 if w == 0 else 1)).strftime("%d %b %Y")
        after_d = (d_obj + timedelta(days=3 if w == 4 else 1)).strftime("%d %b %Y")
    else:
        test_d_std = str(test_date_str)
        d_start = str(test_date_str)
        d_48h = "11 May 2026"
        d_5d = "18 May 2026"
        date_initiated = "18-May-2026"
        date_of_incident = "18-May-2026"
        before_d = "06 May 2026"
        after_d = "08 May 2026"

    return {
        "test_date_std": test_d_std,
        "d_start": d_start,
        "d_48h": d_48h,
        "d_5d": d_5d,
        "date_initiated": date_initiated,
        "date_of_incident": date_of_incident,
        "before_d": before_d,
        "after_d": after_d
    }

def generate_em_narrative():
    """Generates the standardized 3-part Phase I narrative for Environmental Monitoring OOS"""
    s = st.session_state
    analyst_name = s.get("analyst_name", "Gabrielle Surber")
    analyst_init = s.get("analyst_initial", "GS")
    reader_name = s.get("reader_name", "Maraya Chukwumerije and Simin Mohammad")
    sampling_type = s.get("sampling_type", "Settling Sampling")
    bsc_id = s.get("bsc_id", "BSC 1314")
    plate_name = s.get("sample_name", "Sterility GS BSC1314 Sett2 05FEB2026")
    test_date = s.get("test_date", "05 Feb 2026")
    event_id = s.get("event_number", "ETX-260216-0348")
    cfu_count = s.get("cfu_count", "10")
    org_identified = s.get("manual_org", "Staphylococcus capitis (Gram (+) cocci), Staphylococcus hominis (Gram (+) cocci), Kocuria indica (Gram (+) cocci), Micrococcus luteus (Gram (+) cocci) and Staphylococcus epidermidis (Gram (+) cocci)")
    monthly_cleaning_date = s.get("monthly_cleaning_date", "31 Jan 2026")
    cleaner_name = s.get("cleaner_name", "Rey Estrada")
    test_method = s.get("test_method", "USP 71 Sterility" if "Sterility" in plate_name else "ScanRDI")

    # Suite logic based on BSC ID or Plate Name
    combined_id = f"{bsc_id} {plate_name}"
    if "1309" in combined_id or "117A" in combined_id: suite_info, cr_suite = "Suite 117A", "CR117"
    elif "1310" in combined_id or "117B" in combined_id: suite_info, cr_suite = "Suite 117B", "CR117"
    elif "117" in combined_id: suite_info, cr_suite = "Suite 117", "CR117"
    elif "1311" in combined_id or "116A" in combined_id: suite_info, cr_suite = "Suite 116A", "CR116"
    elif "116B" in combined_id: suite_info, cr_suite = "Suite 116B", "CR116"
    elif "116" in combined_id: suite_info, cr_suite = "Suite 116", "CR116"
    elif "1313" in combined_id or "115A" in combined_id: suite_info, cr_suite = "Suite 115A", "CR115"
    elif "1314" in combined_id or "115B" in combined_id: suite_info, cr_suite = "Suite 115B", "CR115"
    elif "115" in combined_id: suite_info, cr_suite = "Suite 115", "CR115"
    elif "114A" in combined_id: suite_info, cr_suite = "Suite 114A", "CR114"
    elif "114B" in combined_id: suite_info, cr_suite = "Suite 114B", "CR114"
    elif "114" in combined_id: suite_info, cr_suite = "Suite 114", "CR114"
    else: suite_info, cr_suite = "Suite 115B", "CR115"

    # Compute Incubation & Milestone Dates
    dates = compute_em_dates(test_date, event_id)
    d_start = dates["d_start"]
    d_48h = dates["d_48h"]
    d_5d = dates["d_5d"]

    is_cleanroom_weekly = any(k in sampling_type.lower() or k in plate_name.lower() for k in ["weekly", "air", "cart", "floor", "cleanroom"])

    # --- 1. Interview & Storage Block (Field 49) ---
    location_desc = f"Cleanroom Suite {suite_info.replace('Suite ', '')}" if is_cleanroom_weekly else f"ISO 5 {bsc_id} located in {suite_info}"
    purpose_desc = "during weekly environmental monitoring" if is_cleanroom_weekly else f"during {test_method} processing"
    
    interview_block = (
        f"The analysts involved in the {sampling_type} plate setup, {analyst_name}, and the analysts involved in "
        f"reading the plate, {reader_name}, were interviewed comprehensively. Their answers are recorded throughout this document. "
        f"The EM plates were stored in compliance with the supplier's recommendations, and their integrity was visually inspected prior to use. "
        f"Furthermore, the plates were confirmed to be within their valid expiration dates. All the supplies were thoroughly disinfected according to SOP 2.600.018. "
        f"The functionality of both incubators was verified through a review of data obtained from our comprehensive in-house continuous monitoring system. "
        f"{sampling_type} was performed by analyst {analyst_name} {purpose_desc}, in {location_desc}, "
        f"on {d_start} as per SOP 2.600.002 - Environmental Monitoring of the Cleanroom Facility. "
        f"The plates were initially incubated at a temperature of 30–35°C in incubator E001031 for a minimum duration of 48 hours, commencing on {d_start}. "
        f"Following completion of minimum of 48 hours of incubation on {d_48h}, the plates were further incubated for minimum of 5 days, with the incubation concluding on {d_5d}. "
        f"Please see Table 1 for detailed information on the observations during respective incubations. "
        f"Based on the observations in Table 1, since the CFU count exceeded the action level for the {sampling_type} plate, the plate was submitted for Microbial Identification under {event_id}. "
        f"The colony was identified as {org_identified}. "
        f"To observe if the organisms identified were transient in nature or recurring, environmental monitoring plates for the analyst {analyst_name} and {location_desc} "
        f"were bracketed to include date before testing and date after testing as detailed in Table 2 (please see attached)."
    )

    # --- 2. EM Records Block (Field 50) ---
    if is_cleanroom_weekly:
        records_block = (
            f"Environmental Monitoring Summary: Weekly surface and active air sampling for {cr_suite} for the previous week and following week of testing showed no microbial growth. "
            f"However, the {sampling_type} for {cr_suite} for the week of testing, performed on {d_start} by analyst {analyst_init}, exhibited {cfu_count} CFUs on {plate_name}, "
            f"which were identified as {org_identified}. Routine monitoring on the date of testing showed no growth across other monitored locations. "
            f"During the interview with the analyst, they indicated that no obvious abnormalities or deviations in the testing procedure were observed. "
            f"All the samples were thoroughly disinfected prior to testing. Moreover, the cleanroom suites in {cr_suite} were thoroughly cleaned and prepared before initiating testing as per SOP 2.600.002 and SOP 2.600.018. "
            f"Monthly cleaning and disinfection of the cleanroom facility and containing Biosafety Cabinets were performed on {monthly_cleaning_date} as per SOP 2.600.018 by analyst {cleaner_name}. "
            f"It was documented that all H2O2 indicators passed. Additionally, cleaning and disinfecting was performed both prior to and after the testing process as per SOP 2.600.018. "
            f"It is important to note that no samples processed within that week in {cr_suite} failed {test_method} component testing that week."
        )
    else:
        records_block = (
            f"Environmental Monitoring Summary: Personnel sampling plate for analyst {analyst_name} for the previous date, date of and following date of testing showed no microbial growth. "
            f"Surface sampling plates for ISO 5 {bsc_id} for the previous date, date of and following date of testing showed no microbial growth either. "
            f"{sampling_type} for the date of testing exhibited {cfu_count} CFU on {plate_name}, performed by analyst {analyst_name}, which was identified as {org_identified}. "
            f"No growth was observed on other routine sampling plates for the date of testing as well as the settling and surface sampling plates for the following and previous date of testing. "
            f"The weekly Active Air and Surface Sampling for Anteroom & Buffer room for {cr_suite} for the week of testing showed no growth either. "
            f"During the interview with the analyst, they indicated that no obvious abnormalities or deviations in the testing procedure were observed. "
            f"All the samples were thoroughly disinfected prior to testing. Moreover, the cleanroom suite and the ISO 5 {bsc_id} were thoroughly cleaned and prepared before initiating the testing as per SOP 2.600.002 and SOP 2.600.018. "
            f"Monthly cleaning and disinfection of the cleanroom facility and containing Biosafety Cabinets were performed on {monthly_cleaning_date} as per SOP 2.600.018 by analyst {cleaner_name}. "
            f"It was documented that all H2O2 indicators passed. Additionally, cleaning and disinfecting was performed both prior to and after the testing process as per SOP 2.600.018. "
            f"It is important to note that no samples processed by analyst {analyst_name} in ISO 5 {bsc_id} on {d_start} failed {test_method} testing that day."
        )

    # --- 3. Phase I Summary / Defensive Conclusion (Field 51) ---
    if is_cleanroom_weekly:
        summary_block = (
            f"Based on the findings outlined in the preceding sections, the Out-Of-Specification (OOS) result observed for the weekly Environmental Monitoring (EM) {sampling_type} plate may be attributed to a potential laboratory or sampling error. "
            f"It is important to note that the product samples are processed exclusively within the ISO 5 Primary Engineering Control (BSC) and are not exposed to the background environment in the ISO 8 or ISO 7 areas. "
            f"Furthermore, all sample containers and supplies are thoroughly disinfected during transfer across the ISO 8 to ISO 7 to ISO 5 cascade. "
            f"Additionally, no sample tested during that week in {cr_suite} failed {test_method} testing, confirming that the elevated environmental reading had no impact on sample integrity. "
            f"Therefore, no preventive and corrective actions are deemed necessary at this time."
        )
    else:
        summary_block = (
            f"Based on the findings outlined in the preceding sections, the Out-Of-Specification (OOS) result observed for the Environmental Monitoring (EM) {sampling_type} plate may be attributed to a potential analyst error or transient laboratory contamination. "
            f"No growth was observed on the analyst's personnel plates as well as on the surface and settling sampling plates collected from the BSC for the following day, indicating that the contamination was transient in nature and that routine daily disinfection procedures were effective in eliminating the contamination. "
            f"Furthermore, no samples processed by analyst {analyst_name} on {d_start} failed {test_method} testing that day, suggesting that the positive EM sample had a minimal impact on the testing environment. "
            f"Additionally, no trend was observed in the analyst’s previous EM data, therefore, no preventive and corrective actions are deemed necessary at this time."
        )

    return interview_block, records_block, summary_block

def build_em_context():
    """Builds a complete context dictionary for rendering DOCX and PDF templates"""
    s = st.session_state
    interview_block, records_block, summary_block = generate_em_narrative()

    analyst_name = s.get('analyst_name', 'Gabrielle Surber')
    analyst_init = s.get('analyst_initial', 'GS')
    reader_name = s.get('reader_name', 'Maraya Chukwumerije & Simin Mohammad')
    sampling_type = s.get('sampling_type', 'Settling Sampling')
    bsc_id = s.get('bsc_id', 'BSC 1314')
    plate_name = s.get('sample_name', 'Sterility GS BSC1314 Sett2 05FEB2026')
    test_date = s.get('test_date', '05 Feb 2026')
    oos_id = s.get('oos_id', 'OOS-260361').replace('OOS-', '')
    event_id = s.get('event_number', 'ETX-260216-0348')
    action_level = s.get('action_level', 'Action Level: ≥ 1 CFU/Plate')
    cfu_count = s.get('cfu_count', '10')
    org_identified = s.get('manual_org', 'Staphylococcus capitis (Gram (+) cocci), Staphylococcus hominis (Gram (+) cocci), Kocuria indica (Gram (+) cocci), Micrococcus luteus (Gram (+) cocci) and Staphylococcus epidermidis (Gram (+) cocci)')
    writer_name = s.get('writer_name', 'Maryam Naeem')

    # Compute Incubation & Milestone Dates
    dates = compute_em_dates(test_date, event_id)
    test_d_std = dates["test_date_std"]
    date_initiated = dates["date_initiated"]
    date_of_incident = dates["date_of_incident"]
    before_d = dates["before_d"]
    after_d = dates["after_d"]

    # Build personnel display block for Section A
    personnel_block = f"{analyst_name}\n({sampling_type} Plate Setup)\n\n{reader_name}\n({sampling_type} Plate Readers)"

    # Signature line
    if writer_name and writer_name.strip() and writer_name.strip() != analyst_name.strip():
        analyst_sig = f"{analyst_name} (written by {writer_name})"
    else:
        analyst_sig = analyst_name

    # Media Plate / Reagent Info
    plate_media_type = s.get('plate_media_type', 'TSA Plate' if ('air' in plate_name.lower() or 'sett' in plate_name.lower()) else 'Contact Plate')
    media_lot = s.get('media_plate_lot', '1011543730')
    media_exp = s.get('media_plate_exp', '29SEP2026')
    reagent_lot_str = f"{plate_media_type}:\n{media_lot}"
    reagent_exp_str = f"{plate_media_type}:\n{media_exp}"

    ctx = {
        # General & Section A
        "oos_id": oos_id,
        "sample_id": event_id,
        "event_number": event_id,
        "etx_id": event_id,
        "sample_name": plate_name,
        "lot_number": plate_name,
        "dosage_form": "Plate",
        "test_date": test_d_std,
        "date_initiated": date_initiated,
        "date_of_incident": date_of_incident,
        "analyst_name": analyst_name,
        "analyst_initial": analyst_init,
        "setup_analyst_initial": analyst_init,
        "reader_name": reader_name,
        "analyst_signature": analyst_sig,
        "analyst_personnel_block": personnel_block,
        "smart_personnel_block": personnel_block,
        "bsc_id": bsc_id,
        "equipment_summary": "Incubator E001031 and Incubator E001034",
        "cr_id": "CR115",
        "action_level": action_level,
        "incident_description": "The CFU count for the environmental monitoring plate exceeded the action level.",
        "smart_incident_opening": "The CFU count for the environmental monitoring plate exceeded the action level.",
        "report_header": event_id,
        "client_name": "Eagle Analytical Internal EM",

        # Narratives
        "smart_comment_interview": f"Yes, analysts {analyst_name} and {reader_name} were interviewed comprehensively.",
        "smart_comment_records": f"Yes, Information is available in EagleTrax under {event_id}.",
        "smart_comment_samples": "Yes, as per SOP 2.600.002",
        "smart_comment_storage": "Yes, as per SOP 2.600.002",
        "narrative_summary": f"{interview_block}\n\n{records_block}\n\n{summary_block}",
        "smart_phase1_summary": summary_block,
        "smart_phase1_continued": "",
        "smart_phase1_part1": interview_block,
        "smart_phase1_part2": summary_block,

        # Media Plate / Reagent Info
        "plate_media_type": plate_media_type,
        "media_plate_lot": media_lot,
        "media_plate_exp": media_exp,
        "reagent_lot": reagent_lot_str,
        "reagent_exp": reagent_exp_str,

        # Table 1 Fields
        "sampling_location": f"{sampling_type} Plate ({bsc_id})",
        "reader_48h": "MC" if "Maraya" in reader_name else "SMO",
        "cfu_obs_48h": f"{cfu_count} CFU on {plate_name}",
        "reader_5d": "SMO",
        "cfu_obs_5d": f"{cfu_count} CFU on {plate_name}",
        "microbial_id": org_identified,

        # Table 2 Bracketing Fields
        "before_date": before_d,
        "after_date": after_d,
        "pers_obs_before": "No growth", "pers_etx_before": "N/A", "pers_id_before": "N/A",
        "pers_obs_during": "No growth", "pers_etx_during": "N/A", "pers_id_during": "N/A",
        "pers_obs_after": "No growth", "pers_etx_after": "N/A", "pers_id_after": "N/A",

        "bsc_surf_analyst_before": analyst_init, "bsc_surf_obs_before": "No growth", "bsc_surf_etx_before": "N/A", "bsc_surf_id_before": "N/A",
        "bsc_surf_analyst_during": analyst_init, "bsc_surf_obs_during": f"{cfu_count} CFU" if "Surface" in sampling_type else "No growth", "bsc_surf_etx_during": event_id if "Surface" in sampling_type else "N/A", "bsc_surf_id_during": org_identified if "Surface" in sampling_type else "N/A",
        "bsc_surf_analyst_after": analyst_init, "bsc_surf_obs_after": "No growth", "bsc_surf_etx_after": "N/A", "bsc_surf_id_after": "N/A",

        "bsc_sett_analyst_before": analyst_init, "bsc_sett_obs_before": "No growth", "bsc_sett_etx_before": "N/A", "bsc_sett_id_before": "N/A",
        "bsc_sett_analyst_during": analyst_init, "bsc_sett_obs_during": f"{cfu_count} CFU" if "Settling" in sampling_type else "No growth", "bsc_sett_etx_during": event_id if "Settling" in sampling_type else "N/A", "bsc_sett_id_during": org_identified if "Settling" in sampling_type else "N/A",
        "bsc_sett_analyst_after": analyst_init, "bsc_sett_obs_after": "No growth", "bsc_sett_etx_after": "N/A", "bsc_sett_id_after": "N/A",
        "date_of_weekly_air": dates["d_start"], "weekly_air_analyst": "SMO", "air_obs": "No growth", "air_etx": "N/A", "air_id": "N/A",
        "date_of_weekly_surf": dates["d_start"], "weekly_surf_analyst": "SMO", "room_surf_obs": "No growth", "room_surf_etx": "N/A", "room_surf_id": "N/A"
    }

    return ctx

def generate_em_tables_page_pdf(context_data):
    """
    Generates a 1-page PDF buffer containing Table 1 and Table 2 using ReportLab Platypus.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        leftMargin=30,
        rightMargin=30,
        topMargin=30,
        bottomMargin=30
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TableTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=8.5,
        leading=10,
        textColor=colors.black,
        spaceAfter=4
    )
    
    cell_hdr_style = ParagraphStyle(
        'HeaderCell',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=6.5,
        leading=8,
        alignment=TA_CENTER,
        textColor=colors.black
    )
    
    cell_body_style = ParagraphStyle(
        'BodyCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=6,
        leading=7.5,
        alignment=TA_CENTER,
        textColor=colors.black
    )

    cell_body_left = ParagraphStyle(
        'BodyCellLeft',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=6,
        leading=7.5,
        alignment=TA_LEFT,
        textColor=colors.black
    )

    cell_section_hdr = ParagraphStyle(
        'SectionHdr',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=6.5,
        leading=8,
        alignment=TA_LEFT,
        textColor=colors.black
    )

    elements = []

    # --- TABLE 1 ---
    elements.append(Paragraph("Table 1: Read Dates & Incubation Observation", title_style))
    
    t1_headers = [
        Paragraph("ETX Submission<br/>ID", cell_hdr_style),
        Paragraph("Set-up Analyst<br/>& Date", cell_hdr_style),
        Paragraph("Sampling Site &<br/>Location", cell_hdr_style),
        Paragraph("Plate Reading<br/>Analyst (≥ 48H)", cell_hdr_style),
        Paragraph("CFUs Observed after 48 Hour Incubation at 30–35°C (E001031)", cell_hdr_style),
        Paragraph("Plate Reading<br/>Analyst (NLT 5 days)", cell_hdr_style),
        Paragraph("CFUs Observed after NLT 5-day Incubation at 20–25°C (E001034)", cell_hdr_style),
        Paragraph("Microbial Identification", cell_hdr_style)
    ]

    t1_row_vals = [
        Paragraph(context_data.get('etx_id', ''), cell_body_style),
        Paragraph(f"{context_data.get('setup_analyst_initial', '')}<br/>{context_data.get('test_date', '')}", cell_body_style),
        Paragraph(context_data.get('sampling_location', ''), cell_body_style),
        Paragraph(context_data.get('reader_48h', 'MC'), cell_body_style),
        Paragraph(context_data.get('cfu_obs_48h', ''), cell_body_style),
        Paragraph(context_data.get('reader_5d', 'SMO'), cell_body_style),
        Paragraph(context_data.get('cfu_obs_5d', ''), cell_body_style),
        Paragraph(context_data.get('microbial_id', ''), cell_body_left)
    ]

    t1_data = [t1_headers, t1_row_vals]
    t1_colWidths = [62, 55, 65, 45, 85, 45, 85, 110]
    
    t1 = Table(t1_data, colWidths=t1_colWidths)
    t1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D9D9D9')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ]))
    elements.append(t1)
    elements.append(Spacer(1, 8))

    # --- TABLE 2 ---
    elements.append(Paragraph("Table 2: Environmental Monitoring for Analyst & Cleanroom", title_style))
    
    t2_headers = [
        Paragraph("Environmental Monitoring<br/>(EM) Sampling Site", cell_hdr_style),
        Paragraph("Frequency", cell_hdr_style),
        Paragraph("Date<br/>(DDMMMYYYY)", cell_hdr_style),
        Paragraph("Analyst<br/>(Initials)", cell_hdr_style),
        Paragraph("Day /Week(s)", cell_hdr_style),
        Paragraph("Observation", cell_hdr_style),
        Paragraph("Plate<br/>ETX ID", cell_hdr_style),
        Paragraph("Microbial ID", cell_hdr_style),
        Paragraph("Notes", cell_hdr_style)
    ]

    t2_colWidths = [115, 38, 48, 38, 62, 65, 52, 98, 36]
    t2_data = [t2_headers]
    table_styles = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D9D9D9')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
    ]

    def add_section(title, rows):
        r_idx = len(t2_data)
        hdr_cell = Paragraph(title, cell_section_hdr)
        t2_data.append([hdr_cell] + [''] * 8)
        table_styles.append(('SPAN', (0, r_idx), (-1, r_idx)))
        table_styles.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.HexColor('#F2F2F2')))
        table_styles.append(('TOPPADDING', (0, r_idx), (-1, r_idx), 2))
        table_styles.append(('BOTTOMPADDING', (0, r_idx), (-1, r_idx), 2))
        
        for r in rows:
            data_row = []
            for col_i, text in enumerate(r):
                if col_i in [0, 7]:
                    data_row.append(Paragraph(str(text), cell_body_left))
                else:
                    data_row.append(Paragraph(str(text), cell_body_style))
            t2_data.append(data_row)

    # 1. Personnel
    add_section("Personnel EM Bracketing", [
        ["Personal (Left & Right Touch)", "Daily", context_data.get('before_date', ''), context_data.get('analyst_initial', ''), "Date Before Testing", context_data.get('pers_obs_before', 'No growth'), context_data.get('pers_etx_before', 'N/A'), context_data.get('pers_id_before', 'N/A'), "None"],
        ["Personal (Left & Right Touch)", "Daily", context_data.get('test_date', ''), context_data.get('analyst_initial', ''), "Date of Testing", context_data.get('pers_obs_during', 'No growth'), context_data.get('pers_etx_during', 'N/A'), context_data.get('pers_id_during', 'N/A'), "None"],
        ["Personal (Left & Right Touch)", "Daily", context_data.get('after_date', ''), context_data.get('analyst_initial', ''), "Date After Testing", context_data.get('pers_obs_after', 'No growth'), context_data.get('pers_etx_after', 'N/A'), context_data.get('pers_id_after', 'N/A'), "None"],
    ])

    # 2. BSC
    add_section("Biological Safety Cabinet (BSC) EM Bracketing", [
        ["Surface Sampling of ISO 5 (4 loc)", "Daily", context_data.get('before_date', ''), context_data.get('bsc_surf_analyst_before', ''), "Date Before Testing", context_data.get('bsc_surf_obs_before', 'No growth'), context_data.get('bsc_surf_etx_before', 'N/A'), context_data.get('bsc_surf_id_before', 'N/A'), "None"],
        ["Surface Sampling of ISO 5 (4 loc)", "Daily", context_data.get('test_date', ''), context_data.get('bsc_surf_analyst_during', ''), "Date of Testing", context_data.get('bsc_surf_obs_during', 'No growth'), context_data.get('bsc_surf_etx_during', 'N/A'), context_data.get('bsc_surf_id_during', 'N/A'), "None"],
        ["Surface Sampling of ISO 5 (4 loc)", "Daily", context_data.get('after_date', ''), context_data.get('bsc_surf_analyst_after', ''), "Date After Testing", context_data.get('bsc_surf_obs_after', 'No growth'), context_data.get('bsc_surf_etx_after', 'N/A'), context_data.get('bsc_surf_id_after', 'N/A'), "None"],
        ["Settling Sampling of ISO 5 (2 loc)", "Daily", context_data.get('before_date', ''), context_data.get('bsc_sett_analyst_before', ''), "Date Before Testing", context_data.get('bsc_sett_obs_before', 'No growth'), context_data.get('bsc_sett_etx_before', 'N/A'), context_data.get('bsc_sett_id_before', 'N/A'), "None"],
        ["Settling Sampling of ISO 5 (2 loc)", "Daily", context_data.get('test_date', ''), context_data.get('bsc_sett_analyst_during', ''), "Date of Testing", context_data.get('bsc_sett_obs_during', '10 CFU'), context_data.get('etx_id', 'ETX-260216-0348'), context_data.get('microbial_id', 'Staphylococcus...'), "None"],
        ["Settling Sampling of ISO 5 (2 loc)", "Daily", context_data.get('after_date', ''), context_data.get('bsc_sett_analyst_after', ''), "Date After Testing", context_data.get('bsc_sett_obs_after', 'No growth'), context_data.get('bsc_sett_etx_after', 'N/A'), context_data.get('bsc_sett_id_after', 'N/A'), "None"],
    ])

    # 3. Weekly Air
    add_section("Weekly Active Air Sampling Bracketing", [
        ["Active Air Sampling of Cleanrooms", "Weekly", context_data.get('date_of_weekly_air', ''), context_data.get('weekly_air_analyst', 'SMO'), "Week (On or After Date)", context_data.get('air_obs', 'No growth'), context_data.get('air_etx', 'N/A'), context_data.get('air_id', 'N/A'), "None"],
    ])

    # 4. Weekly Surface
    add_section("Surface Sampling of Anteroom & Cleanroom Bracketing", [
        ["Surface Sampling of Cleanrooms", "Weekly", context_data.get('date_of_weekly_surf', ''), context_data.get('weekly_surf_analyst', 'SMO'), "Week (On or After Date)", context_data.get('room_surf_obs', 'No growth'), context_data.get('room_surf_etx', 'N/A'), context_data.get('room_surf_id', 'N/A'), "None"],
    ])

    t2 = Table(t2_data, colWidths=t2_colWidths)
    t2.setStyle(TableStyle(table_styles))
    elements.append(t2)

    doc.build(elements)
    buf.seek(0)
    return buf

def generate_em_reports():
    """Generates DOCX and 7-page PDF buffers for EM OOS Report"""
    ctx = build_em_context()
    interview_block, records_block, summary_block = generate_em_narrative()

    docx_buf = None
    pdf_form_buf = None

    # 1. Render Word Template
    target_docx = "EM OOS P1 template.docx" if os.path.exists("EM OOS P1 template.docx") else "EM OOS P1 template 0.docx"
    if os.path.exists(target_docx):
        try:
            from docxtpl import DocxTemplate
            doc = DocxTemplate(target_docx)
            doc.render(ctx)
            docx_buf = io.BytesIO()
            doc.save(docx_buf)
            docx_buf.seek(0)
        except Exception as e:
            st.error(f"DOCX Generation Error: {e}")

    # 2. Render 7-Page PDF Report
    target_pdf = "EM OOS P1 template.pdf"
    if os.path.exists(target_pdf):
        try:
            from pypdf import PdfWriter, PdfReader
            
            # Map 157 Form 3.100.019.F01 fields
            pdf_map = {
                'Text Field57': ctx.get('oos_id', ''),
                'Text Field0': ctx.get('analyst_signature', ''),
                'Date Field0': ctx.get('test_date', ''),
                'Date Field1': ctx.get('date_initiated', ''),
                'Date Field2': ctx.get('date_of_incident', ''),
                'Date Field3': ctx.get('date_initiated', ''),
                'Text Field1': "Environmental Monitoring",
                'Text Field2': ctx['event_number'],
                'Text Field3': ctx['analyst_personnel_block'],
                'Text Field4': ctx['sample_name'],
                'Text Field5': "Plate",
                'Text Field6': ctx['lot_number'],
                'Text Field7': ctx['incident_description'],
                'Text Field8': "2.600.002",
                'Text Field9': "05 Aug 2025",
                'Text Field10': "15",
                'Text Field11': ctx['action_level'].replace("≥", ">="),
                'Text Field12': ctx.get('manager_name', 'Kathan Parikh'),
                'Text Field13': ctx['smart_comment_interview'],
                'Text Field14': "N/A",
                'Text Field15': "Yes, as per SOP 2.600.002",
                'Text Field16': "Yes, as per SOP 2.600.002",
                'Text Field17': ctx['smart_comment_records'],
                'Text Field18': "Yes, the analysts are trained and qualified by quality to perform the test.",
                'Text Field19': "Not Applicable",
                'Text Field20': "Not Applicable",
                'Text Field21': "Yes, as per SOP 2.600.002",
                'Text Field22': ctx.get('reagent_lot', "TSA Plate:\n1011543730"),
                'Text Field23': ctx.get('reagent_exp', "TSA Plate:\n29SEP2026"),
                'Text Field24': "Not Applicable",
                'Text Field25': "Not Applicable",
                'Text Field26': "Not Applicable",
                'Text Field27': "Not Applicable",
                'Text Field28': "Not Applicable",
                'Text Field29': "Not Applicable",
                'Text Field30': "Please see below",
                'Text Field31': "Please see below",
                'Text Field34': "Not Applicable",
                'Text Field35': "Not Applicable",
                'Text Field36': "Not Applicable",
                'Text Field37': "Not Applicable",
                'Text Field38': "Not Applicable",
                'Text Field39': "Not Applicable",
                'Text Field40': "Not Applicable",
                'Text Field41': "Not Applicable",
                'Text Field42': "Not Applicable",
                'Text Field45': "Not Applicable",
                'Text Field46': "Not Applicable",
                'Text Field47': "Not Applicable",
                'Text Field49': interview_block,
                'Text Field50': records_block,
                'Text Field51': summary_block,
                'Text Field52': "Transient contamination. No impact to product quality.",
                'Text Field53': "Maryam Naeem",
                'Text Field54': "Kathan Parikh"
            }

            # Checkbox Yes/No defaults matching production PDF QA standards (EM is internal facility testing, so Client Care and Sample Discrepancy rows are N/A)
            yes_boxes = {4, 9, 10, 13, 16, 19, 24, 27, 28, 33, 36, 39, 42, 43, 48, 51, 52, 55, 60, 63, 66, 69, 72, 73, 78, 79, 87}
            for i in range(100):
                if i in yes_boxes:
                    pdf_map[f'Check Box{i}'] = '/Yes'
                else:
                    pdf_map[f'Check Box{i}'] = ''

            # Fill Form 1-6
            writer = PdfWriter(clone_from=target_pdf)
            for page in writer.pages:
                writer.update_page_form_field_values(page, pdf_map)
                
            # Generate Page 7 Attachment Table
            page7_pdf_buf = generate_em_tables_page_pdf(ctx)
            p7_reader = PdfReader(page7_pdf_buf)
            writer.add_page(p7_reader.pages[0])

            # Output final merged 7-page PDF
            pdf_form_buf = io.BytesIO()
            writer.write(pdf_form_buf)
            pdf_form_buf.seek(0)
        except Exception as e:
            st.error(f"PDF Generation Error: {e}")

    return docx_buf, pdf_form_buf
