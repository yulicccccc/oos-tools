# filename: em_logic.py
import streamlit as st
import os
import re
import json
import io
import sys
import subprocess
import time
from datetime import datetime, timedelta

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
    "oos_id", "client_name", "sample_id", "test_date", "sample_name", "lot_number", 
    "dosage_form", "monthly_cleaning_date", 
    "analyst_initial", "analyst_name", "reader_initial", "reader_name",
    "writer_name", "bsc_id", "cr_id", "shift_number", "sampling_type",
    "org_choice", "manual_org", "action_level", "cfu_count", "event_number", "confirm_number",
    "obs_pers_dur", "etx_pers_dur", "id_pers_dur", 
    "obs_surf_dur", "etx_surf_dur", "id_surf_dur", 
    "obs_sett_dur", "etx_sett_dur", "id_sett_dur", 
    "obs_air_wk_of", "etx_air_wk_of", "id_air_wk_of", 
    "obs_room_wk_of", "etx_room_wk_of", "id_room_wk_of",
    "date_of_weekly", "weekly_initial"
]

def auto_fill_name(initial_key, name_key):
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
        "OOS Number": "oos_id", "Sample / Plate Name": "sample_name", 
        "Test Date": "test_date", "Setup Analyst Name": "analyst_name",
        "Reader Name": "reader_name", "BSC / Cleanroom ID": "bsc_id"
    }
    for label, key in reqs.items():
        if not st.session_state.get(key, "").strip(): warnings.append(label)
    date_val = st.session_state.get("test_date", "").strip()
    if date_val:
        try: datetime.strptime(date_val, "%d%b%y")
        except ValueError: errors.append(f"❌ Date Error: '{date_val}' invalid. Use DDMMMYY (e.g. 17Feb26).")
    return errors, warnings

def clean_filename(text): 
    return re.sub(r'[\\/*?:"<>|]', '_', str(text)).strip() if text else ""

def parse_em_text(text):
    """Smart Paste parser for EM email & notification text"""
    data = {}
    if not text or not text.strip():
        return data
        
    # 1. OOS ID
    oos_match = re.search(r"(OOS-\d+)", text)
    if oos_match:
        data["oos_id"] = oos_match.group(1).strip()
        
    # 2. ETX / Event ID
    etx_match = re.search(r"(ETX-\d{6}-\d{4})", text)
    if etx_match:
        data["event_number"] = etx_match.group(1).strip()
        
    # 3. Plate / Sample Name (e.g. ScanC/O CGS E001309 S1 11MAY2026)
    plate_match = re.search(r"((?:Scan|Sterility|EM)[^\t\r\n]+)", text)
    if plate_match:
        p_name = plate_match.group(1).strip()
        data["sample_name"] = p_name
        
        # Extract date from plate name (e.g. 11MAY2026 or 17FEB2026)
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
            
        # Extract BSC / Equipment ID (e.g. BSC1309, E001309, 1309)
        bsc_match = re.search(r"(?:BSC|E00)?(\d{4})", p_name, re.IGNORECASE)
        if bsc_match:
            data["bsc_id"] = f"BSC {bsc_match.group(1)}"
            
        # Extract Setup Analyst Initial from plate name (e.g. ScanC/O CGS -> CGS -> Clea S. Garza)
        analyst_match = re.search(r"(?:ScanC/O|ScanCO|Scan|Sterility|EM)\s+([A-Z]{2,3})\b", p_name, re.IGNORECASE)
        if analyst_match:
            init = analyst_match.group(1).upper()
            full_n = get_full_name(init)
            if full_n and full_n != init:
                data["analyst_name"] = full_n
                data["analyst_initial"] = init

        # Infer sampling type
        if "sett" in p_name.lower():
            data["sampling_type"] = "Settling Sampling"
        elif "s1" in p_name.lower() or "s2" in p_name.lower() or "s3" in p_name.lower() or "c/o" in p_name.lower() or "surf" in p_name.lower():
            data["sampling_type"] = "Surface Sampling"
        elif "glove" in p_name.lower() or "pers" in p_name.lower():
            data["sampling_type"] = "Personnel Sampling (Glove)"

    # 4. CFU Count
    cfu_match = re.search(r"Total CFU Count on Plate\s*\n?\s*(\d+)", text, re.IGNORECASE)
    if cfu_match:
        data["cfu_count"] = cfu_match.group(1).strip()
        
    # 5. Colony Description & Organism Identification
    desc_match = re.search(r"Colony Description \(Optional\)\s*\n?\s*([^\n\r]+)", text, re.IGNORECASE)
    if desc_match and desc_match.group(1).strip().upper() != "N/A":
        data["manual_org"] = desc_match.group(1).strip()
        
    org_match = re.search(r"Microbial Identification \(Optional\)\s*\n?\s*([^\n\r]+)", text, re.IGNORECASE)
    if org_match and org_match.group(1).strip().upper() != "N/A":
        if "manual_org" in data:
            data["manual_org"] += f" ({org_match.group(1).strip()})"
    # Default Reader Name
    data["reader_name"] = "Simin Mohammad"

    return data

# --- 3. NARRATIVE GENERATION LOGIC ---
def generate_em_narrative():
    s = st.session_state
    
    analyst_name = s.get("analyst_name", "[Analyst Name]")
    reader_name = s.get("reader_name", "[Reader Name]")
    sampling_type = s.get("sampling_type", "Surface Sampling")
    bsc_id = s.get("bsc_id", "BSC 1309")
    plate_name = s.get("sample_name", "EM Plate")
    cfu_count = s.get("cfu_count", "10")
    org_identified = s.get("manual_org", "Staphylococcus epidermidis")
    
    # 1. Interview & Storage Block
    interview_block = (
        f"The analyst involved in the {sampling_type} plate setup, {analyst_name}, and the analyst(s) involved "
        f"in reading the plate, {reader_name}, were interviewed comprehensively. Their answers are recorded throughout this document.\n"
        f"The EM plates were stored in compliance with the supplier's recommendations, and their integrity was visually inspected prior to use. "
        f"Furthermore, the plates were confirmed to be within their valid expiration dates. All the supplies were thoroughly disinfected prior to use."
    )
    
    # 2. EM Records Block
    records_block = (
        f"Environmental Monitoring Summary: Personnel sampling plates for analyst {analyst_name} and "
        f"routine monitoring plates for ISO 5 {bsc_id}, for the previous date, date of, and following date of testing showed no microbial growth. "
        f"However, the {sampling_type} plate for ISO 5 {bsc_id} exhibited {cfu_count} CFUs on {plate_name} for the day of testing "
        f"performed by analyst {analyst_name}, which were identified as {org_identified}. "
        f"No growth was observed on other surface/settling plates for the date of and following date of testing."
    )
    
    # 3. Phase I Summary / Defensive Conclusion
    summary_block = (
        f"Based on the findings outlined in the preceding sections, the Out-Of-Specification (OOS) result observed for "
        f"the Environmental Monitoring (EM) {sampling_type} plate may be attributed to a potential analyst error or transient laboratory contamination.\n"
        f"No growth was observed on the analyst's personnel/surface/settling plates collected from the BSC for the following day, "
        f"indicating that the contamination was transient in nature and that routine daily disinfection procedures were effective in eliminating any residual contamination. "
        f"Furthermore, all testing in the ISO 5 {bsc_id} occurred under controlled conditions with proper physical isolation from the background environment."
    )
    
    return interview_block, records_block, summary_block

def generate_em_reports():
    """Generates DOCX and PDF buffers for EM OOS Report"""
    s = st.session_state
    interview_block, records_block, summary_block = generate_em_narrative()
    
    analyst_name = s.get('analyst_name', '')
    analyst_init = s.get('analyst_initial', '')
    reader_name = s.get('reader_name', 'Simin Mohammad')
    reader_init = "SMO" if "Simin" in reader_name else "MC"
    event_id = s.get("event_number", s.get("sample_id", ""))
    test_date = s.get("test_date", "")
    sample_name = s.get("sample_name", "")
    oos_id = s.get("oos_id", "")
    bsc_id = s.get("bsc_id", "BSC 1309")
    action_level = s.get("action_level", "Action Level: >= 1 CFU/Plate").replace("≥", ">=")
    
    docx_buf = None
    pdf_form_buf = None
    
    # 1. Render DOCX Template
    target_docx = "EM OOS P1 template 0.docx" if os.path.exists("EM OOS P1 template 0.docx") else "EM OOS P1 template.docx"
    if os.path.exists(target_docx):
        try:
            from docxtpl import DocxTemplate
            doc = DocxTemplate(target_docx)
            context = {
                "oos_id": oos_id,
                "sample_id": event_id,
                "sample_name": sample_name,
                "lot_number": sample_name,
                "test_date": test_date,
                "analyst_name": analyst_name,
                "analyst_initial": analyst_init,
                "prepper_name": analyst_name,
                "prepper_initial": analyst_init,
                "changeover_name": analyst_name,
                "changeover_initial": analyst_init,
                "reader_name": reader_name,
                "reader_initial": reader_init,
                "bsc_id": bsc_id,
                "chgbsc_id": bsc_id,
                "cr_id": "CR114",
                "cr_suit": "114",
                "dosage_form": "Plate",
                "client_name": "Eagle Analytical Internal EM",
                "equipment_summary": "Incubator E001034 (Sensor E001501) and Incubator E001031 (Sensor E001505)",
                "narrative_summary": f"{interview_block}\n\n{records_block}\n\n{summary_block}",
                "smart_comment_interview": interview_block,
                "smart_comment_records": records_block,
                "smart_phase1_summary": summary_block,
                "smart_phase1_part1": interview_block,
                "smart_phase1_part2": summary_block,
                "cross_contamination_summary": "No other samples processed concurrently yielded positive growth.",
                "sample_history_paragraph": "Environmental Monitoring plate routine sampling history shows zero prior action level hits.",
                "obs_pers_dur": "0 CFU", "obs_surf_dur": "0 CFU", "obs_sett_dur": "0 CFU", "obs_air_wk_of": "0 CFU", "obs_room_wk_of": "0 CFU",
                "etx_pers_dur": "N/A", "etx_surf_dur": "N/A", "etx_sett_dur": "N/A", "etx_air_wk_of": "N/A", "etx_room_wk_of": "N/A",
                "id_pers_dur": "N/A", "id_surf_dur": "N/A", "id_sett_dur": "N/A", "id_air_wk_of": "N/A", "id_room_wk_of": "N/A"
            }
            doc.render(context)
            docx_buf = io.BytesIO()
            doc.save(docx_buf)
            docx_buf.seek(0)
        except Exception as e:
            st.error(f"DOCX Generation Error: {e}")
            
    # 2. Render PDF Form Template
    target_pdf = "EM OOS P1 template.pdf"
    if os.path.exists(target_pdf):
        try:
            from pypdf import PdfWriter
            initiator_sig = f"{analyst_name} (written by: Qiyue Chen)" if analyst_name else ""
            
            pdf_map = {
                'Text Field57': oos_id,
                'Text Field0': initiator_sig,
                'Date Field0': test_date,
                'Date Field1': test_date,
                'Date Field2': test_date,
                'Text Field1': "Environmental Monitoring",
                'Text Field2': event_id,
                'Text Field3': f"Setup Analyst:\n{analyst_name} ({analyst_init})\n\nReader Analyst:\n{reader_name} ({reader_init})",
                'Text Field4': sample_name,
                'Text Field5': "Plate",
                'Text Field6': sample_name,
                'Text Field7': "The CFU count for the environmental monitoring plate exceeded the action level.",
                'Text Field8': "2.600.002",
                'Text Field11': action_level,
                'Text Field13': f"Yes, {analyst_name} and {reader_name} were interviewed comprehensively.",
                'Text Field14': "N/A",
                'Text Field15': "Yes, as per SOP 2.600.002",
                'Text Field16': "Yes, as per SOP 2.600.002",
                'Text Field17': f"Yes, Information is available in EagleTrax under {event_id}." if event_id else "Yes, as per SOP 2.600.002",
                'Text Field18': "Yes, the analysts are trained and qualified by quality to perform the test.",
                'Text Field19': "Not Applicable",
                'Text Field20': "Not Applicable",
                'Text Field21': "Yes, as per SOP 2.600.002",
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
                'Text Field53': "Maryam Naeem",
                'Text Field54': "Kathan Parikh"
            }
            
            # Checkbox Yes defaults
            yes_boxes = [4, 9, 10, 13, 16, 19, 24, 27, 28, 32, 34, 38, 42, 43, 48, 51, 52, 55, 60, 63, 66, 69, 72, 73, 78, 87, 79]
            for b_num in yes_boxes:
                pdf_map[f'Check Box{b_num}'] = '/Yes'
                
            writer = PdfWriter(clone_from=target_pdf)
            for page in writer.pages:
                writer.update_page_form_field_values(page, pdf_map)
            pdf_form_buf = io.BytesIO()
            writer.write(pdf_form_buf)
            pdf_form_buf.seek(0)
        except Exception as e:
            st.error(f"PDF Form Generation Error: {e}")
            
    return docx_buf, pdf_form_buf
