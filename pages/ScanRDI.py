import streamlit as st
import os
import re
import json
import io
import sys
import subprocess
from datetime import datetime

# --- 1. PAGE CONFIG & IMPORTS ---
st.set_page_config(page_title="ScanRDI Investigation", layout="wide")

try:
    from utils import apply_eagle_style, get_full_name, get_room_logic
except ImportError:
    # Fallback functions if utils.py is missing
    def apply_eagle_style(): pass
    def get_full_name(i): return i
    def get_room_logic(i): return "Unknown", "000", "", "Unknown Loc"

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

# --- 2. GLOBAL CONSTANTS ---
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
    
    # --- Phase 2 Fields ---
    "include_phase2",
    "retest_date", "retest_sample_id", "retest_result",
    "retest_prepper_name", "retest_prepper_initial",
    "retest_analyst_name", "retest_analyst_initial",
    "retest_reader_name", "retest_reader_initial",
    "retest_changeover_name", "retest_changeover_initial",
    "retest_bsc_id", "retest_chgbsc_id", "retest_scan_id",
    "diff_retest_reader", "diff_retest_changeover", "diff_retest_bsc"
]
for i in range(20):
    field_keys.extend([f"other_id_{i}", f"other_order_{i}", f"prior_oos_{i}", f"em_cat_{i}", f"em_obs_{i}", f"em_etx_{i}", f"em_id_{i}"])

# --- 3. STATE FUNCTIONS ---
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

def init_state(key, default=""): 
    if key not in st.session_state: st.session_state[key] = default

# --- 4. ALL HELPER FUNCTIONS ---

def ensure_dependencies():
    required = ["docxtpl", "pypdf", "reportlab"]
    missing = []
    for lib in required:
        try: __import__(lib)
        except ImportError: missing.append(lib)
    if missing:
        st.warning(f"⚙️ Installing: {', '.join(missing)}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)
        st.rerun()

def get_room_logic(bsc_id):
    try: from utils import get_room_logic as u_grl; return u_grl(bsc_id)
    except: return "Unknown", "000", "", "Unknown Loc"

def ordinal(n): return f"{n}th" if 11<=int(n)%100<=13 else f"{n}{['th','st','nd','rd'][int(n)%10 if int(n)%10<=3 else 0]}"

# --- CORE LOGIC: EQUIPMENT GENERATOR ---
def generate_equipment_text(bsc_main, bsc_chg, analyst_main, analyst_chg, date_val, is_retest=False):
    t_room, t_suite, t_suffix, t_loc = get_room_logic(bsc_main)
    c_room, c_suite, c_suffix, c_loc = get_room_logic(bsc_chg)
    
    subj = "Retest sample" if is_retest else "Sample"
    
    # Scenario 1: Same BSC
    if bsc_main == bsc_chg:
        part1 = (f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected "
                 f"sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), "
                 f"and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout "
                 f"the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}.")
        part2 = (f"The ISO 5 BSC E00{bsc_main}, located in the {t_loc}, (Suite {t_suite}{t_suffix}), was used for both testing "
                 f"and changeover steps. It was thoroughly cleaned and disinfected prior to each procedure in accordance with "
                 f"SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). Additionally, BSC E00{bsc_main} "
                 f"was certified and approved by both the Engineering and Quality Assurance teams. {subj} processing and changeover "
                 f"were conducted in the ISO 5 BSC E00{bsc_main} in the {t_loc}, (Suite {t_suite}{t_suffix}) by {analyst_main} on {date_val}.")
        return f"{part1}\n\n{part2}"

    # Scenario 2: Diff BSC, Same Suite
    elif t_suite == c_suite:
        part1 = (f"The cleanroom used for testing and changeover procedures (Suite {t_suite}) comprises three interconnected "
                 f"sections: the innermost ISO 7 cleanroom ({t_suite}B), which connects to the middle ISO 7 buffer room ({t_suite}A), "
                 f"and then to the outermost ISO 8 anteroom ({t_suite}). A positive air pressure system is maintained throughout "
                 f"the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}.")
        part2 = (f"The ISO 5 BSC E00{bsc_main}, located in the {t_loc} (Suite {t_suite}{t_suffix}), and ISO 5 BSC E00{bsc_chg}, "
                 f"located in the {c_loc} (Suite {c_suite}{c_suffix}), were thoroughly cleaned and disinfected prior to their "
                 f"respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). "
                 f"Furthermore, the BSCs used throughout testing, E00{bsc_main} for sample processing and E00{bsc_chg} for the "
                 f"changeover step, were certified and approved by both the Engineering and Quality Assurance teams.")
        if analyst_main == analyst_chg:
            usage = (f"{subj} processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{bsc_main}) "
                     f"and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{bsc_chg}) "
                     f"by {analyst_main} on {date_val}.")
        else:
            usage = (f"{subj} processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{bsc_main}) "
                     f"by {analyst_main} and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{bsc_chg}) "
                     f"by {analyst_chg} on {date_val}.")
        return f"{part1}\n\n{part2} {usage}"

    # Scenario 3: Diff BSC, Diff Suite
    else:
        p1a = (f"The cleanroom used for testing (E00{t_room}) consists of three interconnected sections of cleanroom suite {t_suite}: "
               f"the innermost ISO 7 cleanroom ({t_suite}B), which opens into the middle ISO 7 buffer room ({t_suite}A), and then into the outermost ISO 8 anteroom ({t_suite}). "
               f"A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {t_suite}B through {t_suite}A and into {t_suite}.")
        p1b = (f"The cleanroom used for changeover (E00{c_room}) consists of three interconnected sections of cleanroom suite {c_suite}: "
               f"the innermost ISO 7 cleanroom ({c_suite}B), which opens into the middle ISO 7 buffer room ({c_suite}A), and then into the outermost ISO 8 anteroom ({c_suite}). "
               f"A positive air pressure system is maintained throughout the suite to ensure controlled, unidirectional airflow from {c_suite}B through {c_suite}A and into {c_suite}.")
        part1 = f"{p1a}\n\n{p1b}"
        part2 = (f"The ISO 5 BSC E00{bsc_main}, located in the {t_loc} (Suite {t_suite}{t_suffix}), and ISO 5 BSC E00{bsc_chg}, "
                 f"located in the {c_loc} (Suite {c_suite}{c_suffix}), were thoroughly cleaned and disinfected prior to their "
                 f"respective procedures in accordance with SOP 2.600.018 (Cleaning and Disinfecting Procedure for Microbiology). "
                 f"Furthermore, the BSCs used throughout testing, E00{bsc_main} for sample processing and E00{bsc_chg} for the "
                 f"changeover step, were certified and approved by both the Engineering and Quality Assurance teams.")
        if analyst_main == analyst_chg:
            usage = (f"{subj} processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{bsc_main}) "
                     f"and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{bsc_chg}) "
                     f"by {analyst_main} on {date_val}.")
        else:
            usage = (f"{subj} processing was conducted within the ISO 5 BSC in the innermost section of the cleanroom (Suite {t_suite}{t_suffix}, BSC E00{bsc_main}) "
                     f"by {analyst_main} and the changeover step was conducted within the ISO 5 BSC in the middle section of the cleanroom (Suite {c_suite}{c_suffix}, BSC E00{bsc_chg}) "
                     f"by {analyst_chg} on {date_val}.")
        return f"{part1}\n\n{part2} {usage}"

def generate_history_text():
    if st.session_state.incidence_count == 0 or st.session_state.has_prior_failures == "No": phrase = "no prior failures"
    else:
        pids = [st.session_state.get(f"prior_oos_{i}","").strip() for i in range(st.session_state.incidence_count)]
        phrase = f"{len(pids)} incidents ({', '.join(pids)})"
    return f"Analyzing a 6-month sample history for {st.session_state.client_name}, this specific analyte \"{st.session_state.sample_name}\" has had {phrase} using the Scan RDI method during this period."

def generate_cross_contam_text():
    if st.session_state.other_positives == "No": return "All other samples processed by the analyst and other analysts that day tested negative. These findings suggest that cross-contamination between samples is highly unlikely."
    return f"{st.session_state.total_pos_count_num} samples tested positive. {st.session_state.sample_id} was the {ordinal(st.session_state.current_pos_order)} sample. Cross-contamination is unlikely."

# --- PDF GENERATOR (Layer 3) ---
def generate_filled_pdf(data_dict, template_path, output_path, mapping_config):
    from pypdf import PdfWriter
    try:
        writer = PdfWriter(clone_from=template_path)
        
        # Apply Mapping logic
        pdf_map = {}
        for pdf_key, data_key in mapping_config.items():
            pdf_map[pdf_key] = data_dict.get(data_key, "")

        # Execute filling
        writer.update_page_form_field_values(writer.pages[0], pdf_map)
        
        with open(output_path, "wb") as f:
            writer.write(f)
        return True, None
    except Exception as e:
        return False, str(e)

# --- CORE LOGIC: DOCUMENT GENERATOR ---
def generate_documents():
    from docxtpl import DocxTemplate
    
    # DETERMINE PHASE & FILENAMES
    is_p2 = (st.session_state.include_phase2 == "Yes")
    
    if is_p2:
        file_prefix = "ScanRDI OOS P2"  # Uses P2 files
        main_template = f"{file_prefix} template 0.docx"
        helper_template = f"{file_prefix} template.docx"
        pdf_template = f"{file_prefix} template.pdf"
    else:
        file_prefix = "ScanRDI OOS P1"  # Uses P1 files
        main_template = f"{file_prefix} template 0.docx"
        helper_template = f"{file_prefix} template.docx"
        pdf_template = f"{file_prefix} template.pdf"

    # 1. Prepare Phase 1 Data (Layer 1) - Common for both P1 and P2 reports
    fresh_equip = generate_equipment_text(st.session_state.bsc_id, st.session_state.chgbsc_id, st.session_state.analyst_name, st.session_state.changeover_name, st.session_state.test_date, is_retest=False)
    t_room, t_suite, t_suffix, t_loc = get_room_logic(st.session_state.bsc_id)
    
    data = {k: v for k, v in st.session_state.items()}
    data.update({
        "equipment_summary": fresh_equip,
        "sample_history_paragraph": generate_history_text(),
        "cross_contamination_summary": generate_cross_contam_text(),
        "organism_morphology": f"{st.session_state.org_choice} {st.session_state.manual_org}".strip().title(),
        "cr_id": t_room, "cr_suit": t_suite, "suit": t_suffix, "bsc_location": t_loc,
        "changeover_bsc": st.session_state.chgbsc_id,
        "whole_P1": st.session_state.get("phase1_full_text", "See Phase 1 Report"), 
        "notes": "None"
    })
    
    # 2. Prepare Phase 2 Data (Only if P2)
    mapping_config = {} # Config for PDF mapping

    if is_p2:
        # A. Calculation & Logic
        retest_equip_sum = generate_equipment_text(
            st.session_state.retest_bsc_id,
            st.session_state.retest_chgbsc_id,
            st.session_state.retest_analyst_name,
            st.session_state.retest_changeover_name,
            st.session_state.retest_date,
            is_retest=True
        )
        
        r_room, r_suite, r_suffix, r_loc = get_room_logic(st.session_state.retest_bsc_id)
        rc_room, rc_suite, rc_suffix, rc_loc = get_room_logic(st.session_state.retest_chgbsc_id)
        
        # B. Smart Variable Assembly (Layer 2)
        smart_pers = f"Prepper: {st.session_state.retest_prepper_name} ({st.session_state.retest_prepper_initial})\nProcessors: {st.session_state.retest_analyst_name} ({st.session_state.retest_analyst_initial})\nReader: {st.session_state.retest_reader_name} ({st.session_state.retest_reader_initial})"
        smart_ids = f"Original Test: {st.session_state.sample_id}\n\nRetest: {st.session_state.retest_sample_id}"
        smart_orig_res = f"{st.session_state.sample_id} - Fail"
        smart_retest_res = f"{st.session_state.retest_sample_id} - {st.session_state.retest_result}"
        smart_bsc_list = f"{st.session_state.retest_bsc_id} and {st.session_state.retest_chgbsc_id}"
        smart_suite_list = f"Suite {r_suite}{r_suffix}, Suite {rc_suite}{rc_suffix}"
        smart_p1_block = f"INITIAL TEST UNDER {st.session_state.sample_id}\n\n{data['whole_P1']}"
        
        smart_p2_narrative = (
            f"RETEST UNDER SUBMISSION {st.session_state.retest_sample_id}\n\n"
            f"Analogous to original testing, the analysts involved in prepping, processing and reading the retest samples under "
            f"{st.session_state.retest_sample_id}, {st.session_state.retest_prepper_name}, {st.session_state.retest_analyst_name} and {st.session_state.retest_reader_name} "
            f"confirmed no deviations from standard procedures.\n\n"
            f"The retest sample was stored upon arrival according to the Client’s instructions. Analysts {st.session_state.retest_prepper_name} and {st.session_state.retest_analyst_name} "
            f"confirmed the integrity of the samples throughout both the preparation and processing stages. No leaks or turbidity were observed at any point, verifying that the samples remained intact.\n\n"
            f"All reagents and supplies mentioned in the material section above were stored according to the suppliers’ recommendations, and their integrity was visually verified before utilization. "
            f"Moreover, each reagent and supply had valid expiration dates.\n\n"
            f"During the preparation phase, {st.session_state.retest_prepper_name} disinfected the samples using acidified bleach and placed them into a pre-disinfected storage bin. "
            f"On {st.session_state.retest_date}, prior to sample processing, {st.session_state.retest_analyst_name} performed a second disinfection with acidified bleach, "
            f"allowing a minimum contact time of 10 minutes before transferring the samples into the cleanroom suites.\n\n"
            f"A final disinfection step was completed immediately before the samples were introduced into the ISO 5 Biological Safety Cabinet (BSC), E00{st.session_state.retest_bsc_id}, "
            f"located within the {r_loc}, (Suite {r_suite}{r_suffix}), All activities were performed in accordance with SOP 2.600.023, Rapid Scan RDI® Test Using FIFU Method.\n\n"
            f"{retest_equip_sum}\n\n"
            f"The analyst, {st.session_state.retest_reader_name}, confirmed that the Scan RDI equipment E00{st.session_state.retest_scan_id} was set up as per SOP 2.700.004 "
            f"(Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup and Maintenance), and the negative control and the positive control for the analyst, "
            f"{st.session_state.retest_reader_name}, yielded expected results.\n\n"
            f"On {st.session_state.retest_date}, a rapid sterility test was conducted on the retest sample using the ScanRDI method. The retest sample was initially prepared by Analyst "
            f"{st.session_state.retest_prepper_name}, processed by {st.session_state.retest_analyst_name} and subsequently read by {st.session_state.retest_reader_name}. "
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

        data.update({
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
        
        # Define Mapping for P2 PDF (Based on our investigation)
        mapping_config = {
            "Text Field0": "sample_name",
            "Text Field1": "smart_retest_personnel_block",
            "Text Field2": "smart_sample_id_block",
            "Text Field3": "smart_retest_result_str",
            "Text Field4": "smart_original_result_str",
            "Text Field30": "oos_id",
            "Date Field0": "retest_date",
            "Text Field8": "smart_retest_scan_id",
            "Text Field9": "smart_retest_bsc_list",
            "Text Field10": "smart_retest_suite_list",
            "Text Field22": "smart_phase1_summary_block",
            "Text Field23": "smart_phase2_narrative_block"
        }

    # 3. Generate Layer 1 Doc (Main Report)
    st.session_state.docx_buf = None
    if os.path.exists(main_template):
        try:
            doc0 = DocxTemplate(main_template)
            doc0.render(data)
            buf0 = io.BytesIO(); doc0.save(buf0); buf0.seek(0)
            st.session_state.docx_buf = buf0
        except Exception as e: st.error(f"Main Report Generation Failed: {e}")
    else: st.error(f"⚠️ Template Missing: {main_template}")
    
    # 4. Generate Layer 2 Doc (Helper Fields)
    st.session_state.p2_buf = None
    if os.path.exists(helper_template):
        try:
            docP2 = DocxTemplate(helper_template)
            docP2.render(data)
            bufP2 = io.BytesIO(); docP2.save(bufP2); bufP2.seek(0)
            st.session_state.p2_buf = bufP2
        except Exception as e: st.error(f"Helper Doc Generation Failed: {e}")

    # 5. Generate Layer 3 (Final PDF) - Only if P2 and mapping exists
    st.session_state.pdf_buf = None
    if is_p2 and os.path.exists(pdf_template) and mapping_config:
        success, msg = generate_filled_pdf(data, pdf_template, "temp_filled.pdf", mapping_config)
        if success:
            with open("temp_filled.pdf", "rb") as f:
                st.session_state.pdf_buf = io.BytesIO(f.read())
        else:
            st.error(f"PDF Generation Failed: {msg}")
    elif is_p2 and not os.path.exists(pdf_template):
         st.warning(f"⚠️ PDF Template Missing: {pdf_template}")

# --- 5. INITIALIZE STATE ---
if "data_loaded" not in st.session_state:
    load_saved_state(); st.session_state.data_loaded = True
for k in field_keys: init_state(k, 1 if "count" in k else "N/A" if "etx" in k else "")
if "report_generated" not in st.session_state: st.session_state.report_generated = False

# --- 6. UI ---
st.title("🦠 ScanRDI Investigation")
st.header("1. General Test Details")
c1, c2, c3 = st.columns(3)
with c1: st.text_input("OOS Number", key="oos_id"); st.text_input("Client Name", key="client_name"); st.text_input("Sample ID", key="sample_id")
with c2: st.text_input("Test Date (07Jan26)", key="test_date"); st.text_input("Sample Name", key="sample_name"); st.text_input("Lot Number", key="lot_number")
with c3: st.selectbox("Dosage Form", ["Injectable","Solution"], key="dosage_form"); st.text_input("Monthly Cleaning Date", key="monthly_cleaning_date")

st.header("2. Personnel")
p1, p2 = st.columns(2)
with p1: st.text_input("Prepper Name", key="prepper_name")
with p2: st.text_input("Processor Name", key="analyst_name")
st.radio("Different Reader?", ["No","Yes"], key="diff_reader_analyst", horizontal=True)
if st.session_state.diff_reader_analyst == "Yes": st.text_input("Reader Name", key="reader_name")
else: st.session_state.reader_name = st.session_state.analyst_name
st.radio("Different Changeover?", ["No","Yes"], key="diff_changeover_analyst", horizontal=True)
if st.session_state.diff_changeover_analyst == "Yes": st.text_input("Changeover Name", key="changeover_name")
else: st.session_state.changeover_name = st.session_state.analyst_name

st.divider()
e1, e2 = st.columns(2)
bsc_list = ["1310","1309","1311","1312","1314","1313","Other"]
with e1: st.selectbox("Processing BSC ID", bsc_list, key="bsc_id")
with e2: 
    st.radio("Different Changeover BSC?", ["No","Yes"], key="diff_changeover_bsc", horizontal=True)
    if st.session_state.diff_changeover_bsc == "Yes": st.selectbox("Changeover BSC ID", bsc_list, key="chgbsc_id")
    else: st.session_state.chgbsc_id = st.session_state.bsc_id

st.header("3. Findings & EM")
f1, f2 = st.columns(2)
with f1: st.text_input("Confirmed #", key="confirm_number"); st.selectbox("Org Shape", ["rod","cocci","Other"], key="org_choice")
with f2: st.text_input("Control Lot", key="control_lot"); st.radio("EM Growth?", ["No","Yes"], key="em_growth_observed", horizontal=True)

st.divider()
# --- PHASE 2 SECTION ---
st.subheader("🚦 Phase 2 Investigation")
st.checkbox("Include Phase 2 Investigation?", key="include_phase2")

if st.session_state.include_phase2:
    st.info("Enter details for the Retest below:")
    r1, r2, r3 = st.columns(3)
    with r1: st.text_input("Retest Date (DDMMMYY)", key="retest_date"); st.text_input("Retest Sample ID", key="retest_sample_id")
    with r2: st.text_input("Retest Result", value="Pass", key="retest_result"); st.text_input("Retest Scan ID (e.g. 2017)", key="retest_scan_id")
    
    st.markdown("##### Retest Personnel")
    p2_1, p2_2 = st.columns(2)
    with p2_1: st.text_input("Retest Prepper", key="retest_prepper_name"); st.text_input("Initials", key="retest_prepper_initial")
    with p2_2: st.text_input("Retest Processor", key="retest_analyst_name"); st.text_input("Initials", key="retest_analyst_initial")
    
    st.radio("Different Retest Reader?", ["No","Yes"], key="diff_retest_reader", horizontal=True)
    if st.session_state.diff_retest_reader == "Yes": st.text_input("Retest Reader", key="retest_reader_name"); st.text_input("Reader Initials", key="retest_reader_initial")
    else: st.session_state.retest_reader_name = st.session_state.retest_analyst_name; st.session_state.retest_reader_initial = st.session_state.retest_analyst_initial
    
    st.radio("Different Retest Changeover?", ["No","Yes"], key="diff_retest_changeover", horizontal=True)
    if st.session_state.diff_retest_changeover == "Yes": st.text_input("Retest Changeover Name", key="retest_changeover_name"); st.text_input("Chg Initials", key="retest_changeover_initial")
    else: st.session_state.retest_changeover_name = st.session_state.retest_analyst_name; st.session_state.retest_changeover_initial = st.session_state.retest_analyst_initial

    st.markdown("##### Retest Equipment")
    e2_1, e2_2 = st.columns(2)
    with e2_1: st.selectbox("Retest Proc. BSC", bsc_list, key="retest_bsc_id")
    with e2_2: 
        st.radio("Diff Retest Chg BSC?", ["No","Yes"], key="diff_retest_bsc", horizontal=True)
        if st.session_state.diff_retest_bsc == "Yes": st.selectbox("Retest Chg BSC", bsc_list, key="retest_chgbsc_id")
        else: st.session_state.retest_chgbsc_id = st.session_state.retest_bsc_id

# --- GENERATE ---
if st.button("🚀 GENERATE REPORT"):
    generate_documents()
    st.session_state.report_generated = True

if st.session_state.report_generated:
    st.success("Reports Ready!")
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.session_state.get("docx_buf"):
            # Dynamic Label based on Phase
            label = "P2 Main Report" if st.session_state.include_phase2 == "Yes" else "P1 Main Report"
            st.download_button(f"📄 {label} (Word)", st.session_state.docx_buf, f"Report_Main.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with c2:
        if st.session_state.get("p2_buf"):
            label = "P2 Helper Doc" if st.session_state.include_phase2 == "Yes" else "P1 Helper Doc"
            st.download_button(f"📂 {label} (Word)", st.session_state.p2_buf, f"Helper_Fields.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    with c3:
        if st.session_state.get("pdf_buf"):
            st.download_button("✅ Final P2 Form (PDF)", st.session_state.pdf_buf, f"Final_Form.pdf", "application/pdf")
