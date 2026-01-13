import streamlit as st
from docxtpl import DocxTemplate
from pypdf import PdfWriter
import os, io
from datetime import datetime
from utils import get_full_name, get_room_logic, num_to_words

st.title("🦠 ScanRDI Sterility Investigation")

# --- UI SECTIONS ---
tab1, tab2, tab3 = st.tabs(["Details & Personnel", "EM Observations", "History & Cross-Contam"])

with tab1:
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("OOS Number", key="oos_id")
        st.text_input("Client Name", key="client_name")
        st.text_input("Sample ID", key="sample_id")
        st.text_input("Test Date", key="test_date")
    with c2:
        st.text_input("Prepper Init", key="prepper_initial")
        st.text_input("Processor Init", key="analyst_initial")
        st.selectbox("BSC ID", ["1310", "1309", "1311", "1312", "1314", "1313"], key="bsc_id")

with tab2:
    st.radio("EM Growth Observed?", ["No", "Yes"], key="em_growth_observed", horizontal=True)
    if st.session_state.em_growth_observed == "Yes":
        st.number_input("Count", min_value=1, key="em_growth_count")
        for i in range(st.session_state.get('em_growth_count', 1)):
            st.text_input(f"Obs #{i+1} (e.g. 1 CFU...)", key=f"em_obs_{i}")

with tab3:
    st.radio("Prior Failures?", ["No", "Yes"], key="has_prior_failures", horizontal=True)

# --- REPORT GENERATION ---
if st.button("🚀 GENERATE SMART REPORTS"):
    cr_id, cr_suit, suit, bsc_loc = get_room_logic(st.session_state.bsc_id)
    p_name = get_full_name(st.session_state.prepper_initial)
    a_name = get_full_name(st.session_state.analyst_initial)
    r_name = get_full_name(st.session_state.get('reader_initial', st.session_state.analyst_initial))
    
    # Smart Grammar Logic
    unique_names = list(dict.fromkeys([n for n in [p_name, a_name, r_name] if n]))
    analyst_str = " and ".join([", ".join(unique_names[:-1]), unique_names[-1]]) if len(unique_names) > 1 else unique_names[0]

    # BUILD SMART VARIABLES
    smart_personnel = f"Prepper: {p_name}\nProcessor: {a_name}\nReader: {r_name}"
    smart_phase1_summary = f"All analysts involved – {analyst_str} – were interviewed. Integrity confirmed. Processing in BSC E00{st.session_state.bsc_id}."

    # MAPPING
    pdf_data = {
        'Text Field8': st.session_state.oos_id,
        'Text Field3': smart_personnel,
        'Text Field49': smart_phase1_summary,
        'Text Field13': f"Yes, analysts {analyst_str} were interviewed comprehensively."
    }

    # PDF Logic
    if os.path.exists("templates/ScanRDI_template.pdf"):
        writer = PdfWriter(clone_from="templates/ScanRDI_template.pdf")
        for page in writer.pages:
            writer.update_page_form_field_values(page, pdf_data)
        buf = io.BytesIO()
        writer.write(buf)
        st.download_button("📂 Download PDF", buf.getvalue(), f"OOS-{st.session_state.oos_id}.pdf")
