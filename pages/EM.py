# filename: pages/EM.py
import streamlit as st
import os
import re
import json
import io
from datetime import datetime

# --- 1. SAFE UTILS & LOGIC IMPORT ---
try:
    from utils import apply_eagle_style, get_room_logic, get_full_name
    import em_logic as el
except ImportError as e:
    st.error(f"Import Error: {e}")
    def apply_eagle_style(): pass
    def get_room_logic(i): return "Unknown", "000", "", "Unknown"
    def get_full_name(i): return i

# --- 2. PAGE CONFIG & STYLING ---
st.set_page_config(page_title="EM OOS Investigation", layout="wide")
apply_eagle_style()

st.title("🧫 Environmental Monitoring (EM) OOS Investigation")
st.caption("Form 3.100.019.F01 (Rev 11) - SOP 2.600.002 Standard Automated Report & Table Generator")

# --- 3. SMART EMAIL & DOCX TABLE IMPORT ---
st.markdown("### 📥 Smart Import: EM Table (.docx) or Email Text")
u_col1, u_col2 = st.columns([1, 1])

with u_col1:
    uploaded_table = st.file_uploader("📂 Upload EM Summary Table (.docx):", type=["docx"])
    if uploaded_table is not None:
        if st.button("🪄 Parse Uploaded .docx Table"):
            parsed = el.parse_em_docx_table(uploaded_table)
            if parsed:
                for k, v in parsed.items():
                    st.session_state[k] = v
                st.session_state["em_show_reports"] = False
                st.success("✨ Successfully parsed EM Table .docx! All fields auto-filled.")
                st.rerun()

with u_col2:
    email_text = st.text_area(
        "📧 Paste EM Notification Email / Text:", 
        height=100, 
        placeholder="Paste notification text containing OOS-260361, ETX-260216-0348, Plate Name, CFU Count, etc."
    )
    if st.button("🪄 Parse Pasted Text"):
        parsed = el.parse_em_text(email_text)
        if parsed:
            for k, v in parsed.items():
                st.session_state[k] = v
            st.session_state["em_show_reports"] = False
            st.success("✨ Auto-filled parsed fields! Please review below.")
            st.rerun()

st.markdown("---")

# --- 4. INPUT FORM ---
st.markdown("### 📋 Section A: Test & Environmental Details")

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.text_input("OOS Number", key="oos_id", placeholder="e.g. OOS-260361")
    st.text_input("Setup Analyst Name", key="analyst_name", placeholder="e.g. Simin Mohammad")
with c2:
    st.text_input("Sample / Plate Name", key="sample_name", placeholder="e.g. EM SMO 116A Air 07MAY2026")
    st.text_input("Setup Analyst Initial", key="analyst_initial", placeholder="e.g. SMO")
with c3:
    st.text_input("Event / ETX Number", key="event_number", placeholder="e.g. ETX-260518-0254")
    st.selectbox("Plate Media Type", ["TSA Plate", "Contact Plate"], key="plate_media_type")
with c4:
    st.text_input("Test Date (DDMMMYY)", key="test_date", placeholder="e.g. 07May26")
    st.selectbox(
        "Sampling Type", 
        ["Settling Sampling", "Surface Sampling", "Surface Sampling (Changeover)", "Weekly Cleanroom Sampling", "Personnel Sampling (Glove)"], 
        key="sampling_type"
    )

st.markdown("##### 🔬 Media / Reagents & Investigation Results")
r1, r2, r3, r4 = st.columns(4)
with r1:
    st.text_input("Media Plate Lot #", key="media_plate_lot", value=st.session_state.get("media_plate_lot", "1011543730"))
    st.text_input("Equipment / BSC ID", key="bsc_id", placeholder="e.g. BSC 1314 or CR116")
with r2:
    st.text_input("Media Plate Expiration", key="media_plate_exp", value=st.session_state.get("media_plate_exp", "29SEP2026"))
    st.text_input("Reader Analyst Name(s)", key="reader_name", value=st.session_state.get("reader_name", "Maraya Chukwumerije & Simin Mohammad"))
with r3:
    st.text_input("CFU Count", key="cfu_count", placeholder="e.g. 748")
    st.text_input("Action / Alert Level", key="action_level", value=st.session_state.get("action_level", "Action Level: ≥ 1 CFU/Plate"))
with r4:
    st.text_input("Organism(s) Identified", key="manual_org", placeholder="e.g. Kocuria palustris (Gram (+) cocci)...")

st.markdown("---")
st.markdown("### 📝 Phase I Narrative & Report Generation")

if st.button("🚀 Generate Reports & Documents (Word & 7-Page PDF)"):
    st.session_state["em_show_reports"] = True

if st.session_state.get("em_show_reports", False):
    errors, warnings = el.validate_inputs()
    if errors:
        for err in errors: st.error(err)
        st.session_state["em_show_reports"] = False
    else:
        if warnings:
            st.warning(f"⚠️ Missing recommended fields: {', '.join(warnings)}")
        
        interview_block, records_block, summary_block = el.generate_em_narrative()
        docx_buf, pdf_buf = el.generate_em_reports()
        
        st.success("✅ EM Phase I Complete 7-Page Report Generated Successfully!")
        
        st.markdown("### 📂 Download Reports & Attachments")
        c1, c2, c3 = st.columns(3)
        safe_name = el.clean_filename(st.session_state.get("oos_id", "EM_Report"))
        
        with c1:
            st.subheader("Word Document")
            if docx_buf:
                st.download_button(
                    "📄 EM OOS Full Report (.docx)", 
                    docx_buf, 
                    f"{safe_name}.docx", 
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.error("Word template not found or rendering failed.")

        with c2:
            st.subheader("7-Page PDF Report")
            if pdf_buf:
                st.download_button(
                    "🔴 EM OOS Complete 7-Page PDF (.pdf)", 
                    pdf_buf, 
                    f"{safe_name}.pdf", 
                    "application/pdf"
                )
            else:
                st.error("PDF template not found or rendering failed.")

        with c3:
            st.subheader("Backup Session")
            session_data = {k: st.session_state[k] for k in el.FIELD_KEYS if k in st.session_state}
            st.download_button(
                "💾 Save Session Data (.txt)", 
                json.dumps(session_data, indent=2), 
                f"SAVE_{safe_name}.txt", 
                "text/plain"
            )
            
        st.markdown("---")
        st.subheader("1. Analyst Interview & Storage Narrative (Part 1)")
        st.info(interview_block)
        
        st.subheader("2. Environmental Monitoring Summary Narrative (Part 2)")
        st.info(records_block)
        
        st.subheader("3. Defensive Phase I Summary & Conclusion (Part 3)")
        st.success(summary_block)
