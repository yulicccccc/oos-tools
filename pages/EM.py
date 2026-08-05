# filename: pages/EM.py
import streamlit as st
import os
import re
import json
import io
import sys
import subprocess
import time
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
st.caption("Form 3.100.019.F01 - SOP 2.600.002 Automated Report Generator")

# --- 3. SMART EMAIL IMPORT ---
st.markdown("### 📧 Smart Email Import / Paste Text")
email_text = st.text_area("Paste EM Notification Email or OOS Results text here:", height=130, placeholder="Paste email content containing OOS-261187, Plate Name, CFU Count, Organisms, etc.")

if st.button("🪄 Parse & Auto-Fill Form"):
    parsed = el.parse_em_text(email_text)
    if parsed:
        for k, v in parsed.items():
            st.session_state[k] = v
        st.success("✨ Auto-filled parsed fields! Please review below.")
        st.rerun()
    else:
        st.warning("⚠️ No matching fields found in pasted text. Please enter details manually.")

st.markdown("---")
# --- 4. INPUT FORM ---
st.markdown("### 📋 Section A: Test & Environmental Details")

col1, col2, col3 = st.columns(3)

with col1:
    st.text_input("OOS Number", key="oos_id", placeholder="e.g. OOS-260422")
    st.text_input("Sample / Plate Name", key="sample_name", placeholder="e.g. ScanC/O HS BSC1309 S3 17FEB2026")
    st.text_input("Test Date (DDMMMYY)", key="test_date", placeholder="e.g. 17Feb26")

with col2:
    st.selectbox("Sampling Type", ["Surface Sampling", "Settling Sampling", "Personnel Sampling (Glove)", "Weekly Cleanroom Sampling"], key="sampling_type")
    st.text_input("Equipment / BSC ID", key="bsc_id", placeholder="e.g. BSC 1309 or CR114")
    st.text_input("Setup Analyst Name", key="analyst_name", placeholder="e.g. Simin Mohammad")

with col3:
    reader_options = [
        "Simin Mohammad",
        "Maraya Chukwumerije"
    ]
    current_reader = st.session_state.get("reader_name", reader_options[0])
    idx = reader_options.index(current_reader) if current_reader in reader_options else 0
    selected_reader = st.selectbox("Reader Analyst Name(s)", reader_options, index=idx, key="reader_select")
    st.session_state["reader_name"] = selected_reader

    st.text_input("Action / Alert Level", key="action_level", value="Action Level: ≥ 1 CFU/Plate")
    st.text_input("CFU Count & Organism", key="manual_org", placeholder="e.g. 10 CFUs - Staphylococcus epidermidis")

st.markdown("---")
st.markdown("### 📝 Phase I Narrative Preview")

if st.button("🚀 Generate Reports & Documents (Word & PDF)"):
    errors, warnings = el.validate_inputs()
    if errors:
        for err in errors: st.error(err)
    else:
        if warnings:
            st.warning(f"⚠️ Missing recommended fields: {', '.join(warnings)}")
        
        interview_block, records_block, summary_block = el.generate_em_narrative()
        docx_buf, pdf_buf = el.generate_em_reports()
        
        st.success("✅ EM Phase I Report Generated Successfully!")
        
        st.markdown("### 📂 Download Reports")
        c1, c2, c3 = st.columns(3)
        safe_name = el.clean_filename(st.session_state.get("oos_id", "EM_Report"))
        
        with c1:
            st.subheader("Word Document")
            if docx_buf:
                st.download_button(
                    "📄 EM OOS Report (.docx)", 
                    docx_buf, 
                    f"{safe_name}.docx", 
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.error("Word template not found or rendering failed.")

        with c2:
            st.subheader("PDF Form Document")
            if pdf_buf:
                st.download_button(
                    "🔴 EM OOS Report (.pdf)", 
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
        st.subheader("1. Analyst Interview & Storage Narrative")
        st.info(interview_block)
        
        st.subheader("2. Environmental Monitoring Summary Narrative")
        st.info(records_block)
        
        st.subheader("3. Defensive Phase I Summary")
        st.success(summary_block)
