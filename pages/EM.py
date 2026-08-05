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

# --- 3. INPUT FORM ---
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
    st.text_input("Reader Analyst Name(s)", key="reader_name", placeholder="e.g. Maraya Chukwumerije & Simin Mohammad")
    st.text_input("Action / Alert Level", key="action_level", value="Action Level: ≥ 1 CFU/Plate")
    st.text_input("CFU Count & Organism", key="manual_org", placeholder="e.g. 10 CFUs - Staphylococcus epidermidis")

st.markdown("---")
st.markdown("### 📝 Phase I Narrative Preview")

if st.button("🔄 Generate Narrative Preview"):
    errors, warnings = el.validate_inputs()
    if errors:
        for err in errors: st.error(err)
    else:
        if warnings:
            st.warning(f"⚠️ Missing recommended fields: {', '.join(warnings)}")
        
        interview_block, records_block, summary_block = el.generate_em_narrative()
        
        st.success("✅ EM Phase I Narrative Generated Successfully!")
        
        st.subheader("1. Analyst Interview & Storage Narrative")
        st.info(interview_block)
        
        st.subheader("2. Environmental Monitoring Summary Narrative")
        st.info(records_block)
        
        st.subheader("3. Defensive Phase I Summary")
        st.success(summary_block)
