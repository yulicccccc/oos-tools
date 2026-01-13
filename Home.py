import streamlit as st
import json, os
from utils import parse_email_text

st.set_page_config(page_title="LabOps Master Tool", layout="wide")

st.title("🧪 LabOps Sterility Investigation Tool")
st.markdown("---")

st.header("📧 1. Smart Email Import")
st.info("Paste your OOS Notification email here to pre-fill the ScanRDI, Celsis, or USP 71 pages.")

email_input = st.text_area("Email Content", height=200)

if st.button("🪄 Parse & Distribute Data"):
    if email_input:
        data = parse_email_text(email_input)
        for k, v in data.items():
            st.session_state[k] = v
        st.success("✅ Data parsed! You can now navigate to the platform pages in the sidebar.")
    else:
        st.warning("Please paste an email first.")

st.markdown("---")
st.subheader("Instructions")
st.write("1. Paste email and click 'Parse'.")
st.write("2. Select your platform (e.g., ScanRDI) from the sidebar.")
st.write("3. Review auto-filled info and generate your reports.")
