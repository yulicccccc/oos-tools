import streamlit as st
from utils import get_full_name, get_room_logic

st.set_page_config(page_title="ScanRDI Investigation", layout="wide")

# Sidebar Header consistency
st.sidebar.markdown("<h2 style='color: #66CC33;'>ScanRDI</h2>", unsafe_allow_html=True)

st.title("ScanRDI Sterility Investigation Report")
st.markdown("---")

# Section 1: General Details
with st.expander("GENERAL TEST DETAILS", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        st.text_input("OOS ID", key="oos_id")
    with col2:
        st.text_input("Test Date", key="test_date")
    with col3:
        st.text_input("Sample ID", key="sample_id")

# Section 2: Environmental Monitoring Results
with st.container():
    st.subheader("ENVIRONMENTAL MONITORING DATA")
    st.radio("Was microbial growth observed in EM?", ["No", "Yes"], key="em_growth_observed", horizontal=True)
    # Dynamic rows logic remains here...
