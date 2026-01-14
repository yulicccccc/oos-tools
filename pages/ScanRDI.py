import streamlit as st
from utils import apply_eagle_style, get_full_name, get_room_logic

# 1. 必须是第一行
st.set_page_config(page_title="ScanRDI", layout="wide")

# 2. 立即应用 Eagle 统一侧边栏
apply_eagle_style()

# 3. 页面内容
st.title("🦠 ScanRDI Investigation")

tab1, tab2 = st.tabs(["📋 General Details", "🔍 EM Observations"])

with tab1:
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("OOS Number", key="oos_id")
        st.text_input("Sample ID", key="sample_id")
    with c2:
        st.selectbox("BSC ID", ["1310", "1309", "1311", "1312"], key="bsc_id")
        st.text_input("Analyst Initials", key="analyst_int", on_change=lambda: st.write(f"Name: {get_full_name(st.session_state.analyst_int)}"))

with tab2:
    st.write("Environmental Monitoring data goes here...")
    st.button("Generate ScanRDI Report")
