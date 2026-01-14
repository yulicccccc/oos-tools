import streamlit as st

st.set_page_config(page_title="Microbiology Investigation Tool", layout="wide")

# --- CUSTOM CSS FOR EAGLE TRAX SIDEBAR ---
st.markdown("""
    <style>
    /* 1. Make the entire sidebar Dark Blue */
    [data-testid="stSidebar"] {
        background-color: #003366 !important;
    }
    
    /* 2. Hide the default navigation */
    [data-testid="stSidebarNav"] {
        display: none;
    }

    /* 3. Style for the BIG 'Micro' header */
    .micro-header {
        color: white !important;
        font-size: 24px !important;
        font-weight: bold !important;
        margin-bottom: 10px !important;
        padding-left: 10px;
    }

    /* 4. Style for the sub-options */
    div[data-testid="stPageLink"] p {
        color: white !important;
        font-size: 16px !important;
        font-weight: 400 !important;
    }
    
    /* 5. Add some spacing to the sidebar top */
    .st-emotion-cache-16umgzp {
        padding-top: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    # Top Logo/Branding
    st.markdown("<h2 style='color: #66CC33; padding-left:10px;'>EAGLE</h2>", unsafe_allow_html=True)
    st.markdown("---")

    # BIG CATEGORY: Micro
    st.markdown('<div class="micro-header">Micro</div>', unsafe_allow_html=True)
    
    # SUB-OPTIONS: Smaller font links
    st.page_link("Home.py", label="Dashboard")
    st.page_link("pages/ScanRDI.py", label="ScanRDI")
    st.page_link("pages/USP_71.py", label="USP <71>")
    st.page_link("pages/Celsis.py", label="Celsis")
    st.page_link("pages/Environmental_Monitoring.py", label="EM Portal")

# --- MAIN CONTENT ---
st.title("Microbiology Sterility Investigation Platform")
