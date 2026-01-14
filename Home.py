import streamlit as st

st.set_page_config(page_title="Microbiology Platform", layout="wide")

# --- PROFESSIONAL CSS: BLUE SIDEBAR & WHITE TEXT ---
st.markdown("""
    <style>
    /* Force Sidebar Background to Blue */
    [data-testid="stSidebar"] {
        background-color: #003366 !important;
    }
    
    /* Hide default Streamlit navigation */
    [data-testid="stSidebarNav"] {
        display: none;
    }

    /* BIG FONT for Category Header */
    .category-header {
        color: white !important;
        font-size: 26px !important;
        font-weight: bold !important;
        padding: 10px 0px 5px 15px;
    }

    /* SMALLER FONT for Sub-items */
    div[data-testid="stPageLink"] p {
        color: white !important;
        font-size: 16px !important;
        padding-left: 10px;
    }
    
    /* Highlight the Micro link to look like a main header */
    .micro-link p {
        font-size: 26px !important;
        font-weight: bold !important;
    }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    # Eagle Trax Branding
    st.markdown("<h2 style='color: #66CC33; padding-left:15px;'>EAGLE</h2>", unsafe_allow_html=True)
    st.markdown("---")

    # PRIMARY CATEGORY: Micro (Links back to Home)
    st.page_link("Home.py", label="Micro")
    
    # SUB-ITEMS: OOS Types & Environmental Monitoring
    st.page_link("pages/ScanRDI.py", label="ScanRDI")
    st.page_link("pages/USP_71.py", label="USP <71>")
    st.page_link("pages/Celsis.py", label="Celsis")
    st.page_link("pages/Environmental_Monitoring.py", label="Environmental Monitoring")

# --- PAGE CONTENT ---
st.title("Microbiology Dashboard")
st.write("Please select a specific investigation type from the menu.")
