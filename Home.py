import streamlit as st

# MUST BE THE FIRST LINE
st.set_page_config(page_title="Microbiology Platform", layout="wide")

# --- THE ULTIMATE EAGLE CSS ---
st.markdown("""
    <style>
    /* 1. Dark Blue Sidebar */
    [data-testid="stSidebar"] {
        background-color: #003366 !important;
    }
    
    /* 2. Hide default nav list */
    [data-testid="stSidebarNav"] {
        display: none;
    }

    /* 3. MASSIVE MICRO HEADER (2 sizes larger + Bold) */
    .st-emotion-cache-p5mtransition p {
        font-size: 48px !important; 
        font-weight: 900 !important; 
        color: white !important;
        text-transform: uppercase;
        letter-spacing: 2px;
    }

    /* 4. Sub-items: ALL BOLD & Professional Font Size */
    div[data-testid="stPageLink"] p {
        color: white !important;
        font-size: 22px !important;
        font-weight: 800 !important;
        margin-left: 20px !important;
        transition: 0.2s ease-in-out;
    }

    /* 5. THE GREEN SELECTION/HOVER FIX */
    /* Forces Eagle Green (#66CC33) when selected or hovered */
    div[data-testid="stPageLink"] a:hover p,
    div[data-testid="stPageLink"] a:focus p,
    div[data-testid="stPageLink"] a:active p {
        color: #66CC33 !important;
    }

    /* Green highlight background box */
    div[data-testid="stPageLink"] a:hover,
    div[data-testid="stPageLink"] a:focus {
        background-color: rgba(102, 204, 51, 0.15) !important;
        border-radius: 10px;
        text-decoration: none !important;
    }

    /* Expand/Fold Arrow Styling */
    summary svg {
        fill: white !important;
        transform: scale(1.8);
    }
    
    .st-emotion-cache-p5mtransition {
        background-color: transparent !important;
        border: none !important;
    }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    # Branding
    st.markdown("<h1 style='color: #66CC33; padding-left:10px; font-weight:900;'>EAGLE</h1>", unsafe_allow_html=True)
    st.markdown("---")

    # MICRO: Mind-Map Folded Category
    with st.expander("Micro", expanded=True):
        st.page_link("pages/ScanRDI.py", label="ScanRDI")
        st.page_link("pages/USP_71.py", label="USP <71>")
        st.page_link("pages/Celsis.py", label="Celsis")
        st.page_link("pages/EM.py", label="EM")
