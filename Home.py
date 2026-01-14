import streamlit as st

# MUST BE THE FIRST LINE
st.set_page_config(page_title="Microbiology Platform", layout="wide")

# --- ULTIMATE PROFESSIONAL CSS ---
st.markdown("""
    <style>
    /* 1. Solid Dark Blue Sidebar */
    [data-testid="stSidebar"] {
        background-color: #003366 !important;
    }
    
    /* 2. Remove default navigation */
    [data-testid="stSidebarNav"] {
        display: none;
    }

    /* 3. Style for the Extra Large 'Micro' Expander (Mind-Map Style) */
    .st-emotion-cache-p5mtransition {
        background-color: transparent !important;
        border: none !important;
    }
    
    summary {
        color: white !important;
        font-size: 40px !important; /* 2 sizes larger than before */
        font-weight: 800 !important;
        padding-left: 5px !important;
        list-style: none !important;
    }

    /* 4. Sub-items styling (White & Professional) */
    div[data-testid="stPageLink"] p {
        color: white !important;
        font-size: 18px !important;
        font-weight: 400 !important;
        margin-left: 20px !important;
    }
    
    /* Ensure the expander arrow is white */
    .st-emotion-cache-p5mtransition svg {
        fill: white !important;
    }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    # Branding header
    st.markdown("<h2 style='color: #66CC33; padding-left:10px; margin-bottom:0;'>EAGLE</h2>", unsafe_allow_html=True)
    st.markdown("---")

    # FOLDED MIND-MAP: Micro
    with st.expander("Micro", expanded=False):
        # NO Dashboard here - strictly the 4 tests
        st.page_link("pages/ScanRDI.py", label="ScanRDI")
        st.page_link("pages/USP_71.py", label="USP <71>")
        st.page_link("pages/Celsis.py", label="Celsis")
        st.page_link("pages/EM.py", label="EM")

# --- MAIN CONTENT ---
st.title("Microbiology Investigation Platform")
st.write("Please click the extra-large **Micro** header in the sidebar to select an investigation type.")
