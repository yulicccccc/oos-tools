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
        font-size: 40px !important; /* Extra Large */
        font-weight: 800 !important;
        padding-left: 5px !important;
        list-style: none !important;
    }
    
    /* Change 'Micro' to Green when hovered */
    summary:hover {
        color: #66CC33 !important;
        cursor: pointer;
    }

    /* 4. Sub-items styling (White by default) */
    div[data-testid="stPageLink"] p {
        color: white !important;
        font-size: 18px !important;
        font-weight: 400 !important;
        margin-left: 20px !important;
        transition: 0.3s;
    }
    
    /* 5. GREEN SELECTION/HOVER EFFECT */
    /* Changes the text to green when you hover over a link */
    div[data-testid="stPageLink"]:hover p {
        color: #66CC33 !important;
    }
    
    /* Changes the background to a subtle green tint or just keeps the text green */
    div[data-testid="stPageLink"] button:hover {
        background-color: rgba(102, 204, 51, 0.1) !important;
        border: 1px solid #66CC33 !important;
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

    # FOLDED MIND-MAP: Micro (Default is folded)
    with st.expander("Micro", expanded=False):
        # All links turn GREEN when hovered
        st.page_link("pages/ScanRDI.py", label="ScanRDI")
        st.page_link("pages/USP_71.py", label="USP <71>")
        st.page_link("pages/Celsis.py", label="Celsis")
        st.page_link("pages/EM.py", label="EM")

# --- MAIN CONTENT ---
st.title("Microbiology Investigation Platform")
st.write("Please click the extra-large **Micro** header in the sidebar to select an investigation type.")
