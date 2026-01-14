import streamlit as st

# MUST BE THE FIRST LINE
st.set_page_config(page_title="Microbiology Platform", layout="wide")

# --- ULTIMATE EAGLE TRAX CSS ---
st.markdown("""
    <style>
    /* 1. Solid Dark Blue Sidebar Background */
    [data-testid="stSidebar"] {
        background-color: #003366 !important;
    }
    
    /* 2. Hide the default Streamlit sidebar menu */
    [data-testid="stSidebarNav"] {
        display: none;
    }

    /* 3. Style for the BOLD & EXTRA LARGE 'Micro' Expander */
    .st-emotion-cache-p5mtransition {
        background-color: transparent !important;
        border: none !important;
    }
    
    summary {
        color: white !important;
        font-size: 45px !important; /* Extra Large Font */
        font-weight: 900 !important; /* Extra Bold */
        padding-left: 5px !important;
        list-style: none !important;
    }
    
    summary:hover {
        color: #66CC33 !important; /* Turns Green on Hover */
    }

    /* 4. Sub-items styling (Bold & Professional) */
    div[data-testid="stPageLink"] p {
        color: white !important;
        font-size: 20px !important;
        font-weight: 700 !important; /* All Bold as requested */
        margin-left: 20px !important;
        transition: 0.3s;
    }
    
    /* 5. THE GREEN SELECTION FIX */
    /* This targets the button container when hovered or active */
    div[data-testid="stPageLink"] > a:hover, 
    div[data-testid="stPageLink"] > a:focus {
        background-color: rgba(102, 204, 51, 0.2) !important;
        border-radius: 5px;
    }

    /* Turns the text green when hovered */
    div[data-testid="stPageLink"] > a:hover p,
    div[data-testid="stPageLink"] > a:focus p {
        color: #66CC33 !important;
    }

    /* Ensure the expander arrow icon is white */
    .st-emotion-cache-p5mtransition svg {
        fill: white !important;
    }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    # Official Branding
    st.markdown("<h1 style='color: #66CC33; padding-left:10px; font-weight:900;'>EAGLE</h1>", unsafe_allow_html=True)
    st.markdown("---")

    # THE FOLDED MIND-MAP: Micro (starts folded)
    with st.expander("Micro", expanded=False):
        # All items are now Bold and turn Green on hover/selection
        st.page_link("pages/ScanRDI.py", label="ScanRDI")
        st.page_link("pages/USP_71.py", label="USP <71>")
        st.page_link("pages/Celsis.py", label="Celsis")
        st.page_link("pages/EM.py", label="EM")

# --- PAGE CONTENT ---
st.title("Microbiology Dashboard")
st.write("Click the extra-large **Micro** header to reveal investigation options.")
