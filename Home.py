import streamlit as st

# MUST BE THE FIRST LINE
st.set_page_config(page_title="Microbiology Platform", layout="wide")

# --- CUSTOM CSS: EAGLE TRAX BLUE & SMART FONT SCALING ---
st.markdown("""
    <style>
    /* 1. Force Sidebar Background to Eagle Blue */
    [data-testid="stSidebar"] {
        background-color: #003366 !important;
    }
    
    /* 2. Hide the default Streamlit page list */
    [data-testid="stSidebarNav"] {
        display: none;
    }

    /* 3. Style for the BIG 'Micro' Header (2 sizes larger) */
    .st-emotion-cache-p5mtransition {
        background-color: transparent !important;
        border: none !important;
    }
    
    summary {
        color: white !important;
        font-size: 34px !important; /* Extra large font */
        font-weight: 800 !important;
        padding-left: 5px !important;
        list-style: none !important;
    }

    /* 4. Style for the unfolded sub-items */
    div[data-testid="stPageLink"] p {
        color: white !important;
        font-size: 16px !important;
        font-weight: 400 !important;
        margin-left: 15px !important;
    }
    
    /* Clean up the expander arrow color */
    .st-emotion-cache-p5mtransition svg {
        fill: white !important;
    }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    # Official Branding
    st.markdown("<h2 style='color: #66CC33; padding-left:10px; margin-bottom:0;'>EAGLE</h2>", unsafe_allow_html=True)
    st.markdown("---")

    # THE EXPANDABLE "MICRO" CATEGORY
    # expanded=False ensures it is "Folded" by default like a mind map
    with st.expander("Micro", expanded=False):
        st.page_link("Home.py", label="Dashboard")
        st.page_link("pages/ScanRDI.py", label="ScanRDI")
        st.page_link("pages/USP_71.py", label="USP <71>")
        st.page_link("pages/Celsis.py", label="Celsis")
        st.page_link("pages/Environmental_Monitoring.py", label="Environmental Monitoring")

# --- MAIN CONTENT ---
st.title("Microbiology Dashboard")
st.write("Please click the large **Micro** header in the sidebar to reveal investigation types.")
