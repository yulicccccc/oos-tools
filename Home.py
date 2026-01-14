import streamlit as st

# Hide the default sidebar navigation so our custom one is the only one visible
st.set_page_config(page_title="Microbiology Platform", layout="wide")

# --- CUSTOM CSS FOR EAGLETRAX LOOK ---
st.markdown("""
    <style>
    [data-testid="stSidebarNav"] {display: none;} /* Hides default nav */
    .big-nav {
        font-size: 1.5rem !important;
        font-weight: 700 !important;
        color: #ffffff !important;
        margin-top: 20px !important;
        margin-bottom: 5px !important;
    }
    .sub-nav {
        font-size: 0.9rem !important;
        color: #66CC33 !important; /* Green from Eagle Trax */
        padding-left: 15px !important;
        text-decoration: none !important;
    }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    # Top Branding Header
    st.markdown(
        """
        <div style="background-color: #003366; padding: 15px; border-radius: 5px; margin-bottom: 25px;">
            <h2 style="color: white; margin: 0;">EAGLETRAX</h2>
            <p style="color: #66CC33; margin: 0; font-weight: bold;">MICROBIOLOGY</p>
        </div>
        """, unsafe_allow_html=True
    )

    # BIG HEADER
    st.markdown('<p class="big-nav">Home</p>', unsafe_allow_html=True)
    st.page_link("Home.py", label="Dashboard Summary")
    
    # BIG HEADER
    st.markdown('<p class="big-nav">Investigations</p>', unsafe_allow_html=True)
    # Smaller font sub-items
    st.page_link("pages/ScanRDI.py", label="ScanRDI")
    st.page_link("pages/USP_71.py", label="USP <71>")
    st.page_link("pages/Celsis.py", label="Celsis")
    
    # BIG HEADER
    st.markdown('<p class="big-nav">Environmental</p>', unsafe_allow_html=True)
    st.page_link("pages/Environmental_Monitoring.py", label="EM Portal")
