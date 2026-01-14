import re
from datetime import datetime

import streamlit as st

def apply_eagle_sidebar():
    st.markdown("""
        <style>
        /* 1. Solid Dark Blue Sidebar */
        [data-testid="stSidebar"] {
            background-color: #003366 !important;
        }
        
        /* 2. Hide default nav */
        [data-testid="stSidebarNav"] {
            display: none;
        }

        /* 3. MASSIVE & BOLD 'Micro' Header */
        .st-emotion-cache-p5mtransition p {
            font-size: 48px !important; 
            font-weight: 900 !important; 
            color: white !important;
            text-transform: uppercase;
        }

        /* 4. Sub-items: ALL BOLD */
        div[data-testid="stPageLink"] p {
            color: white !important;
            font-size: 22px !important;
            font-weight: 800 !important;
            margin-left: 20px !important;
        }

        /* 5. GREEN SELECTION/HOVER FIX */
        div[data-testid="stPageLink"] a:hover p,
        div[data-testid="stPageLink"] a:focus p {
            color: #66CC33 !important;
        }

        div[data-testid="stPageLink"] a:hover {
            background-color: rgba(102, 204, 51, 0.15) !important;
            border-radius: 10px;
        }

        summary svg {
            fill: white !important;
            transform: scale(1.8);
        }
        </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("<h1 style='color: #66CC33; padding-left:10px; font-weight:900;'>EAGLE</h1>", unsafe_allow_html=True)
        st.markdown("---")
        with st.expander("Micro", expanded=True):
            st.page_link("pages/ScanRDI.py", label="ScanRDI")
            st.page_link("pages/USP_71.py", label="USP <71>")
            st.page_link("pages/Celsis.py", label="Celsis")
            st.page_link("pages/EM.py", label="EM")

def get_full_name(initials):
    if not initials: return ""
    lookup = {
        "HS": "Halaina Smith", "DS": "Devanshi Shah", "GS": "Gabbie Surber",
        "MRB": "Muralidhar Bythatagari", "KSM": "Karla Silva", "DT": "Debrework Tassew",
        "PG": "Pagan Gary", "GA": "Gerald Anyangwe", "DH": "Domiasha Harrison",
        "TK": "Tamiru Kotisso", "AO": "Ayomide Odugbesi", "CCD": "Cuong Du",
        "ES": "Alex Saravia", "MJ": "Mukyang Jang", "KA": "Kathleen Aruta",
        "SMO": "Simin Mohammad", "VV": "Varsha Subramanian", "CSG": "Clea S. Garza",
        "GL": "Guanchen Li", "QYC": "Qiyue Chen"
    }
    return lookup.get(initials.upper().strip(), "")

def get_room_logic(bsc_id):
    try:
        num = int(bsc_id)
        suffix = "B" if num % 2 == 0 else "A"
        location = "innermost ISO 7 room" if suffix == "B" else "middle ISO 7 buffer room"
    except: suffix, location = "B", "innermost ISO 7 room"
    
    room_map = {"1310": "117", "1309": "117", "1311": "116", "1312": "116", "1314": "115", "1313": "115", "1316": "114", "1798": "114"}
    suite = room_map.get(bsc_id, "Unknown")
    room_id = {"117": "1739", "116": "1738", "115": "1737", "114": "1736"}.get(suite, "Unknown")
    return room_id, suite, suffix, location

def num_to_words(n):
    return {1: "one", 2: "two", 3: "three", 4: "four", 5: "five"}.get(n, str(n))

def parse_email_text(text):
    data = {}
    oos = re.search(r"OOS-(\d+)", text)
    if oos: data['oos_id'] = oos.group(1)
    client = re.search(r"([A-Za-z\s]+\(E\d+\))", text)
    if client: data['client_name'] = client.group(1).strip()
    etx = re.search(r"(ETX-\d{6}-\d{4})", text)
    if etx: data['sample_id'] = etx.group(1).strip()
    name = re.search(r"Sample\s*Name:\s*(.*)", text, re.IGNORECASE)
    if name: data['sample_name'] = name.group(1).strip()
    lot = re.search(r"(?:Lot|Batch)\s*[:\.]?\s*([^\n\r]+)", text, re.IGNORECASE)
    if lot: data['lot_number'] = lot.group(1).strip()
    date = re.search(r"testing\s*on\s*(\d{2}\s*\w{3}\s*\d{4})", text, re.IGNORECASE)
    if date:
        try:
            d_obj = datetime.strptime(date.group(1).strip(), "%d %b %Y")
            data['test_date'] = d_obj.strftime("%d%b%y")
        except: pass
    return data
