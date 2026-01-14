import streamlit as st
import re
from datetime import datetime

# --- 1. 统一的界面样式函数 ---
def apply_eagle_style():
    """
    在每个页面调用此函数，即可获得完全一致的 Eagle Trax 侧边栏。
    """
    # 强制 CSS 样式
    st.markdown("""
        <style>
        /* 1. 侧边栏背景：深蓝 */
        [data-testid="stSidebar"] {
            background-color: #003366 !important;
        }
        
        /* 2. 隐藏 Streamlit 自带导航 */
        [data-testid="stSidebarNav"] {
            display: none;
        }

        /* 3. Micro 标题：超大 (48px)、超粗 (900)、大写 */
        .st-emotion-cache-p5mtransition p {
            font-size: 48px !important;
            font-weight: 900 !important;
            color: white !important;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 0px !important;
        }

        /* 4. 子选项：粗体 (22px)、白色 */
        div[data-testid="stPageLink"] p {
            color: white !important;
            font-size: 22px !important;
            font-weight: 700 !important;
            margin-left: 15px !important;
            transition: all 0.2s ease;
        }

        /* 5. 选中/悬停时的绿色高亮 (EAGLE GREEN #66CC33) */
        div[data-testid="stPageLink"] a:hover p,
        div[data-testid="stPageLink"] a:focus p,
        div[data-testid="stPageLink"] a:active p {
            color: #66CC33 !important;
        }
        
        /* 选中时的背景微光 */
        div[data-testid="stPageLink"] a:hover,
        div[data-testid="stPageLink"] a:focus {
            background-color: rgba(102, 204, 51, 0.15) !important;
            border-radius: 8px;
            text-decoration: none !important;
        }

        /* 修正折叠箭头的颜色和大小 */
        summary svg {
            fill: white !important;
            transform: scale(1.5) !important;
        }
        .st-emotion-cache-p5mtransition {
            background-color: transparent !important;
            border: none !important;
        }
        </style>
    """, unsafe_allow_html=True)

    # 统一的侧边栏结构
    with st.sidebar:
        st.markdown("<h1 style='color: #66CC33; padding-left:10px; font-weight:900; margin-bottom:0;'>EAGLE</h1>", unsafe_allow_html=True)
        st.markdown("---")
        
        # 默认展开 Micro
        with st.expander("Micro", expanded=True):
            st.page_link("pages/ScanRDI.py", label="ScanRDI")
            st.page_link("pages/USP_71.py", label="USP <71>")
            st.page_link("pages/Celsis.py", label="Celsis")
            st.page_link("pages/EM.py", label="EM")

# --- 2. 业务逻辑工具函数 ---

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

def parse_email_text(text):
    data = {}
    oos = re.search(r"OOS-(\d+)", text)
    if oos: data['oos_id'] = oos.group(1)
    client = re.search(r"([A-Za-z\s]+\(E\d+\))", text)
    if client: data['client_name'] = client.group(1).strip()
    etx = re.search(r"(ETX-\d{6}-\d{4})", text)
    if etx: data['sample_id'] = etx.group(1).strip()
    return data
