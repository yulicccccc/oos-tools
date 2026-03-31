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

def get_full_name(initial):
    """(终极版) 缩写转全名翻译器"""
    if not initial: 
        return ""
    mapping = {
        "KA": "Kathleen Aruta", "DH": "Domiasha Harrison", "GL": "Guanchen Li", "DS": "Devanshi Shah",
        "QC": "Qiyue Chen", "HS": "Halaina Smith", "MJ": "Mukyung Jang", "AS": "Alex Saravia",
        "CSG": "Clea S. Garza", "RS": "Robin Seymour", "CCD": "Cuong Du", "VV": "Varsha Subramanian",
        "KS": "Karla Silva", "GS": "Gabbie Surber", "PG": "Pagan Gary", "DT": "Debrework Tassew",
        "GA": "Gerald Anyangwe", "MRB": "Muralidhar Bythatagari", "TK": "Tamiru Kotisso", "OA": "Olugbenga Ajayi",
        "RE": "Rey Estrada", "AOD": "Ayomide Odugbesi", "EN": "Elysse Nioupin", "SU": "Sonal Uprety", "AC": "Andrew Carrillo",
        "KC": "Kira C"
    }
    return mapping.get(initial.strip().upper(), "")

def get_room_logic(bsc_id):
    """单双号规则推断洁净室编号"""
    try:
        num = int(bsc_id)
        suffix = "B" if num % 2 == 0 else "A"
        location = "innermost ISO 7 room" if suffix == "B" else "middle ISO 7 buffer room"
    except: 
        suffix, location = "B", "innermost ISO 7 room"
    
    if bsc_id in ["1310", "1309"]: suite = "117"
    elif bsc_id in ["1311", "1312"]: suite = "116"
    elif bsc_id in ["1314", "1313"]: suite = "115"
    elif bsc_id in ["1316", "1798"]: suite = "114"
    else: suite = "Unknown"
    
    room_id = {"117": "1739", "116": "1738", "115": "1737", "114": "1736"}.get(suite, "Unknown")
    return room_id, suite, suffix, location

def num_to_words(n):
    """数字转单词"""
    return {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 
            6: "six", 7: "seven", 8: "eight", 9: "nine", 10: "ten"}.get(n, str(n))

def ordinal(n):
    """数字转序数词"""
    try: 
        n = int(n) 
    except ValueError: 
        return str(n)
    if 11 <= (n % 100) <= 13: 
        suffix = 'th'
    else: 
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"

def parse_email_text(text):
    """纯文本邮件解析工具，返回基础数据字典"""
    data = {}
    oos = re.search(r"OOS-(\d+)", text)
    if oos: data['oos_id'] = oos.group(1).strip()
    client = re.search(r"^(?:.*\n)?(.*\bE\d{5}\b.*)$", text, re.MULTILINE)
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
