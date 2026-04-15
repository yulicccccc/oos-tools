# filename: utils.py
import streamlit as st
import re
from datetime import datetime, timedelta

# --- 1. 统一的界面样式函数 ---
def apply_eagle_style():
    """
    在每个页面调用此函数，即可获得完全一致的 Eagle Trax 侧边栏。
    """
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

    with st.sidebar:
        st.markdown("<h1 style='color: #66CC33; padding-left:10px; font-weight:900; margin-bottom:0;'>EAGLE</h1>", unsafe_allow_html=True)
        st.markdown("---")
        
        with st.expander("Micro", expanded=True):
            st.page_link("pages/ScanRDI.py", label="ScanRDI")
            st.page_link("pages/USP_71.py", label="USP <71>")
            st.page_link("pages/Celsis.py", label="Celsis")
            st.page_link("pages/EM.py", label="EM")

# --- 2. 业务逻辑工具函数 ---

def get_full_name(initial):
    """(终极版) 缩写转全名翻译器 - 已包含 Celsis 人员"""
    if not initial: 
        return ""
    mapping = {
        "KA": "Kathleen Aruta", "DH": "Domiasha Harrison", "GL": "Guanchen Li", "DS": "Devanshi Shah",
        "QC": "Qiyue Chen", "HS": "Halaina Smith", "MJ": "Mukyung Jang", "AS": "Alex Saravia",
        "CSG": "Clea S. Garza", "RS": "Robin Seymour", "CCD": "Cuong Du", "VV": "Varsha Subramanian",
        "KS": "Karla Silva", "GS": "Gabbie Surber", "PG": "Pagan Gary", "DT": "Debrework Tassew",
        "GA": "Gerald Anyangwe", "MRB": "Muralidhar Bythatagari", "TK": "Tamiru Kotisso", "OA": "Olugbenga Ajayi",
        "RE": "Rey Estrada", "AOD": "Ayomide Odugbesi", "EN": "Elysse Nioupin", "SU": "Sonal Uprety", 
        "AC": "Andrew Carrillo", "KC": "Kira C",
        "AA": "America Alanis",  # Celsis 主力
        "SMO": "SMO"             # EM 表格特定缩写
    }
    return mapping.get(initial.strip().upper(), "")

def get_room_logic(bsc_id):
    """单双号规则推断洁净室编号 + 1798 强制拦截"""
    bsc_str = str(bsc_id).strip()
    
    # 强制拦截 Celsis 专属设备 (1798 是偶数，但它在 A 房)
    if bsc_str == "1798":
        return "1736", "114", "A", "middle ISO 7 buffer room"
        
    try:
        num = int(bsc_str)
        suffix = "B" if num % 2 == 0 else "A"
        location = "innermost ISO 7 room" if suffix == "B" else "middle ISO 7 buffer room"
    except: 
        suffix, location = "B", "innermost ISO 7 room"
    
    if bsc_str in ["1310", "1309"]: suite = "117"
    elif bsc_str in ["1311", "1312"]: suite = "116"
    elif bsc_str in ["1314", "1313"]: suite = "115"
    elif bsc_str in ["1316"]: suite = "114"
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
    
    # 修复了贪婪匹配 Client 的 Bug
    client = re.search(r"^(?:.*\n)?(.*\bE\d{5}\b.*)$", text, re.MULTILINE)
    if client: 
        raw_client = client.group(1).strip()
        data['client_name'] = re.sub(r"^Client:\s*", "", raw_client, flags=re.IGNORECASE)
        
    etx = re.search(r"(ETX-\d{6}-\d{4})", text)
    if etx: data['sample_id'] = etx.group(1).strip()
    name = re.search(r"Sample\s*Name:\s*(.*)", text, re.IGNORECASE)
    if name: data['sample_name'] = name.group(1).strip()
    lot = re.search(r"(?:Lot|Batch)\s*[:\.]?\s*([^\n\r]+)", text, re.IGNORECASE)
    if lot: data['lot_number'] = lot.group(1).strip()
    date = re.search(r"testing\s*on\s*(\d{2}\s*\w{3}\s*\d{4})", text, re.IGNORECASE)
    if date:
        try:
            d_obj = datetime.strptime(date.group(1).strip().replace(" ", ""), "%d%b%Y")
            data['test_date'] = d_obj.strftime("%d%b%y")
        except: pass
    return data

# --- 3. 时间与日期高级计算工具 (Celsis 专属工作日引擎) ---

def get_business_day_back(start_date, days_to_back):
    """从起始日向前回溯指定的工作日数量，跳过周末"""
    current_date = start_date
    count = 0
    while count < days_to_back:
        current_date -= timedelta(days=1)
        if current_date.weekday() < 5:  # 0-4 是周一到周五
            count += 1
    return current_date

def get_celsis_dates(test_date_str):
    """
    根据发现阳性的 test_date，自动倒推接种日 (T-7 工作日) 和收样日 (T-8 工作日)。
    """
    try:
        fmt = "%d%b%y" if len(test_date_str) <= 7 else "%d%b%Y"
        t_anchor = datetime.strptime(test_date_str, fmt)
        
        process_dt = get_business_day_back(t_anchor, 7)
        received_dt = get_business_day_back(t_anchor, 8)
        
        return {
            "process_date": process_dt.strftime("%d%b%y"),
            "received_data": received_dt.strftime("%d%b%y")
        }
    except Exception:
        return {"process_date": "[Error]", "received_data": "[Error]"}
