import streamlit as st

# ==========================================
# 1. UI 样式控制 (UI Styling)
# ==========================================
def apply_eagle_style():
    """统一设置所有页面的全局 UI 样式，保持 Eagle Analytical 的品牌一致性"""
    st.markdown("""
        <style>
        /* 隐藏 Streamlit 默认的右上角菜单和底部标志 */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        /* 全局按钮的圆角和颜色 */
        .stButton>button { 
            border-radius: 5px; 
            font-weight: bold;
        }
        </style>
    """, unsafe_allow_html=True)


# ==========================================
# 2. 核心数据字典 (Shared Data Dictionaries)
# ==========================================
def get_full_name(initial):
    """
    (核心部件) 缩写转全名翻译器。
    后续实验室如果有新人入职，只需在此处添加一行即可，所有子页面自动生效。
    """
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


# ==========================================
# 3. 物理逻辑与算法 (Business & Physics Logic)
# ==========================================
def get_room_logic(bsc_id):
    """根据 BSC 的单双号推算所在的 Cleanroom 和 Suite"""
    try:
        num = int(bsc_id)
        # 偶数为 B，奇数为 A
        suffix = "B" if num % 2 == 0 else "A"
        location = "innermost ISO 7 room" if suffix == "B" else "middle ISO 7 buffer room"
    except: 
        suffix, location = "B", "innermost ISO 7 room"
    
    # 机器对应的 Suite 组
    if bsc_id in ["1310", "1309"]: suite = "117"
    elif bsc_id in ["1311", "1312"]: suite = "116"
    elif bsc_id in ["1314", "1313"]: suite = "115"
    elif bsc_id in ["1316", "1798"]: suite = "114"
    else: suite = "Unknown"
    
    room_map = {"117": "1739", "116": "1738", "115": "1737", "114": "1736"}
    return room_map.get(suite, "Unknown"), suite, suffix, location


# ==========================================
# 4. 格式化小工具 (Formatting Helpers)
# ==========================================
def ordinal(n):
    """将阿拉伯数字转换为英语序数词 (例如: 1 -> 1st, 2 -> 2nd)"""
    try: 
        n = int(n) 
    except ValueError: 
        return str(n)
        
    if 11 <= (n % 100) <= 13: 
        suffix = 'th'
    else: 
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"

def num_to_words(n):
    """将 1-10 的数字转换为英文单词 (例如: 1 -> one)"""
    mapping = {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 
               6: "six", 7: "seven", 8: "eight", 9: "nine", 10: "ten"}
    return mapping.get(n, str(n))
