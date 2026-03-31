import streamlit as st

# 1. 大门招牌 (Page Configuration)
# 使用了你更专业的标题 "Microbiology Platform"
st.set_page_config(
    page_title="Microbiology Platform", 
    page_icon="🦅",
    layout="wide"
)

# 2. 呼叫后勤部刷漆 (Apply the Standard Eagle Sidebar)
try:
    from utils import apply_eagle_style
    apply_eagle_style()
except ImportError:
    pass

# 3. 迎宾主屏幕 (Professional Main Content)
st.title("🦅 Microbiology Investigation Platform")
st.markdown("---")

st.markdown("""
### Welcome to the Quality Control Portal

This platform is designed to streamline OOS investigations and data trending for the Microbiology department.
All investigation modules are synchronized with the **Eagle Trax** standards.

**Available Modules:**
* **ScanRDI**: Rapid Sterility Testing Investigation
* **USP <71>**: Traditional Sterility Testing
* **Celsis**: Rapid Microbial Detection
* **EM**: Environmental Monitoring Data Entry

*Please select a module from the sidebar on the left to begin your work.*
""")

# 4. 底部状态栏 (Clean info box at the bottom)
st.markdown("<br><br>", unsafe_allow_html=True) # 增加一点留白，让页面更呼吸
st.info("🟢 System Status: Online | Version: 1.0.2 | Maintained by LabOps")
