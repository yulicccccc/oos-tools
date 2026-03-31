import streamlit as st

# 必须放在第一行：设置浏览器标签页的名字、图标和宽屏模式
st.set_page_config(
    page_title="Eagle OOS Tools",
    page_icon="🦅",
    layout="wide"
)

# 导入我们昨天刚写好的公共 UI 样式
from utils import apply_eagle_style

# ---------------------------------------------------------
# 1. 渲染侧边栏导航 (调用 utils.py 里的功能)
# ---------------------------------------------------------
apply_eagle_style()

# ---------------------------------------------------------
# 2. 迎宾主页面 (Lobby / Landing Page)
# ---------------------------------------------------------
st.title("🦅 Welcome to Eagle Analytical OOS Tools")
st.markdown("---")

st.markdown("""
### 👋 Hello there! 
Welcome to the automated investigation report generator. 

👈 **Please select a module from the sidebar on the left to begin your work.**

#### Available Modules:
* **ScanRDI**: For Rapid Sterility Testing OOS investigations.
* **USP <71>**: For traditional sterility testing investigations.
* **Celsis**: For rapid microbial detection investigations.
* **EM**: For Environmental Monitoring excursions.

---
*If you need to update the analyst name list or room logic, please modify the `utils.py` file.*
""")
