# filename: app.py
import streamlit as st

# --- 1. 全局配置 (必须是脚本运行的第一条命令) ---
st.set_page_config(
    page_title="Eagle Analytical OOS Tools",
    page_icon="🦅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. 侧边栏视觉与路由提示 ---
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1a/Blank_square.svg/1200px-Blank_square.svg.png", width=50) # 你可以换成你们 Eagle 的 Logo URL
st.sidebar.title("🔬 OOS Modules")
st.sidebar.markdown("Select a module above to begin.")
st.sidebar.markdown("---")
st.sidebar.caption("System Architect: Qiyue Chen")

# --- 3. 大堂主视觉与系统说明 (Landing Page) ---
st.title("🦅 Eagle Analytical OOS Generator")

st.markdown("""
### Welcome to the Automated OOS Investigation System

This system is designed to seamlessly generate Phase I OOS investigation reports (Form 3.100.019.F01) by combining user inputs with robust environmental monitoring logic.

#### Available Modules:

1.  **🦠 ScanRDI Investigation**
    * **Methodology**: Rapid Scan RDI® Test (FIFU Method).
    * **SOP Reference**: 2.600.023 / 2.700.004
    * **Key Metrics**: Morphological count and differential staining.

2.  **🧪 Celsis Investigation**
    * **Methodology**: Celsis Sterility Testing.
    * **SOP Reference**: 2.600.059
    * **Key Metrics**: Relative Luminescence Units (RLU) and %CV.

---
**Directions for Use**: 
Please expand the sidebar on the left and click on the specific module you need to run today's investigation. Ensure your Word and PDF templates are correctly placed in the root directory.
""")

st.info("💡 **Tip**: The system uses dynamic date calculation. Please ensure all dates entered follow the strict DDMMMYY format (e.g., 17Mar26) unless otherwise specified.")
