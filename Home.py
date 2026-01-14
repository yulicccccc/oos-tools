import streamlit as st
from utils import apply_eagle_style

# 1. 必须是第一行
st.set_page_config(page_title="Microbiology Platform", layout="wide")

# 2. 调用 utils 里的统一函数 (这就是“变身”的关键)
# 这行代码会强制把侧边栏变成深蓝色，把 Micro 变成超大号，把子选项变成绿色高亮
apply_eagle_style()

# 3. 页面内容 (作为 Landing Page，保持简单)
st.title("Microbiology Investigation Platform")
st.info("⬅️ Please verify the sidebar style is correct, then select a test from the **MICRO** menu.")
