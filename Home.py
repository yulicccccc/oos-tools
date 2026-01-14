import streamlit as st
from utils import apply_eagle_style

# 1. Page Configuration
st.set_page_config(page_title="Microbiology Platform", layout="wide")

# 2. Apply the Standard Eagle Sidebar (Blue Background, Huge Micro, Green Hover)
apply_eagle_style()

# 3. Professional Main Content
st.title("Microbiology Investigation Platform")
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

*Please select a module from the sidebar to begin.*
""")

# Optional: Add a clean info box at the bottom
st.info("System Status: Online | Version: 1.0.2")
