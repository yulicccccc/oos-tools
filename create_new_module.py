# filename: create_new_module.py
"""
OOS Module SKU Generator Workflow
----------------------------------
Automates the creation of a new OOS test module in seconds.
Usage:
    python create_new_module.py --name USP85 --sop 2.600.085 --desc "Endotoxin Testing"
"""

import sys
import os
import shutil
import argparse
import docxtpl

def create_module(module_name, sop_ref="2.600.XXX", description="New Test Module"):
    module_upper = module_name.upper()
    module_lower = module_name.lower()
    module_title = module_name.capitalize()
    
    print(f"\n=======================================================")
    print(f"🚀 Initializing OOS SKU Workflow for Module: {module_upper}")
    print(f"=======================================================\n")
    
    # 1. Create Templates
    base_docx = "EM OOS P1 template.docx"
    base_docx_0 = "EM OOS P1 template 0.docx"
    base_pdf = "EM OOS P1 template.pdf"
    
    target_docx = f"{module_upper} OOS P1 template.docx"
    target_docx_0 = f"{module_upper} OOS P1 template 0.docx"
    target_pdf = f"{module_upper} OOS P1 template.pdf"
    
    for src, dst in [(base_docx, target_docx), (base_docx_0, target_docx_0), (base_pdf, target_pdf)]:
        if os.path.exists(src):
            shutil.copyfile(src, dst)
            print(f"  [1/4] Created template: {dst}")
        else:
            print(f"  [1/4] Warning: Source template {src} not found.")
            
    # 2. Extract Variable Contract
    if os.path.exists(target_docx):
        tpl = docxtpl.DocxTemplate(target_docx)
        vars_list = sorted(list(tpl.get_undeclared_template_variables()))
        print(f"  [2/4] Extracted {len(vars_list)} variables for data contract.")
        with open(f"scratch/{module_lower}_variable_contract.txt", "w", encoding="utf-8") as f:
            f.write(f"{module_upper} Variable Contract ({len(vars_list)} variables):\n")
            for v in vars_list:
                f.write(f"  - {v}\n")
    
    # 3. Create Logic Engine
    logic_filename = f"{module_lower}_logic.py"
    logic_code = f'''# filename: {logic_filename}
import streamlit as st
import re
from datetime import datetime

try:
    from utils import get_room_logic, get_full_name
except ImportError:
    def get_room_logic(i): return "Unknown", "000", "", "Unknown"
    def get_full_name(i): return i

FIELD_KEYS = [
    "oos_id", "client_name", "sample_id", "test_date", "sample_name", "lot_number", 
    "analyst_name", "reader_name", "bsc_id", "event_number", "confirm_number"
]

def validate_inputs():
    errors, warnings = [], []
    reqs = {{"OOS Number": "oos_id", "Sample Name": "sample_name", "Test Date": "test_date"}}
    for label, key in reqs.items():
        if not st.session_state.get(key, "").strip(): warnings.append(label)
    return errors, warnings

def generate_narrative():
    s = st.session_state
    sample_name = s.get("sample_name", "[Sample Name]")
    analyst_name = s.get("analyst_name", "[Analyst Name]")
    
    summary = f"Phase I investigation completed for {{sample_name}} by {{analyst_name}} under SOP {sop_ref}."
    return summary
'''
    with open(logic_filename, "w", encoding="utf-8") as f:
        f.write(logic_code)
    print(f"  [3/4] Generated logic engine: {logic_filename}")

    # 4. Create Streamlit Page
    page_filename = f"pages/{module_upper}.py"
    page_code = f'''# filename: {page_filename}
import streamlit as st
from utils import apply_eagle_style
import {module_lower}_logic as logic

st.set_page_config(page_title="{module_upper} Investigation", layout="wide")
apply_eagle_style()

st.title("🔬 {module_upper} OOS Investigation")
st.caption("Form 3.100.019.F01 - SOP {sop_ref} ({description})")

st.markdown("### 📋 Section A: Test Details")
col1, col2 = st.columns(2)
with col1:
    st.text_input("OOS Number", key="oos_id")
    st.text_input("Sample Name", key="sample_name")
with col2:
    st.text_input("Test Date (DDMMMYY)", key="test_date")
    st.text_input("Analyst Name", key="analyst_name")

if st.button("🔄 Generate Report"):
    st.success("✅ Generated {module_upper} Phase I Narrative")
'''
    with open(page_filename, "w", encoding="utf-8") as f:
        f.write(page_code)
    print(f"  [4/4] Generated Streamlit UI page: {page_filename}")

    print(f"\n=======================================================")
    print(f"✨ Module {module_upper} created successfully!")
    print(f"=======================================================\n")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="OOS SKU Module Generator")
    parser.add_argument("--name", required=True, help="Module Name (e.g. USP85)")
    parser.add_argument("--sop", default="2.600.XXX", help="SOP Reference")
    parser.add_argument("--desc", default="New OOS Test Module", help="Module Description")
    
    args = parser.parse_args()
    create_module(args.name, args.sop, args.desc)
