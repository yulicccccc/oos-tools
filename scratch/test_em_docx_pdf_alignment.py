import os
import sys
import io
import docxtpl
import pypdf

sys.path.insert(0, os.getcwd())
sys.stdout.reconfigure(encoding='utf-8')

# Dummy session state simulation
class DummySession(dict):
    def __getattr__(self, key):
        return self.get(key, "")

import streamlit as st

st.session_state = {
    "oos_id": "OOS-261187",
    "sample_name": "ScanC/O CGS E001309 S1 11MAY2026",
    "event_number": "ETX-260518-0273",
    "test_date": "11May26",
    "analyst_name": "Clea S. Garza",
    "analyst_initial": "CGS",
    "reader_name": "Simin Mohammad",
    "bsc_id": "BSC 1309",
    "action_level": "Action Level: ≥ 1 CFU/Plate",
    "manual_org": "White, shiny, hardened, smooth area embeded on agar"
}

import em_logic as el

interview_block, records_block, summary_block = el.generate_em_narrative()

context = {
    "oos_id": st.session_state["oos_id"],
    "sample_id": st.session_state["event_number"],
    "sample_name": st.session_state["sample_name"],
    "lot_number": st.session_state["sample_name"],
    "test_date": st.session_state["test_date"],
    "analyst_name": st.session_state["analyst_name"],
    "analyst_initial": st.session_state["analyst_initial"],
    "reader_name": st.session_state["reader_name"],
    "bsc_id": st.session_state["bsc_id"],
    "dosage_form": "Plate",
    "smart_comment_interview": interview_block,
    "smart_comment_records": records_block,
    "smart_phase1_summary": summary_block,
    "smart_phase1_part1": interview_block,
    "smart_phase1_part2": summary_block
}

# Test DOCX rendering
doc = docxtpl.DocxTemplate("EM OOS P1 template 0.docx")
doc.render(context)
doc.save("scratch/test_output_em.docx")
print("Saved test_output_em.docx")

# Test PDF filling
pdf_map = {
    'Text Field57': st.session_state["oos_id"],
    'Text Field0': f"{st.session_state['analyst_name']} (written by: Qiyue Chen)",
    'Date Field0': st.session_state["test_date"],
    'Date Field1': st.session_state["test_date"],
    'Date Field2': st.session_state["test_date"],
    'Text Field1': "Environmental Monitoring",
    'Text Field2': st.session_state["event_number"],
    'Text Field3': f"Setup Analyst:\n{st.session_state['analyst_name']} ({st.session_state['analyst_initial']})\n\nReader Analyst:\n{st.session_state['reader_name']}",
    'Text Field4': st.session_state["sample_name"],
    'Text Field5': "Plate",
    'Text Field6': st.session_state["sample_name"],
    'Text Field7': "The CFU count for the environmental monitoring plate exceeded the action level.",
    'Text Field8': "2.600.002",
    'Text Field11': st.session_state["action_level"],
    'Text Field15': "Yes, as per SOP 2.600.002",
    'Text Field16': "Yes, as per SOP 2.600.002",
    'Text Field21': "Yes, as per SOP 2.600.002",
    'Text Field49': interview_block,
    'Text Field50': records_block,
    'Text Field51': summary_block
}

writer = pypdf.PdfWriter(clone_from="EM OOS P1 template.pdf")
for page in writer.pages:
    writer.update_page_form_field_values(page, pdf_map)
with open("scratch/test_output_em.pdf", "wb") as f:
    writer.write(f)
print("Saved test_output_em.pdf")

# Inspect PDF fields in output
reader = pypdf.PdfReader("scratch/test_output_em.pdf")
fields = reader.get_fields()
print("\nInspecting Key Fields in Generated PDF:")
print(f"  - Initiator (Text Field0): {repr(fields['Text Field0'].get('/V'))}")
print(f"  - Sample Name (Text Field4): {repr(fields['Text Field4'].get('/V'))}")
print(f"  - Lot # (Text Field6): {repr(fields['Text Field6'].get('/V'))}")
