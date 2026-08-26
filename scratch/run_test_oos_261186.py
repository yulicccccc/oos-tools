import os
import sys
import io
sys.path.insert(0, os.path.abspath("."))

import streamlit as st
import em_logic as el
import pypdf
import docxtpl

def run_test_case():
    print("="*70)
    print("RUNNING EM OOS TEST CASE: OOS-261186 (EM SMO 116A Air 07MAY2026)")
    print("="*70)

    # Populate Session State with actual data from user files
    st.session_state["oos_id"] = "OOS-261186"
    st.session_state["sample_name"] = "EM SMO 116A Air 07MAY2026"
    st.session_state["event_number"] = "ETX-260518-0254"
    st.session_state["test_date"] = "07May26"
    st.session_state["sampling_type"] = "Weekly Active Air Sampling"
    st.session_state["bsc_id"] = "Cleanroom Suite 116 (ISO 7 116A)"
    st.session_state["analyst_name"] = "Simin Mohammad"
    st.session_state["analyst_initial"] = "SMO"
    st.session_state["reader_name"] = "Samera A. Salim"
    st.session_state["action_level"] = "Alert Level: ≥ 8 CFU/Plate\nAction Level: > 10 CFU/Plate"
    st.session_state["cfu_count"] = "748"
    st.session_state["manual_org"] = "Kocuria palustris (Gram (+) cocci)"
    st.session_state["monthly_cleaning_date"] = "30 Apr 2026"
    st.session_state["cleaner_name"] = "Rey Estrada"
    st.session_state["writer_name"] = "Maryam Naeem"
    st.session_state["manager_name"] = "Kathan Parikh"
    st.session_state["test_method"] = "Sterility"

    # Generate Narratives
    p1, p2, p3 = el.generate_em_narrative()
    print("\n--- NARRATIVE PART 1 (Interview & Storage) ---")
    print(p1)
    print("\n--- NARRATIVE PART 2 (EM Summary & Bracketing) ---")
    print(p2)
    print("\n--- NARRATIVE PART 3 (Phase I Summary & Conclusion) ---")
    print(p3)

    # Generate Reports
    docx_buf, pdf_buf = el.generate_em_reports()
    
    out_docx = "scratch/OOS_261186_Report.docx"
    out_pdf = "scratch/OOS_261186_Report.pdf"

    with open(out_docx, "wb") as f:
        f.write(docx_buf.getvalue())
    print(f"\nSaved Word Report: {out_docx} ({os.path.getsize(out_docx)} bytes)")

    with open(out_pdf, "wb") as f:
        f.write(pdf_buf.getvalue())
    print(f"Saved 7-Page PDF Report: {out_pdf} ({os.path.getsize(out_pdf)} bytes)")

    # Verify PDF details
    pdf_reader = pypdf.PdfReader(out_pdf)
    print(f"\nPDF Total Pages: {len(pdf_reader.pages)}")
    assert len(pdf_reader.pages) == 7, "PDF must have 7 pages!"
    
    fields = pdf_reader.get_fields()
    print("PDF OOS ID field:", fields.get('Text Field57', {}).get('/V'))
    print("PDF Event ID field:", fields.get('Text Field2', {}).get('/V'))
    print("PDF Action Level field:", fields.get('Text Field11', {}).get('/V'))

    print("\n" + "="*70)
    print("SUCCESSFULLY GENERATED BOTH WORD (.docx) AND 7-PAGE PDF (.pdf) REPORTS!")
    print("="*70)

if __name__ == "__main__":
    run_test_case()
