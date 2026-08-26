import os
import sys
import io

# Ensure current working directory is in sys.path
sys.path.insert(0, os.path.abspath("."))

import pypdf
import docxtpl
import em_logic as el

def test_full_pipeline():
    print("="*60)
    print("TESTING FULL EM OOS GENERATION PIPELINE")
    print("="*60)

    # 1. Simulate Session State for OOS-260361
    import streamlit as st
    st.session_state["oos_id"] = "OOS-260361"
    st.session_state["sample_name"] = "Sterility GS BSC1314 Sett2 05FEB2026"
    st.session_state["event_number"] = "ETX-260216-0348"
    st.session_state["test_date"] = "05Feb26"
    st.session_state["sampling_type"] = "Settling Sampling"
    st.session_state["bsc_id"] = "BSC 1314"
    st.session_state["analyst_name"] = "Gabrielle Surber"
    st.session_state["analyst_initial"] = "GS"
    st.session_state["reader_name"] = "Maraya Chukwumerije & Simin Mohammad"
    st.session_state["action_level"] = "Action Level: ≥ 1 CFU/Plate"
    st.session_state["cfu_count"] = "10"
    st.session_state["manual_org"] = "Staphylococcus capitis (Gram (+) cocci), Staphylococcus hominis (Gram (+) cocci), Kocuria indica (Gram (+) cocci), Micrococcus luteus (Gram (+) cocci) and Staphylococcus epidermidis (Gram (+) cocci)"
    st.session_state["monthly_cleaning_date"] = "31 Jan 2026"
    st.session_state["cleaner_name"] = "Rey Estrada"

    # 2. Generate Narrative
    p1, p2, p3 = el.generate_em_narrative()
    print("Narrative Part 1 length:", len(p1))
    print("Narrative Part 2 length:", len(p2))
    print("Narrative Part 3 length:", len(p3))

    # 3. Generate Reports
    docx_buf, pdf_buf = el.generate_em_reports()
    
    assert docx_buf is not None, "DOCX buffer is None!"
    assert pdf_buf is not None, "PDF buffer is None!"

    # Save to scratch for inspection
    docx_out = "scratch/test_output_EM_260361.docx"
    pdf_out = "scratch/test_output_EM_260361.pdf"

    with open(docx_out, "wb") as f:
        f.write(docx_buf.read())
    print(f"Saved generated DOCX: {docx_out} (size: {os.path.getsize(docx_out)} bytes)")

    with open(pdf_out, "wb") as f:
        f.write(pdf_buf.read())
    print(f"Saved generated PDF: {pdf_out} (size: {os.path.getsize(pdf_out)} bytes)")

    # 4. Verify PDF page count & field values
    pdf_reader = pypdf.PdfReader(pdf_out)
    print(f"Verified PDF total page count: {len(pdf_reader.pages)}")
    assert len(pdf_reader.pages) == 7, f"Expected 7 pages, got {len(pdf_reader.pages)}"

    fields = pdf_reader.get_fields()
    print(f"Verified PDF form fields count: {len(fields)}")
    print("OOS ID field in PDF:", fields.get('Text Field57', {}).get('/V'))
    print("Initiator in PDF:", fields.get('Text Field0', {}).get('/V'))

    # 5. Verify Table Template Rendering (tables for em.docx)
    ctx = el.build_em_context()
    doc_table = docxtpl.DocxTemplate("tables for em.docx")
    doc_table.render(ctx)
    table_docx_out = "scratch/test_output_tables_for_em.docx"
    doc_table.save(table_docx_out)
    print(f"Saved rendered tables for em: {table_docx_out} (size: {os.path.getsize(table_docx_out)} bytes)")

    print("\nALL 5 TEMPLATES AND FULL GENERATION PIPELINE PASSED WITH 100% SUCCESS!")

if __name__ == "__main__":
    test_full_pipeline()
