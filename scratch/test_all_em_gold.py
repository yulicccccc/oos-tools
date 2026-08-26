import sys
sys.path.append('.')
import streamlit as st
import em_logic
import docx
import pypdf

# Case 1: OOS-261401 (GL / BSC E001314 S1 / Artifact)
print("="*80)
print("TESTING OOS-261401 (GL / BSC E001314 S1 / Colony-like Artifact)")
print("="*80)
st.session_state.clear()
st.session_state.update({
    "oos_id": "OOS-261401",
    "sample_name": "Sterility GL E001314 S1 04JUN2026",
    "test_date": "04 Jun 2026",
    "sampling_type": "Surface Sampling",
    "bsc_id": "BSC E001314",
    "analyst_name": "Guanchen (David) Li",
    "analyst_initial": "GL",
    "reader_name": "Maraya Chukwumerije and Simin Mohammad",
    "event_number": "ETX-260615-0424",
    "action_level": "Action level: ≥ 1CFU/Plate",
    "cfu_count": "1",
    "manual_org": "colony-like artifact",
    "monthly_cleaning_date": "26 April 2026",
    "cleaner_name": "Rey Estrada",
    "writer_name": "Dhvanir Kansara",
    "manager_name": "Robin Seymour",
    "media_plate_lot": "1011834770",
    "media_plate_exp": "25 Sep 2026",
    "plate_media_type": "Contact Plate"
})

docx_buf, pdf_buf = em_logic.generate_em_reports()
if docx_buf and pdf_buf:
    with open("scratch/OOS_261401_Test.docx", "wb") as f:
        f.write(docx_buf.read())
    with open("scratch/OOS_261401_Test.pdf", "wb") as f:
        f.write(pdf_buf.read())
    print("Successfully generated OOS_261401 Test Word & PDF!")

# Case 2: OOS-261403 (RE / BSC E001313 S3 / Staphylococcus warneri)
print("="*80)
print("TESTING OOS-261403 (RE / BSC E001313 S3 / Staphylococcus warneri)")
print("="*80)
st.session_state.clear()
st.session_state.update({
    "oos_id": "OOS-261403",
    "sample_name": "Scan RCLE BSC1313 S3 05JUN2026",
    "test_date": "05 Jun 2026",
    "sampling_type": "Surface Sampling",
    "bsc_id": "BSC E001313",
    "analyst_name": "Rey Estrada",
    "analyst_initial": "RE",
    "reader_name": "Maraya Chukwumerije and Simin Mohammad",
    "event_number": "ETX-260615-0435",
    "action_level": "Action level: ≥ 1CFU/Plate",
    "cfu_count": "1",
    "manual_org": "Staphylococcus warneri",
    "monthly_cleaning_date": "26 April 2026",
    "cleaner_name": "Rey Estrada",
    "writer_name": "Dhvanir Kansara",
    "manager_name": "Robin Seymour",
    "media_plate_lot": "1011834770",
    "media_plate_exp": "25 Sep 2026",
    "plate_media_type": "Contact Plate"
})

docx_buf, pdf_buf = em_logic.generate_em_reports()
if docx_buf and pdf_buf:
    with open("scratch/OOS_261403_Test.docx", "wb") as f:
        f.write(docx_buf.read())
    with open("scratch/OOS_261403_Test.pdf", "wb") as f:
        f.write(pdf_buf.read())
    print("Successfully generated OOS_261403 Test Word & PDF!")

# Case 3: OOS-261186 (SMO / 116A Air / Kocuria palustris)
print("="*80)
print("TESTING OOS-261186 (SMO / 116A Air / Kocuria palustris)")
print("="*80)
st.session_state.clear()
st.session_state.update({
    "oos_id": "OOS-261186",
    "sample_name": "EM SMO 116A Air 07MAY2026",
    "test_date": "07 May 2026",
    "sampling_type": "Weekly Active Air Sampling",
    "bsc_id": "CR116",
    "analyst_name": "Simin Mohammad",
    "analyst_initial": "SMO",
    "reader_name": "Samera A. Salim and Simin Mohammad",
    "event_number": "ETX-260518-0254",
    "action_level": "Alert Level: >= 8 CFU/Plate\nAction Level: > 10 CFU/Plate",
    "cfu_count": "748",
    "manual_org": "Kocuria palustris (Gram (+) cocci)",
    "monthly_cleaning_date": "30 Apr 2026",
    "cleaner_name": "Rey Estrada",
    "writer_name": "Maryam Naeem",
    "manager_name": "Kathan Parikh",
    "media_plate_lot": "1011543730",
    "media_plate_exp": "29SEP2026",
    "plate_media_type": "TSA Plate"
})

docx_buf, pdf_buf = em_logic.generate_em_reports()
if docx_buf and pdf_buf:
    with open("scratch/OOS_261186_Test.docx", "wb") as f:
        f.write(docx_buf.read())
    with open("scratch/OOS_261186_Test.pdf", "wb") as f:
        f.write(pdf_buf.read())
    print("Successfully generated OOS_261186 Test Word & PDF!")
