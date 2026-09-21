import os, json
from docxtpl import DocxTemplate

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docs_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"

save_data = {
    "test_date": "08-May-2026",
    "monthly_cleaning_date": "26-Apr-2026",
    "writer_name": "Simin Mohammad",
    "reader_name": "Maraya Chukwumerije and Sophia Santamaria",
    "cfu_count": "11",
    "suite_num": "115",
    "section_c_other": "N/A SMO 28JUL2026",
    "before_date": "29APR 2026",
    "reader_48h": "MC",
    "oos_id": "OOS-261185",
    "sample_name": "EM SMO 115B Air 08MAY2026",
    "after_date": "14MAY 2026",
    "analyst_name": "Simin Mohammad",
    "analyst_initial": "SMO",
    "sampling_type": "Active Air Sampling of Cleanrooms",
    "event_number": "ETX-260518-0250",
    "writer_initial": "SMO",
    "manual_org": "Penicillium decumbens, Cladosporium tenuissimum, Cladosporium langeronii, Cladosporium halotolerans",
    "cleaner_name": "Tamiru Kotisso and Cuong Du",
    "manager_notified": "Kathan Parikh",
    "media_plate_lot": "1011691160",
    "manager_signer": "Robin Seymour",
    "reader_5d": "SAS",
    "bsc_num": "N/A",
    "action_level": "Action level: >= 10 CFU/Plate",
    "media_plate_exp": "17-Dec-2026",
    "bsc_id": "CR115 (Sensor E001737)"
}

# 1. Save JSON session files
save_txt_desktop = os.path.join(desktop_dir, "SAVE_OOS-261185 EM SMO 115B Air 08MAY2026 - EM.txt")
save_txt_docs = os.path.join(docs_dir, "SAVE_OOS-261185 EM SMO 115B Air 08MAY2026 - EM.txt")

with open(save_txt_desktop, "w", encoding="utf-8") as f:
    json.dump(save_data, f, indent=2)
with open(save_txt_docs, "w", encoding="utf-8") as f:
    json.dump(save_data, f, indent=2)
print("Saved SAVE JSON session files.")

# 2. Render Word Document
part1 = (
    "The analyst involved in the Active Air Sampling plate setup, Simin Mohammad, and the analysts involved in reading the plate, "
    "Maraya Chukwumerije and Sophia Santamaria, were interviewed comprehensively. Their answers are recorded throughout this document.\n\n"
    "The EM plates were stored in compliance with the supplier's recommendations, and their integrity was visually inspected prior to use. "
    "Furthermore, the plates were confirmed to be within their valid expiration dates. All the supplies were thoroughly disinfected according "
    "to MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology). The functionality of both incubators was verified through a review "
    "of data obtained from our comprehensive in-house continuous monitoring system.\n\n"
    "Active Air Sampling was performed by analyst Simin Mohammad during weekly environmental monitoring processing in Suite 115 (ISO 8), "
    "Suite 115A (ISO 7), and Suite 115B (ISO 7) on 08 May 2026 as per MICRO-SOP-2 (Environmental Monitoring of the Cleanroom Facility). "
    "The plates were initially incubated at a temperature of 30–35°C in incubator E001031 for a minimum duration of 48 hours, commencing on "
    "08 May 2026. Following completion of a minimum of 48 hours of incubation on 11 May 2026, the plates were further incubated for a minimum of "
    "5 days at 20–25°C in incubator E001034, with the incubation concluding on 18 May 2026. Please see Table 1 for detailed information on the "
    "observations during respective incubations.\n\n"
    "Based on the observations in Table 1, since the CFU count exceeded the action level for one of the Active Air Sampling plates (115B), "
    "the plate was submitted for Microbial Identification under ETX-260518-0250. The colonies were identified as Penicillium decumbens, "
    "Cladosporium tenuissimum, Cladosporium langeronii, and Cladosporium halotolerans. To observe if the organisms identified were transient "
    "in nature or recurring, weekly environmental monitoring plates for the clean room were bracketed to include the week before testing, "
    "the week of testing, and the week after testing as detailed in Table 2 (please see attached)."
)

part2 = (
    "Environmental Monitoring Summary: Weekly environmental sampling for the previous week showed 1 CFU for 115 ISO 8 identified as Micrococcus luteus. "
    "Week of testing showed 1 CFU in 115A ISO 7 and 11 CFU in 115B ISO 7 identified as Staphylococcus aureus, Penicillium decumbens, "
    "Cladosporium tenuissimum, Cladosporium langeronii, and Cladosporium halotolerans. Lastly, the following week of testing showed 1 CFU in 115 ISO 8, "
    "11 CFU in 115B ISO 7, 1 CFU on table in ISO 8 115, and 1 CFU on the cart in 115 ISO 8 identified as Gram (+) cocci, Candida orthopsilosis, "
    "Budding yeast, Gram (+) rods, and Staphylococcus capitis.\n\n"
    "During the interview with the analyst, they indicated that no obvious abnormalities or deviations in the testing procedure were observed. "
    "All the samples were thoroughly disinfected prior to testing. Moreover, the CR115 was thoroughly cleaned and prepared before initiating "
    "the testing as per MICRO-SOP-2 (Environmental Monitoring of the Cleanroom Facility) and MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology).\n\n"
    "Monthly cleaning and disinfection of the outermost ISO 8 Anteroom (Suite 115), the middle ISO 7 Buffer room (Suite 115A), the innermost "
    "ISO 7 clean room (Suite 115B), and its containing ISO 5 Biosafety Cabinets were performed on 26-Apr-2026 by analysts Tamiru Kotisso and "
    "Cuong Du as per MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology). It was documented that all H2O2 indicators passed. "
    "This confirms the efficient monthly cleaning of all three suites 115, 115A, and 115B. Additionally, cleaning and disinfecting was "
    "performed both prior to and after the testing process as per MICRO-SOP-9."
)

part3 = (
    "Based on the findings outlined in the preceding sections, the Out-Of-Specification (OOS) result observed for the Environmental Monitoring (EM) "
    "Active Air Sampling plate may be attributed to a potential analyst error. It is to be noted that the growth observed on the weekly active air "
    "and weekly surface sampling plates for the day of testing do not follow a trend, indicating that the contamination was transient in nature and "
    "that routine daily disinfection procedures were effective in eliminating the contamination. Furthermore, no trend was observed in the clean "
    "rooms previous weekly EM data, therefore, no preventive and corrective actions are deemed necessary at this time."
)

personnel_block = (
    "Simin Mohammad\n(Weekly Active Air Sampling Plate Setup)\n\n"
    "Maraya Chukwumerije\n(Weekly Active Air Sampling Plate Reader)\n\n"
    "Sophia Santamaria\n(Weekly Active Air Sampling Plate Reader)"
)

ctx = {
    "oos_id": "261185",
    "sample_id": "ETX-260518-0250",
    "event_number": "ETX-260518-0250",
    "etx_id": "ETX-260518-0250",
    "sample_name": "EM SMO 115B Air 08MAY2026",
    "lot_number": "EM SMO 115B Air 08MAY2026",
    "dosage_form": "Plate",
    "test_date": "08-May-2026",
    "d_start": "08-May-2026",
    "d_48h": "11-May-2026",
    "d_5d": "18-May-2026",
    "date_initiated": "18-May-2026",
    "date_of_incident": "18-May-2026",
    "analyst_name": "Simin Mohammad",
    "analyst_initial": "SMO",
    "setup_analyst_initial": "SMO",
    "reader_name": "Maraya Chukwumerije and Sophia Santamaria",
    "analyst_signature": "Simin Mohammad",
    "analyst_personnel_block": personnel_block,
    "smart_personnel_block": personnel_block,
    "bsc_id": "CR115 (Sensor E001737)",
    "equipment_summary": "Incubator E001031 and Incubator E001034",
    "cr_display": "CR115 (Sensor E001737)",
    "cr_exp": "December 2026",
    "cr_id": "CR115 (Sensor E001737)",
    "action_level": "Action level: >= 10 CFU/Plate",
    "incident_description": "The CFU count for the environmental monitoring plate exceeded the action level.",
    "smart_incident_opening": "The CFU count for the environmental monitoring plate exceeded the action level.",
    "report_header": "ETX-260518-0250",
    "client_name": "Eagle Analytical Internal EM",
    "smart_comment_interview": "Yes, analysts Simin Mohammad, Maraya Chukwumerije, and Sophia Santamaria were comprehensively interviewed.",
    "smart_comment_records": "Yes, Information is available in EagleTrax under ETX-260518-0250",
    "smart_comment_samples": "Yes, as per MICRO-SOP-2",
    "smart_comment_storage": "Yes, as per MICRO-SOP-2",
    "narrative_summary": f"{part1}\n\n{part2}\n\n{part3}",
    "smart_phase1_summary": f"{part1}\n\n{part2}",
    "smart_phase1_continued": part3,
    "smart_phase1_part1": part1,
    "smart_phase1_part2": f"{part2}\n\n{part3}",
    "plate_media_type": "TSA Plate",
    "media_plate_lot": "1011691160",
    "media_plate_exp": "17-Dec-2026",
    "reagent_lot_str": "TSA Plate:\n1011691160",
    "reagent_exp_str": "TSA Plate:\n17-Dec-2026",
    "manager_notified": "Kathan Parikh",
    "manager_signer": "Robin Seymour",
    "section_c_other": "N/A SMO 28JUL2026"
}

tpl_path = os.path.join(docs_dir, "EM OOS P1 template.docx")
doc = DocxTemplate(tpl_path)
doc.render(ctx)

target_doc_desktop = os.path.join(desktop_dir, "OOS-261185 EM SMO 115B Air 08MAY2026 - EM.docx")
target_doc_docs = os.path.join(docs_dir, "OOS-261185 EM SMO 115B Air 08MAY2026 - EM.docx")

doc.save(target_doc_desktop)
doc.save(target_doc_docs)
print(f"Rendered Word report saved to:\n  {target_doc_desktop}\n  {target_doc_docs}")
