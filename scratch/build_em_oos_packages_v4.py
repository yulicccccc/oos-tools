import os, sys, shutil
import fitz
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, TextStringObject
import docx
import win32com.client as win32

from build_docx_table_helper import build_docx_table

ref_pdf_path = r'references/OOS-261553 EM SMO 142 Air 26JUN2026.pdf'

# Build docx and pdf tables for 262133 and 262134
docx_33 = r"scratch\EM table OOS-262133 04SEP2026.docx"
docx_34 = r"scratch\EM table OOS-262134 04SEP2026.docx"
pdf_table_33 = r"scratch\EM table OOS-262133 04SEP2026.pdf"
pdf_table_34 = r"scratch\EM table OOS-262134 04SEP2026.pdf"

print("1. Generating Word master tables...")
build_docx_table("OOS-262133", "262133", docx_33)
build_docx_table("OOS-262134", "262134", docx_34)

print("2. Converting tables to PDF via Word COM...")
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

for d_path, p_path in [(docx_33, pdf_table_33), (docx_34, pdf_table_34)]:
    doc = word.Documents.Open(os.path.abspath(d_path))
    doc.SaveAs(os.path.abspath(p_path), FileFormat=17)
    doc.Close()
    doc_fitz = fitz.open(p_path)
    print(f"  {p_path} -> Pages: {len(doc_fitz)}")
    assert len(doc_fitz) == 1, f"Table PDF {p_path} must be 1 page!"

# Build narratives for OOS-262133 (CR143)
p3_narrative_262133 = (
    "The analyst involved in the weekly active air sampling plate setup, Simin Mohammad, and the analyst involved in reading the plate, "
    "Maraya Chukwumerije, were interviewed comprehensively. Their answers are recorded throughout this document.\r \r"
    "The EM plates were stored in compliance with supplier recommendations, and their integrity was visually inspected prior to use. "
    "Plates were confirmed to be within valid expiration dates, and all supplies were thoroughly disinfected per MICRO SOP-9 (Cleaning "
    "and Disinfecting Procedure for Microbiology). Functionality of both incubators was verified through our continuous monitoring system.\r \r"
    "Weekly active air sampling was performed by analyst Simin Mohammad in ISO 8 Cleanroom Suite 143 on 04 Sep 2026 as per MICRO-SOP-2 - "
    "Environmental Monitoring of the Cleanroom Facility.\r \r"
    "The plates were initially incubated at 30–35°C in incubator E001031 for a minimum duration of 48 hours, commencing on 04 Sep 2026. "
    "Following 48 hours of incubation on 08 Sep 2026, the plates were further incubated for a minimum of 5 days, concluding on 14 Sep 2026. "
    "Please see Table 1 for detailed incubation observations.\r \r"
    "Based on Table 1, since the CFU count exceeded the action level for the ISO 8 CR143 active air plates, they were submitted for "
    "Microbial Identification under ETX-260914-0493 and ETX-260914-0496. Recovered isolates were confirmed as Staphylococcus capitis, "
    "Micrococcus luteus, Moraxella osloensis, Corynebacterium segmentosum, Kytococcus sedentarius, Staphylococcus hominis, Paracoccus yeei, "
    "Brevibacterium sp. CS2, Corynebacterium ureicelerivorans, Micrococcus sp., Corynebacterium sanguinis, Staphylococcus haemolyticus, "
    "Corynebacterium genitalium, Cordyceps memorabilis, Allophoma tropica, Corynebacterium jeikeium, Pseudomonas fulva, Lophiotremataceae, "
    "Acrodontium salmoneum, Quambalaria cyanescens, Sordariales, and Streptococcus sp. A12 (see Table 1).\r \r"
    "To observe if the organisms identified were transient in nature or recurring, weekly environmental monitoring plates for cleanroom "
    "suite 143 were bracketed to include date before testing and date after testing as detailed in Table 2 (please see attached)."
)

p4_narrative_262133 = (
    "Environmental Monitoring Summary: During the preceding week of active air monitoring (28 Aug 2026), 5 CFUs were recovered from CR142, "
    "while zero growth (0 CFU) was recovered across CR143 Section I, CR143 Section II, and all ISO 7 cleanrooms (CR144 and CR145), conforming to applicable limits.\r \r"
    "During the week under investigation (04 Sep 2026), weekly active air monitoring in CR143 yielded 53 CFUs in Section I (ETX-260914-0493) "
    "and 44 CFUs in Section II (ETX-260914-0496), exceeding the ISO 8 action level of > 10 CFU/plate. Additionally, 17 CFUs were recovered in CR142 "
    "(ETX-260914-0498). Notably, zero growth (0 CFU) was recovered across all ISO 7 buffer areas (CR144 and CR145), confirming the excursion was strictly "
    "confined to the ISO 8 anterooms. During the subsequent week (09 Sep 2026), 16 CFUs were recovered in CR143 Section I and 17 CFUs in Section II, "
    "while CR142 dropped to 3 CFUs (below alert level), and CR144 Section II had 1 CFU. Recovered colonies were characterized and confirmed as "
    "Corynebacterium simulans, Staphylococcus epidermidis, Staphylococcus pasteuri, Micrococcus luteus, Staphylococcus hominis, Rothia dentocariosa, "
    "Corynebacterium imitans, Staphylococcus sp. M0911, Moraxella osloensis, Streptococcus thermophilus, Aeromicrobium erythreum, Roseomonas sp. FDAARGOS_362, "
    "Agrococcus carbonis, Corynebacterium segmentosum, and Streptomyces xiamenensis (see Table 2).\r \r"
    "Weekly surface monitoring during the week before testing (28 Aug 2026) recovered 1 CFU from Table 2 in CR143 (Gram-positive rods). Crucially, "
    "weekly surface monitoring contact plates collected during the week under investigation (04 Sep 2026) and the subsequent week (09 Sep 2026) "
    "consistently yielded No Growth (0 CFU) across all cleanrooms and anterooms.\r \r"
    "A summary of all relevant environmental monitoring results is provided in Table 2.\r \r"
    "During the interview with the analyst, they indicated that no obvious abnormalities or deviations in the testing procedure were observed. "
    "All the supplies and sampling equipment were thoroughly disinfected prior to testing. The ISO 8 CR 143 was prepared, cleaned, and disinfected "
    "before initiation of testing in accordance with MICRO-SOP-2 and MICRO-SOP-9.\r \r"
    "Monthly cleaning and disinfection of the ISO 8 Anteroom (CR142), ISO 8 Anteroom (CR143), ISO 7 Buffer Room (CR144), ISO 7 Buffer Room (CR145), "
    "and associated ISO 5 Biosafety Cabinets were performed on 30 Aug 2026 in accordance with MICRO-SOP-9, by laboratory technician Andrew Carrillo. "
    "Documentation confirmed that all hydrogen peroxide (H2O2) indicators passed acceptance criteria, demonstrating successful completion of the "
    "monthly decontamination process. Routine cleaning and disinfection were performed both before and after testing.\r \r"
    "Notably, a comprehensive review of laboratory testing records confirmed that no client samples processed within that week in CR 143 failed sterility testing (including ScanRDI, Celsis, and USP <71>)."
)

p5_narrative_262133 = (
    "Based on the available evidence, the recovery of 53 CFUs and 44 CFUs from the weekly active air sampling plates in CR143 collected on 04 Sep 2026 "
    "represents an isolated environmental monitoring excursion localized to the ISO 8 anteroom area. The primary ISO 7 cleanroom suites and critical ISO 5 "
    "processing zones demonstrated continuous absence of microbial recovery on the date of testing, and all surface monitoring plates yielded No Growth.\r \r"
    "The environmental monitoring results and organism identification did not indicate an ongoing or systemic contamination issue in the critical testing "
    "environments. The investigation identified no impact to client samples, testing activities, or laboratory operations attributable to the event. "
    "A review of sample results, environmental monitoring data, and associated laboratory activities did not identify any evidence that samples were "
    "adversely affected by the positive environmental monitoring result.\r \r"
    "Taken together, the available data indicate that the environmental monitoring event was transient and non-recurring with respect to sample integrity, "
    "with no evidence of sample impact or a systemic environmental control deficiency in the critical processing suites. The findings are most consistent "
    "with an isolated environmental monitoring excursion, with potential laboratory-related factors considered as a possible contributor. Therefore, "
    "no additional corrective or preventive actions are warranted at this time beyond continued routine environmental monitoring, adherence to established "
    "cleaning and disinfection controls, and compliance with existing procedures."
)

# Build narratives for OOS-262134 (CR142)
p3_narrative_262134 = (
    "The analyst involved in the weekly active air sampling plate setup, Simin Mohammad, and the analyst involved in reading the plate, "
    "Maraya Chukwumerije, were interviewed comprehensively. Their answers are recorded throughout this document.\r \r"
    "The EM plates were stored in compliance with supplier recommendations, and their integrity was visually inspected prior to use. "
    "Plates were confirmed to be within valid expiration dates, and all supplies were thoroughly disinfected per MICRO SOP-9 (Cleaning "
    "and Disinfecting Procedure for Microbiology). Functionality of both incubators was verified through our continuous monitoring system.\r \r"
    "Weekly active air sampling was performed by analyst Simin Mohammad in ISO 8 Cleanroom Suite 142 on 04 Sep 2026 as per MICRO-SOP-2 - "
    "Environmental Monitoring of the Cleanroom Facility.\r \r"
    "The plates were initially incubated at 30–35°C in incubator E001031 for a minimum duration of 48 hours, commencing on 04 Sep 2026. "
    "Following 48 hours of incubation on 08 Sep 2026, the plates were further incubated for a minimum of 5 days, concluding on 14 Sep 2026. "
    "Please see Table 1 for detailed incubation observations.\r \r"
    "Based on Table 1, since the CFU count exceeded the action level for the ISO 8 CR142 active air plate, it was submitted for "
    "Microbial Identification under ETX-260914-0498. Recovered isolates were confirmed as Staphylococcus capitis, "
    "Micrococcus luteus, Moraxella osloensis, Corynebacterium segmentosum, Kytococcus sedentarius, Staphylococcus hominis, Paracoccus yeei, "
    "Brevibacterium sp. CS2, Corynebacterium ureicelerivorans, Micrococcus sp., Corynebacterium sanguinis, Staphylococcus haemolyticus, "
    "Corynebacterium genitalium, Cordyceps memorabilis, Allophoma tropica, Corynebacterium jeikeium, Pseudomonas fulva, Lophiotremataceae, "
    "Acrodontium salmoneum, Quambalaria cyanescens, Sordariales, and Streptococcus sp. A12 (see Table 1).\r \r"
    "To observe if the organisms identified were transient in nature or recurring, weekly environmental monitoring plates for cleanroom "
    "suite 142 were bracketed to include date before testing and date after testing as detailed in Table 2 (please see attached)."
)

p4_narrative_262134 = (
    "Environmental Monitoring Summary: During the preceding week of active air monitoring (28 Aug 2026), 5 CFUs were recovered from CR142, "
    "while zero growth (0 CFU) was recovered across CR143 Section I, CR143 Section II, and all ISO 7 cleanrooms (CR144 and CR145), conforming to applicable limits.\r \r"
    "During the week under investigation (04 Sep 2026), weekly active air monitoring in CR142 yielded 17 CFUs (ETX-260914-0498), "
    "exceeding the ISO 8 action level of > 10 CFU/plate. Additionally, 53 CFUs were recovered in CR143 Section I (ETX-260914-0493) "
    "and 44 CFUs in Section II (ETX-260914-0496). Crucially, zero growth (0 CFU) was recovered across all ISO 7 buffer cleanrooms (CR144 and CR145), "
    "demonstrating that the excursion was strictly confined to the outermost ISO 8 anteroom areas. During the subsequent week (09 Sep 2026), "
    "active air monitoring in CR142 dropped significantly to 3 CFUs (ETX-260916-0737), returning well below the alert level. Other recoveries "
    "included 16 CFUs in CR143 Section I, 17 CFUs in Section II, and 1 CFU in CR144 Section II. Recovered colonies were characterized and confirmed as "
    "Corynebacterium simulans, Staphylococcus epidermidis, Staphylococcus pasteuri, Micrococcus luteus, Staphylococcus hominis, Rothia dentocariosa, "
    "Corynebacterium imitans, Staphylococcus sp. M0911, Moraxella osloensis, Streptococcus thermophilus, Aeromicrobium erythreum, Roseomonas sp. FDAARGOS_362, "
    "Agrococcus carbonis, Corynebacterium segmentosum, and Streptomyces xiamenensis (see Table 2).\r \r"
    "Weekly surface monitoring during the week before testing (28 Aug 2026) recovered 1 CFU from Table 2 in CR143 (Gram-positive rods). "
    "Surface sampling contact plates collected during the week under investigation (04 Sep 2026) and the subsequent week (09 Sep 2026) "
    "consistently yielded No Growth (0 CFU) across all cleanrooms and anterooms.\r \r"
    "A summary of all relevant environmental monitoring results is provided in Table 2.\r \r"
    "During the interview with the analyst, they indicated that no obvious abnormalities or deviations in the testing procedure were observed. "
    "All the supplies and sampling equipment were thoroughly disinfected prior to testing. The ISO 8 CR 142 was prepared, cleaned, and disinfected "
    "before initiation of testing in accordance with MICRO-SOP-2 and MICRO-SOP-9.\r \r"
    "Monthly cleaning and disinfection of the ISO 8 Anteroom (CR142), ISO 8 Anteroom (CR143), ISO 7 Buffer Room (CR144), ISO 7 Buffer Room (CR145), "
    "and associated ISO 5 Biosafety Cabinets were performed on 30 Aug 2026 in accordance with MICRO-SOP-9, by laboratory technician Andrew Carrillo. "
    "Documentation confirmed that all hydrogen peroxide (H2O2) indicators passed acceptance criteria, demonstrating successful completion of the "
    "monthly decontamination process. Routine cleaning and disinfection were performed both before and after testing.\r \r"
    "Notably, a comprehensive review of laboratory testing records confirmed that no client samples processed within that week in CR 142 failed sterility testing (including ScanRDI, Celsis, and USP <71>)."
)

p5_narrative_262134 = (
    "Based on the available evidence, the recovery of 17 CFUs from the weekly active air sampling plate in CR142 collected on 04 Sep 2026 represents "
    "an isolated environmental monitoring excursion within the ISO 8 anteroom. The subsequent weekly active air monitoring performed on 09 Sep 2026 "
    "demonstrated rapid return to normal levels with only 3 CFUs (well below the alert level of 8 CFU/plate), and all critical ISO 7 cleanrooms and surface "
    "contact plates maintained absence of adverse contamination.\r \r"
    "The environmental monitoring results and organism identification did not indicate an ongoing or systemic contamination issue. The investigation "
    "identified no impact to client samples, testing activities, or laboratory operations attributable to the event. A review of sample results, "
    "environmental monitoring data, and associated laboratory activities did not identify any evidence that samples were adversely affected by the positive "
    "environmental monitoring result.\r \r"
    "Taken together, the available data indicate that the environmental monitoring event was transient and non-recurring, with no evidence of sample impact "
    "or a systemic environmental control deficiency. The findings are most consistent with an isolated environmental monitoring excursion, with potential "
    "laboratory-related factors considered as a possible contributor. Therefore, no additional corrective or preventive actions are warranted at this time "
    "beyond continued routine environmental monitoring, adherence to established cleaning and disinfection controls, and compliance with existing procedures."
)

def build_pdf_package(oos_num, sample_id, room_name, p3_text, p4_text, p5_text, specific_table_pdf, out_pdf_path):
    reader = PdfReader(ref_pdf_path)
    writer = PdfWriter()
    writer.append(reader, pages=(0, 6))
    writer.set_need_appearances_writer(True)
    
    field_updates = {
        'Text Field57': str(oos_num),
        'Text Field0': 'Simin Mohammad (Written by: Qiyue Chen)',
        'Date Field0': '04-Sep-2026',
        'Date Field1': '14-Sep-2026',
        'Date Field2': '14-Sep-2026',
        'Text Field1': 'Environmental Monitoring',
        'Text Field2': sample_id,
        'Text Field3': 'Simin Mohammad (Active Air Sampling Plate Setup)\r\rMaraya Chukwumerije\r(Active Air Sampling Plate Reader)',
        'Text Field4': f'EM SMO {room_name} Air 04SEP2026',
        'Text Field5': 'Plate',
        'Text Field6': f'EM SMO {room_name} Air 04SEP2026',
        'Text Field7': 'The CFU count for the environmental monitoring plate exceeded the action level.',
        'Text Field8': 'MICRO-SOP-2',
        'Text Field9': '23 Jul 2026',
        'Text Field10': 'Rev. 16',
        'Text Field11': 'Alert Level: ≥ 8 CFU/Plate\rAction Level: > 10 CFU/Plate',
        'Text Field12': 'Kathan Parikh',
        'Date Field3': '14-Sep-2026',
        'Text Field13': 'Yes, analysts Simin Mohammad and Maraya Chukwumerije were comprehensively interviewed.',
        'Text Field14': 'N/A',
        'Text Field15': 'Yes, as per MICRO-SOP-2',
        'Text Field16': 'Yes, as per MICRO-SOP-2',
        'Text Field17': f'Yes, Information is available in EagleTrax under {sample_id}.',
        'Text Field18': 'Yes, the analysts are trained and qualified by quality to perform the test.',
        'Text Field19': 'Not applicable',
        'Text Field20': 'Not applicable',
        'Text Field21': 'Yes, as per MICRO-SOP-2',
        # Page 2
        'Text Field22': 'TSA Plate:\r1011953660',
        'Text Field23': 'TSA Plate:\r28 May 2027',
        'Text Field24': 'Not Applicable',
        'Text Field25': 'Not Applicable',
        'Text Field26': 'Not Applicable',
        'Text Field27': 'Not Applicable',
        'Text Field28': 'Not Applicable',
        'Text Field29': 'Not Applicable',
        'Text Field30': 'Please see below',
        'Text Field31': 'Please see below',
        'Text Field32': f'CR{room_name}',
        'Text Field33': 'Nov 2026',
        'Text Field34': 'Not Applicable',
        'Text Field35': 'Not Applicable',
        'Text Field36': 'Not Applicable',
        'Text Field37': 'Not Applicable',
        'Text Field38': 'Not Applicable',
        'Text Field39': 'Not Applicable',
        'Text Field40': 'No',
        'Text Field41': 'No',
        'Text Field42': 'Not Applicable',
        'Text Field43': 'Incubator:\rE001031\r \r(Sensor\rE001505)\r \rIncubator:\rE001034\r \r(Sensor\rE001501)',
        'Text Field44': '\rAug 2027\r \r \rFeb 2027\r \r \rAug 2027\r \r \rFeb 2027',
        'Text Field45': 'Not Applicable',
        # Page 3
        'Text Field46': 'Not Applicable',
        'Text Field47': 'Not Applicable',
        'Text Field48': 'N/A QYC 14Sep26',
        'Text Field49': p3_text,
        # Page 4
        'Text Field50': p4_text,
        # Page 5
        'Text Field51': p5_text,
        # Page 6
        'Text Field53': 'Qiyue Chen',
        'Text Field54': 'Robin Seymour'
    }
    
    for page_idx in range(6):
        page = writer.pages[page_idx]
        writer.update_page_form_field_values(page, field_updates)
        
    temp_p1_6 = f"scratch/temp_p1_6_{oos_num}.pdf"
    with open(temp_p1_6, 'wb') as f:
        writer.write(f)
        
    # Open with fitz to set precise font sizes to prevent any clipping/overflow
    final_doc = fitz.open(temp_p1_6)
    for p_idx, field_map in [
        (0, {"Text Field11": 7.0}),
        (2, {"Text Field49": 8.5}),
        (3, {"Text Field50": 8.5}),
        (4, {"Text Field51": 9.0}),
    ]:
        p = final_doc[p_idx]
        for w in p.widgets():
            if w.field_name in field_map:
                w.text_fontsize = field_map[w.field_name]
                w.update()
                
    # Merge with specific 1-page table PDF
    table_doc = fitz.open(specific_table_pdf)
    final_doc.insert_pdf(table_doc)
    final_doc.save(out_pdf_path)
    final_doc.close()
    table_doc.close()
    
    doc_check = fitz.open(out_pdf_path)
    print(f"Generated {out_pdf_path} -> Total pages: {len(doc_check)}")
    assert len(doc_check) == 7, f"Package {out_pdf_path} must be 7 pages!"

print("\n3. Building 7-page PDF packages...")
pdf_pkg_33 = r"scratch\OOS-262133 EM SMO 143 Air 04SEP2026.pdf"
pdf_pkg_34 = r"scratch\OOS-262134 EM SMO 142 Air 04SEP2026.pdf"

build_pdf_package(262133, "ETX-260914-0493, ETX-260914-0496", "143", p3_narrative_262133, p4_narrative_262133, p5_narrative_262133, pdf_table_33, pdf_pkg_33)
build_pdf_package(262134, "ETX-260914-0498", "142", p3_narrative_262134, p4_narrative_262134, p5_narrative_262134, pdf_table_34, pdf_pkg_34)

# Deploy to Desktop locations
onedrive_desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
local_desktop = r"C:\Users\qchen\Desktop"

print("\n4. Deploying to Desktops (Rule 0M compliance)...")
for d_dir in [onedrive_desktop, local_desktop]:
    if os.path.exists(d_dir):
        # 1. Copy Word tables
        shutil.copy2(docx_33, os.path.join(d_dir, "EM table OOS-262133 04SEP2026.docx"))
        shutil.copy2(docx_34, os.path.join(d_dir, "EM table OOS-262134 04SEP2026.docx"))
        # 2. Copy 7-page PDFs
        shutil.copy2(pdf_pkg_33, os.path.join(d_dir, "OOS-262133 EM SMO 143 Air 04SEP2026.pdf"))
        shutil.copy2(pdf_pkg_34, os.path.join(d_dir, "OOS-262134 EM SMO 142 Air 04SEP2026.pdf"))
        
        # 3. Strictly remove any standalone table PDF
        for bad_f in ["EM table OOS-262133 04SEP2026.pdf", "EM table OOS-262134 04SEP2026.pdf"]:
            bp = os.path.join(d_dir, bad_f)
            if os.path.exists(bp):
                os.remove(bp)
                print(f"  Removed prohibited: {bp}")

print("Build and Deployment completed successfully!")
