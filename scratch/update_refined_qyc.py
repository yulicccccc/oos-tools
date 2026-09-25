import fitz
import os
import shutil

qyc_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261242 EM SMO 115B Air 14MAY2026 - EM - QYC.pdf"
backup_path = r".history\OOS_261242_QYC_backup_before_refine.pdf"
shutil.copy2(qyc_path, backup_path)
print(f"Backed up current QYC to: {backup_path}")

doc = fitz.open(qyc_path)

# Page 2: ensure Check Box68 = 'Yes' (No), Check Box69 = '' (N/A)
for w in doc[1].widgets():
    if w.field_name == "Check Box68":
        w.field_value = "Yes"
        w.update()
    elif w.field_name == "Check Box69":
        w.field_value = ""
        w.update()

text_p4_refined = (
    "Environmental Monitoring Bracketing & Trend Assessment:\n"
    "Weekly environmental monitoring results for the cleanroom were reviewed using a bracketing approach that "
    "included the week before testing (08-May-2026), the week of testing (14-May-2026), and the week after "
    "testing (22-May-2026), as detailed in Table 2 (attached).\n\n"
    "The bracketing data demonstrate that ISO 7 Room 115B had 11 CFU during the week before testing (08-May-2026, "
    "OOS-261185) and again 11 CFU during the week of testing (14-May-2026, OOS-261242). Both results exceeded the "
    "established action level of >= 10 CFU/plate.\n\n"
    "The microbial populations recovered during the two sampling events were also evaluated. The 08-May-2026 isolates "
    "included filamentous molds, including Penicillium decumbens and Cladosporium spp., whereas the 14-May-2026 "
    "active-air plate yielded Candida orthopsilosis, a yeast. Although consecutive weekly action-level excursions occurred "
    "in Room 115B, the organisms recovered during the two events represented different microbial groups. Therefore, the "
    "available identification data do not support persistence or recurrence of the same organism across the two sampling events.\n\n"
    "Analyst Interview & Cleanroom Disinfection:\n"
    "During the comprehensive interview with the setup analyst (Simin Mohammad) and reader (Sophia Santamaria), no "
    "procedural deviations, aseptic breaches, or equipment anomalies were identified. All samples and supplies were "
    "disinfected prior to introduction, and Suite 115 was thoroughly cleaned and prepared before testing as per "
    "MICRO-SOP-2 (Environmental Monitoring of the Cleanroom Facility) and MICRO-SOP-9 (Cleaning and Disinfecting Procedure "
    "for Microbiology).\n\n"
    "Cleanroom Monthly Disinfection & State of Control:\n"
    "Prior to the event, monthly cleaning and disinfection of the outermost ISO 8 Anteroom (Suite 115), the middle ISO 7 "
    "Buffer room (Suite 115A), the innermost ISO 7 clean room (Suite 115B), and its containing ISO 5 Biosafety Cabinets were "
    "performed on 26-Apr-2026 by analysts Rey Estrada and Tamiru Kotisso as per MICRO-SOP-9. All H2O2 chemical indicator strips "
    "passed, demonstrating established pre-event cleaning and routine maintenance of environmental control."
)

text_p5_refined = (
    "Evaluation of Concurrent Suite 115 Sterility Samples & Cladosporium halotolerans Correlation:\n"
    "A collective review of cleanroom operations and concurrent sterility testing in Suite 115 was performed for the week of testing:\n"
    "1. On 11-May-2026, sample ETX-251218-0360 was processed for USP <71> sterility testing in BSC E001314 by Guanchen Li. On 26-May-2026, "
    "the FTM vial yielded Staphylococcus lugdunensis (Gram (+) cocci, 2 colonies under ETX-260526-0342), which is taxonomically unrelated to "
    "either the mold or yeast recoveries.\n"
    "2. On 12-May-2026, sample ETX-260508-0478 was processed for USP <71> sterility testing in BSC E001314 by Devanshi Shah. On 19-May-2026, "
    "the TSB vial yielded Cladosporium halotolerans (Hyphae, 2 colonies under ETX-260519-0388).\n\n"
    "Significantly, Cladosporium halotolerans was recovered from both the preceding active air monitoring event in Room 115B on 08-May-2026 "
    "and from sterility sample ETX-260508-0478 processed inside Suite 115 on 12-May-2026. Evaluating this relationship collectively indicates "
    "that Cladosporium halotolerans fungal spores had a localized presence within Suite 115 during the second week of May, demonstrating an "
    "adverse environmental trend for that specific mold species during that timeframe (addressed and investigated under its respective "
    "sterility failure investigation). However, in evaluating the 14-May-2026 active air excursion under current investigation, the recovered "
    "isolate was Candida orthopsilosis, an asexual budding yeast that is biologically distinct from Cladosporium halotolerans.\n\n"
    "Cleanroom Clearance & Phase I Root Cause Conclusion:\n"
    "The consecutive 11 CFU recoveries observed on 08-May-2026 and 14-May-2026 demonstrate elevated microbial recovery in Room 115B "
    "during this period. However, the organisms recovered from the two sampling events were different microbial populations, and the "
    "available identification data do not indicate persistence or recurrence of the same organism.\n\n"
    "Subsequent active-air monitoring performed in Room 115B on 22-May-2026 yielded No Growth (0 CFU). Surface contact plates collected "
    "across Suite 115 on 22-May-2026 likewise yielded No Growth (0 CFU), demonstrating that the elevated recoveries observed during the "
    "preceding two weeks were not observed during subsequent routine monitoring.\n\n"
    "Facility-wide monthly disinfection and H2O2 fogging were subsequently completed on 31-May-2026, with all applicable chemical "
    "indicators meeting acceptance criteria, providing an additional facility-wide disinfection measure following the excursion.\n\n"
    "Based on the organism identification results, subsequent environmental monitoring data, and completion of routine disinfection "
    "activities, the 14-May-2026 yeast recovery was not observed again during subsequent monitoring, and the available evidence does not "
    "support persistence of the recovered organism within Room 115B. No additional corrective action is considered warranted at this time "
    "beyond continued routine environmental monitoring and adherence to established cleanroom cleaning and disinfection procedures."
)

for w in doc[3].widgets():
    if w.field_name == "Text Field50":
        w.field_value = text_p4_refined.replace("₂", "2")
        w.update()

for w in doc[4].widgets():
    if w.field_name == "Text Field51":
        w.field_value = text_p5_refined.replace("₂", "2")
        w.text_fontsize = 8.2
        w.update()

# Save updated QYC PDF via temp file
temp_out = os.path.join("scratch", "temp_refined_qyc.pdf")
doc.save(temp_out)
doc.close()
shutil.copy2(temp_out, qyc_path)

# Synchronize to Desktop standard names
std_full = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261242 EM SMO 115B Air 14MAY2026 - EM.pdf"
std_short = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261242.pdf"
shutil.copy2(qyc_path, std_full)
shutil.copy2(qyc_path, std_short)
print("Synchronized all Desktop PDF copies!")

# Re-render verification preview images
doc_v = fitz.open(qyc_path)
os.makedirs("scratch/qyc_refined_pages", exist_ok=True)
for i in [1, 3, 4, 5, 7]:
    pix = doc_v[i].get_pixmap(dpi=150)
    pix.save(f"scratch/qyc_refined_pages/page_{i+1}.png")
print("Rendered verified pages 2, 4, 5, 6, 8 to scratch/qyc_refined_pages/")
