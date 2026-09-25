import fitz
import os
import shutil

qyc_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261242 EM SMO 115B Air 14MAY2026 - EM - QYC.pdf"
backup_path = r".history\OOS_261242_QYC_backup.pdf"
os.makedirs(os.path.dirname(backup_path), exist_ok=True)
shutil.copy2(qyc_path, backup_path)
print(f"Backed up QYC to: {backup_path}")

doc = fitz.open(qyc_path)

# Page 2: Update Check Box68 = "Yes", Check Box69 = ""
for w in doc[1].widgets():
    if w.field_name == "Check Box68":
        w.field_value = "Yes"
        w.update()
    elif w.field_name == "Check Box69":
        w.field_value = ""
        w.update()

text_p4 = (
    "Environmental Monitoring Bracketing & Trend Assessment:\n"
    "Weekly environmental monitoring plates for the clean room were bracketed to include the week before testing (08-May-2026), "
    "the week of testing (14-May-2026), and the week after testing (22-May-2026) as detailed in Table 2 (attached). The bracketing data "
    "demonstrate that ISO 7 Room 115B had 11 CFUs during the week before testing (08-May-2026, under OOS-261185) and again 11 CFUs "
    "during the week of testing (14-May-2026, under OOS-261242), both reaching/exceeding the established action level of >= 10 CFU/plate. "
    "Acknowledging these two consecutive weekly active-air excursions in Room 115B, the recovered microbial populations were evaluated: "
    "on 08-May-2026, the isolates were filamentous fungal molds identified as Penicillium decumbens, Cladosporium tenuissimum, "
    "Cladosporium langeronii, and Cladosporium halotolerans (along with Staphylococcus aureus), whereas on 14-May-2026, the active air plate "
    "exclusively yielded a budding yeast identified as Candida orthopsilosis (along with Gram (+) cocci in ISO 8 115, and Gram (+) rods and "
    "Staphylococcus capitis on Suite 115 surfaces). While both consecutive weeks exhibited excursions meeting/exceeding the action level in "
    "Room 115B, the recovered organisms belonged to distinctly different microbial classes (filamentous molds vs. budding yeast).\n\n"
    "Analyst Interview & Cleanroom Disinfection:\n"
    "During the comprehensive interview with the setup analyst (Simin Mohammad) and reader (Sophia Santamaria), no procedural deviations, "
    "aseptic breaches, or equipment anomalies were identified. All samples and supplies were disinfected prior to introduction, and Suite 115 "
    "was thoroughly cleaned and prepared before testing as per MICRO-SOP-2 (Environmental Monitoring of the Cleanroom Facility) and "
    "MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology).\n\n"
    "Cleanroom Monthly Disinfection & State of Control:\n"
    "Prior to the event, monthly cleaning and disinfection of the outermost ISO 8 Anteroom (Suite 115), the middle ISO 7 Buffer room "
    "(Suite 115A), the innermost ISO 7 clean room (Suite 115B), and its containing ISO 5 Biosafety Cabinets were performed on 26-Apr-2026 "
    "by analysts Rey Estrada and Tamiru Kotisso as per MICRO-SOP-9. All H2O2 chemical indicator strips passed, confirming established "
    "pre-event facility control. Furthermore, subsequent monthly cleaning and disinfection with H2O2 fogging was performed on 31-May-2026 "
    "by analysts Tamiru Kotisso and Cuong Du with all H2O2 indicators passing, verifying complete facility restoration and ongoing "
    "environmental control. Routine daily and between-session disinfection was actively maintained."
)

text_p5 = (
    "Evaluation of Concurrent Suite 115 Sterility Samples & Cladosporium halotolerans Correlation:\n"
    "A collective review of cleanroom operations and concurrent sterility testing in Suite 115 was performed for the week of testing:\n"
    "1. On 11-May-2026, sample ETX-251218-0360 was processed for USP <71> sterility testing in BSC E001314 by Guanchen Li. On 26-May-2026, "
    "the FTM vial yielded Staphylococcus lugdunensis (Gram (+) cocci, 2 colonies under ETX-260526-0342), which is taxonomically unrelated to "
    "either the mold or yeast recoveries.\n"
    "2. On 12-May-2026, sample ETX-260508-0478 was processed for USP <71> sterility testing in BSC E001314 by Devanshi Shah. On 19-May-2026, "
    "the TSB vial yielded Cladosporium halotolerans (Hyphae, 2 colonies under ETX-260519-0388).\n\n"
    "Collective Assessment of Environmental and Sterility Data:\n"
    "Significantly, Cladosporium halotolerans was recovered from both the preceding active air monitoring event in Room 115B on 08-May-2026 "
    "and from sterility sample ETX-260508-0478 processed inside Suite 115 on 12-May-2026. Evaluating this relationship collectively indicates "
    "that Cladosporium halotolerans fungal spores had a localized presence within Suite 115 during the second week of May, demonstrating an "
    "adverse environmental trend for that specific mold species during that timeframe (addressed and investigated under its respective "
    "sterility failure investigation). However, in evaluating the 14-May-2026 active air excursion under current investigation, the recovered "
    "isolate was Candida orthopsilosis, an asexual budding yeast that is biologically and phylogenetically distinct from Cladosporium halotolerans. "
    "While the consecutive 11 CFU counts on 08-May and 14-May reflect an elevated bioburden in Room 115B during this period, the transition from "
    "mold to yeast indicates that the 14-May excursion did not arise from persistent colonization or spreading of the Cladosporium mold, but "
    "rather represented a separate, transient yeast recovery.\n\n"
    "Cleanroom Clearance & Phase I Root Cause Conclusion:\n"
    "Subsequent active air sampling performed in Room 115B on 22-May-2026 yielded No Growth (0 CFU), and surface contact plates across Suite 115 "
    "on 22-May-2026 likewise yielded No Growth (0 CFU). Furthermore, the facility-wide monthly disinfection and H2O2 fogging completed on "
    "31-May-2026 (all chemical indicators verified passing) successfully remediated any residual mold or yeast bioburden. Based on these "
    "collective findings, the 14-May active air excursion was non-recurring following routine sanitization, and the cleanroom environment has "
    "returned to a verified state of microbiological control. No additional corrective actions are deemed necessary beyond ongoing adherence "
    "to strict cleanroom disinfection protocols."
)

for w in doc[3].widgets():
    if w.field_name == "Text Field50":
        w.field_value = text_p4.replace("₂", "2")
        w.update()

for w in doc[4].widgets():
    if w.field_name == "Text Field51":
        w.field_value = text_p5.replace("₂", "2")
        w.update()

out_pdf = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261242 EM SMO 115B Air 14MAY2026 - EM.pdf"
doc.save(out_pdf)
doc.close()

# Also synchronize to OOS-261242.pdf and QYC copy
std_pdf = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261242.pdf"
shutil.copy2(out_pdf, std_pdf)
print(f"Saved finalized PDF to {out_pdf} and {std_pdf}")

# Render verification pages
doc_v = fitz.open(out_pdf)
os.makedirs("scratch/qyc_verified_pages", exist_ok=True)
for i in range(len(doc_v)):
    pix = doc_v[i].get_pixmap(dpi=150)
    pix.save(f"scratch/qyc_verified_pages/page_{i+1}.png")
print("Rendered all 8 pages to scratch/qyc_verified_pages/")
