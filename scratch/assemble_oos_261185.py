import os, sys, shutil
import fitz
from pypdf import PdfReader, PdfWriter

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docs_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
history_dir = os.path.join(docs_dir, ".history")
os.makedirs(history_dir, exist_ok=True)

src_form_pdf = os.path.join(desktop_dir, "OOS-261185.pdf")
src_table_pdf = os.path.join(desktop_dir, "EM table OOS-261185 08MAY2026.pdf")

# 1. Use clean original backup from history
backup_pdf = os.path.join(history_dir, "OOS_261185_backup_before_field_update.pdf")

doc = fitz.open(backup_pdf)

field_updates = {
    # Page 1
    "Text Field0": "Simin Mohammad",
    "Text Field3": (
        "Simin Mohammad (Weekly Active Air Sampling Plate Setup)\r"
        "Maraya Chukwumerije (Weekly Active Air Sampling Plate Reader)\r"
        "Sophia Santamaria (Weekly Active Air Sampling Plate Reader)"
    ),
    "Text Field8": "MICRO-SOP-2",
    "Text Field9": "23-Jul-2026",
    "Text Field10": "16",
    "Text Field11": "Action level: >= 10 CFU/Plate",
    "Text Field13": "Yes, analysts Simin Mohammad, Maraya Chukwumerije, and Sophia Santamaria were comprehensively interviewed.",
    "Text Field15": "Yes, as per MICRO-SOP-2",
    "Text Field16": "Yes, as per MICRO-SOP-2",
    "Text Field21": "Yes, as per MICRO-SOP-2",
    
    # Page 2
    "Text Field43": "Incubator E001034 (Sensor E001501)\rIncubator E001031 (Sensor E001505)",
    "Text Field44": "Aug 2026 / Feb 2027\rAug 2026 / Feb 2027",
    
    # Page 3
    "Text Field49": (
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
    ),
    
    # Page 4
    "Text Field50": (
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
    ),
    
    # Page 5
    "Text Field51": (
        "Based on the findings outlined in the preceding sections, the Out-Of-Specification (OOS) result observed for the Environmental Monitoring (EM) "
        "Active Air Sampling plate may be attributed to a potential analyst error. It is to be noted that the growth observed on the weekly active air "
        "and weekly surface sampling plates for the day of testing do not follow a trend, indicating that the contamination was transient in nature and "
        "that routine daily disinfection procedures were effective in eliminating the contamination. Furthermore, no trend was observed in the clean "
        "rooms previous weekly EM data, therefore, no preventive and corrective actions are deemed necessary at this time."
    )
}

for page in doc:
    for w in page.widgets():
        if w.field_name in field_updates:
            new_val = field_updates[w.field_name]
            w.field_value = new_val
            if w.field_name == "Text Field11":
                w.text_fontsize = 6.5
            w.update()
            print(f"Updated {w.field_name}")

temp_updated_pdf = os.path.join(docs_dir, "scratch", "temp_updated_form.pdf")
doc.save(temp_updated_pdf)
doc.close()
print("Saved temporary updated form PDF.")

# 3. Assemble full 7-page PDF (Pages 1-6 form + Page 7 table)
reader_form = PdfReader(temp_updated_pdf)
reader_table = PdfReader(src_table_pdf)

writer = PdfWriter()
for p in reader_form.pages:
    writer.add_page(p)

for p in reader_table.pages:
    writer.add_page(p)

target_full_pdf = os.path.join(desktop_dir, "OOS-261185 EM SMO 115B Air 08MAY2026 - EM.pdf")
with open(target_full_pdf, "wb") as f:
    writer.write(f)
print(f"Saved full assembled package to: {target_full_pdf}")

# Overwrite original OOS-261185.pdf on Desktop
with open(src_form_pdf, "wb") as f:
    writer.write(f)
print(f"Synchronized original Desktop PDF: {src_form_pdf}")

# 4. Render verification images of all 7 pages
doc_final = fitz.open(target_full_pdf)
print(f"Final assembled PDF page count: {len(doc_final)}")
for i, p in enumerate(doc_final):
    pix = p.get_pixmap(dpi=150)
    out_img = os.path.join(docs_dir, "scratch", f"final_261185_p{i+1}.png")
    pix.save(out_img)
print("Saved all 7 verification images to scratch/")
