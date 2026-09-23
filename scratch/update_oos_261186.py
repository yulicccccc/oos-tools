import os, sys, shutil
import fitz

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docs_dir = r"c:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
history_dir = os.path.join(docs_dir, ".history")
scratch_dir = os.path.join(docs_dir, "scratch")

# Master signed PDF on Desktop
master_pdf_path = os.path.join(desktop_dir, "OOS-261186 EM SMO 116A Air 07MAY2026 - EM - SMO - QYC.pdf")

# Backup before modifying
backup_path = os.path.join(history_dir, "OOS_261186_before_qa_feedback_update.pdf")
shutil.copy2(master_pdf_path, backup_path)
print(f"Backed up master PDF to: {backup_path}")

doc = fitz.open(master_pdf_path)
print(f"Loaded master PDF, total pages: {len(doc)}")

# 1. Page 3: Text Field49 (Monthly Cleaning Update)
p3 = doc[2]
w49 = [w for w in p3.widgets() if w.field_name == "Text Field49"][0]
val49 = w49.field_value

old_cleaning_text = (
    "Monthly cleaning and disinfection of the outermost ISO 8 Anteroom (Suite 116), the middle ISO 7 Buffer room (Suite 116A), "
    "the innermost ISO 7 clean room (Suite 116B), and its containing ISO 5 Biosafety Cabinets were performed on 31-May-2026 "
    "as per MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology) by analyst Tamiru Kotisso and Cuong Du on 31-May-2026. "
    "It was documented that all H₂O₂ indicators passed. This confirms the efficient monthly cleaning of all three suites 116, 116A, "
    "and 116B. Additionally, cleaning and disinfecting was performed both prior to and after the testing process as per MICRO-SOP-9."
)

new_cleaning_text = (
    "Prior to the event, monthly cleaning and disinfection of the outermost ISO 8 Anteroom (Suite 116), the middle ISO 7 Buffer room "
    "(Suite 116A), the innermost ISO 7 clean room (Suite 116B), and its containing ISO 5 Biosafety Cabinets were performed on 26-Apr-2026 "
    "by analysts Rey Estrada and Tamiru Kotisso as per MICRO-SOP-9 (Cleaning and Disinfecting Procedure for Microbiology). It was documented "
    "that all H2O2 indicators passed, confirming the efficient pre-event cleaning and established state of control of Suite 116. Furthermore, "
    "subsequent monthly cleaning was performed on 31-May-2026 by analysts Tamiru Kotisso and Cuong Du with all H2O2 indicators passing, "
    "confirming ongoing facility control. Additionally, cleaning and disinfecting was performed both prior to and after the testing process "
    "as per MICRO-SOP-9."
)

# Replace cleaning text cleanly
if old_cleaning_text in val49:
    val49_new = val49.replace(old_cleaning_text, new_cleaning_text)
elif "performed on 31-May-2026" in val49:
    # Find start of monthly cleaning
    idx = val49.find("Monthly cleaning and disinfection")
    if idx != -1:
        val49_new = val49[:idx] + new_cleaning_text
    else:
        val49_new = val49 + "\r \r" + new_cleaning_text
else:
    val49_new = val49 + "\r \r" + new_cleaning_text

val49_new = val49_new.replace("₂", "2")
w49.field_value = val49_new
w49.update()
print("Updated Page 3 Text Field49.")

# 2. Page 4: Text Field50 (Root Cause Conclusion Update)
p4 = doc[3]
w50 = [w for w in p4.widgets() if w.field_name == "Text Field50"][0]

new_val50 = (
    "It is also important to note that no samples processed in Suite 116 for the week of testing "
    "(03-May-2026 to 09-May-2026) failed testing.\r\n\r\n"
    "Based on the available investigation findings, no specific analyst-related, procedural, "
    "equipment-related, or other laboratory-related cause was identified. The excursion appears "
    "to have been transient and non-recurring, and no assignable root cause was established. "
    "It is to be noted that the growth observed on the weekly active air and weekly surface sampling "
    "plates for the day of testing did not follow a trend, indicating that the contamination was "
    "transient in nature and that routine daily disinfection procedures were effective in eliminating "
    "the contamination. Furthermore, no trend was observed in the clean room's previous weekly EM "
    "data, therefore, no preventive and corrective actions are deemed necessary at this time."
)

new_val50 = new_val50.replace("₂", "2")
w50.field_value = new_val50
w50.update()
print("Updated Page 4 Text Field50.")

# Save updated master PDF
temp_out = os.path.join(scratch_dir, "OOS_261186_updated.pdf")
doc.save(temp_out)
doc.close()

# Copy to targets on Desktop
shutil.copy2(temp_out, master_pdf_path)

sync_targets = [
    os.path.join(desktop_dir, "OOS-261186 EM SMO 116A Air 07MAY2026 - EM.pdf"),
    os.path.join(desktop_dir, "OOS-261186.pdf")
]
for target in sync_targets:
    shutil.copy2(temp_out, target)
    print(f"Synchronized: {target}")

print("=== VERIFICATION: RENDER PREVIEWS ===")
doc_v = fitz.open(master_pdf_path)
print(f"Verified PDF page count: {len(doc_v)}")

for p_num in [3, 4, 6, 8]:
    pix = doc_v[p_num - 1].get_pixmap(dpi=150)
    out_img = os.path.join(scratch_dir, f"oos_261186_updated_p{p_num}.png")
    pix.save(out_img)
    print(f"Saved: {out_img}")

doc_v.close()
print("SUCCESS: OOS-261186 update complete!")
