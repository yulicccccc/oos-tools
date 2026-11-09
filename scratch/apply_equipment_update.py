with open('scratch/generate_master_oos_262394.py', 'r', encoding='utf-8') as f:
    text = f.read()

target1 = '''text_field_51 = "N/A QYC 20Oct26"

tf43_equipment = (
    "Incubator E001034\\r \\r"
    "(Sensor E001501)\\r \\r"
    "Incubator E001031\\r \\r"
    "(Sensor E001505)\\r \\r"
    f"BSC E00{BSC_PROC}\\r \\r"
    f"BSC E00{BSC_CHG}"
)
tf44_cal_due = (
    "Aug 2027\\r \\r"
    "Feb 2027\\r \\r"
    "Aug 2027\\r \\r"
    "Feb 2027\\r \\r"
    "Jul 2027\\r \\r"
    "Jul 2027"
)'''

rep1 = '''text_field_51 = "N/A QYC 09Nov26"

tf43_equipment = (
    "Incubator E001034\\r \\r"
    "(Sensor E001501)\\r \\r"
    "Incubator E001031 \\r \\r"
    "(Sensor E001505)\\r \\r"
    "Incubator E001933\\r \\r"
    "Incubator E001932\\r \\r"
    "Incubator E001033 \\r \\r"
    "(Sensor E001491)\\r \\r"
    "BSC E001312\\r \\r"
    "BSC E001937"
)
tf44_cal_due = (
    "Aug 2027\\r \\r"
    "Feb 2027 \\r \\r"
    "Aug 2027\\r \\r"
    "Feb 2027\\r \\r"
    "Mar 2027\\r \\r"
    "Mar 2027 \\r \\r"
    "Jul 2027 \\r \\r"
    "Feb 2027\\r \\r"
    "Jun 2027\\r \\r"
    "Jul 2027"
)'''

assert target1 in text, 'target1 not found'
text = text.replace(target1, rep1)

target2 = "'Text Field48': \"No assignable cause\","
rep2 = "'Text Field48': \"N/A QYC 09Nov26\","
assert target2 in text, 'target2 not found'
text = text.replace(target2, rep2)

target3 = "'Text Field53': \"N/A\",\n    'Text Field54': \"N/A\","
rep3 = "'Text Field53': \"Qiyue Chen\",\n    'Text Field54': \"\","
assert target3 in text, 'target3 not found'
text = text.replace(target3, rep3)

target4 = "'Check Box78': 'Yes', 'Check Box80': 'Yes', 'Check Box84': 'Yes',"
rep4 = "'Check Box78': 'Yes', 'Check Box79': 'Yes', 'Check Box80': 'Off', 'Check Box84': 'Off',"
assert target4 in text, 'target4 not found'
text = text.replace(target4, rep4)

# Also update desktop copy in generate_master
target_copy = '''# 4. Sync to Desktop per Rule 0M
dst_desktop = os.path.join(DESKTOP_DIR, os.path.basename(final_master_pdf))
shutil.copy2(final_master_pdf, dst_desktop)
print("Master PDF synced to Desktop:", dst_desktop)'''

rep_copy = '''# 4. Sync to Desktop per Rule 0M
dst_desktop = os.path.join(DESKTOP_DIR, os.path.basename(final_master_pdf))
try:
    shutil.copy2(final_master_pdf, dst_desktop)
    print("Master PDF synced to Desktop:", dst_desktop)
except Exception as e:
    print(f"Direct overwrite of {dst_desktop} failed (locked by PDF reader): {e}")
    dst_desktop_v = os.path.join(DESKTOP_DIR, f"CORP-FORM-21- 20 Oct 2026 OOS-{OOS_ID} Optio RX (E18631) - ScanRDI (v12.4.24).pdf")
    shutil.copy2(final_master_pdf, dst_desktop_v)
    print("Master PDF synced to versioned Desktop copy:", dst_desktop_v)'''

if target_copy in text:
    text = text.replace(target_copy, rep_copy)

with open('scratch/generate_master_oos_262394.py', 'w', encoding='utf-8') as f:
    f.write(text)

print('Updated generate_master_oos_262394.py successfully!')
