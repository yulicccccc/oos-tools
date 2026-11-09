for script in ['scratch/rebuild_oos_262394_with_form9.py', 'scratch/generate_master_oos_262394.py']:
    with open(script, 'r', encoding='utf-8') as f:
        text = f.read()

    old_p16 = '''p16 = (
    "All samples processed within the same batch were also evaluated for evidence of sample-to-sample cross-contamination. "
    f"The test sample {SAMPLE_ID} was processed as the 12th sample by {PROCESSOR_NAME} on 20Oct26 within BSC E00{BSC_PROC} "
    f"(Cleanroom 144, Parallel Suite). One other sample within the same testing batch (ETX-261019-0668, hs53, documented under OOS-262395 for CasaPharma Rx) "
    f"yielded a positive result for cocci-shaped morphology but was processed as the 53rd sample. The two testing events were separated in sequence "
    f"by 41 intermediate negative samples and exhibited different cellular morphologies (rod-shaped morphology in {SAMPLE_ID} versus cocci-shaped morphology in "
    f"ETX-261019-0668 under OOS-262395). A third positive sample tested on the same date (ETX-261017-0024, documented under OOS-262396 for Link Pharmacy) "
    f"was processed by a different analyst (Perris Gary, 36th sample) in a separate biological safety cabinet (BSC E001317, Room 145) with changeover by America Alanis (BSC E001937, Room 144). "
    f"Furthermore, analyst {PROCESSOR_NAME} verified that gloves were sanitized between samples and standard aseptic techniques were strictly maintained. "
    "Consequently, there is no evidence of sample-to-sample cross-contamination or a shared laboratory vector."
)'''

    new_p16 = '''p16 = (
    "All samples processed within the same batch were also evaluated for evidence of sample-to-sample cross-contamination. "
    f"The test sample {SAMPLE_ID} was processed as the 12th sample by {PROCESSOR_NAME} on 20Oct26 within BSC E00{BSC_PROC} "
    "(Cleanroom 144, Parallel Suite). One other sample within the same testing batch (ETX-261019-0668 under OOS-262395) "
    "yielded a positive result for cocci-shaped morphology but was processed as the 53rd sample. The two testing events were separated in sequence "
    f"by 41 intermediate negative samples and exhibited different cellular morphologies (rod-shaped morphology in {SAMPLE_ID} versus cocci-shaped morphology in "
    "ETX-261019-0668 under OOS-262395). A third positive sample tested on the same date (ETX-261017-0024 under OOS-262396) "
    "was processed by a different analyst (Pagan Gray, 36th sample) in a separate biological safety cabinet (BSC E001317, Room 145). "
    f"Furthermore, analyst {PROCESSOR_NAME} verified that gloves were sanitized between samples and standard aseptic techniques were strictly maintained. "
    "Consequently, there is no evidence of sample-to-sample cross-contamination or a shared laboratory vector."
)'''

    assert old_p16 in text, f'old_p16 not found in {script}'
    text = text.replace(old_p16, new_p16)

    # Also update desktop copy version to v12.4.26
    text = text.replace('(v12.4.25).pdf', '(v12.4.26).pdf')

    with open(script, 'w', encoding='utf-8') as f:
        f.write(text)
    print(f'Updated {script} with clean p16 successfully!')
