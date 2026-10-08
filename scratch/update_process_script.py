with open('scratch/process_oos_262098.py', 'r', encoding='utf-8') as f:
    text = f.read()

# Replace p18
old_p18 = '''p18 = (
    "Based on the observations outlined above, it is unlikely that the failing results were due to reagents, supplies, the cleanroom environment, "
    "the process, or analyst involvement. Consequently, the possibility of laboratory error contributing to this failure is minimal and the original "
    "result is deemed to be valid."
)'''

new_p18 = '''p18 = (
    "During the preliminary interview and investigation, a potential laboratory error was identified during the resuspension of the lyophilized sample. "
    "Consequently, there is a likelihood of a potential source of laboratory error from the processing analyst. Importantly, the same sample lot "
    "(2608-216) was previously submitted for ScanRDI Sterility Testing under ETX-260807-0602 with passing results. As part of the Phase II investigation, "
    "additional sample vials were requested to perform retesting. The remaining samples processed in the same batch on 04 Sep 2026 continued to be "
    "monitored with no additional failures. Therefore, a Phase II investigation and retesting were initiated under Form 3.100.019.F02."
)'''

assert old_p18 in text, 'old_p18 not found'
text = text.replace(old_p18, new_p18)

# Replace smart_comment_interview
old_sci = 'data["smart_comment_interview"] = "Yes, analysts Alex Saravia and Elysse Nioupin were interviewed comprehensively."'
new_sci = 'data["smart_comment_interview"] = "Yes. During preliminary interview, a potential laboratory error was identified during resuspension of lyophilized sample."'
assert old_sci in text, 'old_sci not found'
text = text.replace(old_sci, new_sci)

# Update checkboxes on Page 6
old_cb = "'Check Box88': '/Yes',  # No lab error - original result valid"
new_cb = "'Check Box87': '/Yes',  # Lab error identified\n    'Check Box89': '/Yes',  # Initiate Phase II Form 3.100.019.F02\n    'Check Box91': '/Yes',  # Cannot close - initiate Phase II"
assert old_cb in text, 'old_cb not found'
text = text.replace(old_cb, new_cb)

# Insert Text Field52 comment
text = text.replace(
    "'Text Field53': data['writer_name'],",
    "'Text Field52': 'A potential laboratory error was identified during resuspension of lyophilized sample. Initiating Phase II (Form 3.100.019.F02).',\n    'Text Field53': data['writer_name'],"
)

# Add font size for Text Field52
text = text.replace(
    "'Text Field53': 8.5,",
    "'Text Field52': 7.5,\n        'Text Field53': 8.5,"
)

with open('scratch/process_oos_262098.py', 'w', encoding='utf-8') as f:
    f.write(text)

print('Updated scratch/process_oos_262098.py successfully!')
