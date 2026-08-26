import sys
import os

sys.path.insert(0, os.getcwd())

import em_logic as el

text = """OOS-261187	ScanC/O CGS E001309 S1 11MAY2026	ETX-260518-0273 Test Results
 Total CFU Count on Plate
0
 Number of Organisms Identified
0
 Colony Description (Optional)
White, shiny, hardened, smooth area embeded on agar
 Microbial Identification (Optional)
N/A
 Stain Results (Optional)"""

parsed = el.parse_em_text(text)
print("Updated EM Parser Test Results:")
for k, v in parsed.items():
    print(f"  - {k}: {v}")
