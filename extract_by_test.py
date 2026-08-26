import fitz
import os
import glob
import re
from collections import defaultdict

pdf_dir = r"G:\CRO\OOS\2026"
pdf_files = glob.glob(os.path.join(pdf_dir, "**", "*.pdf"), recursive=True)

results = defaultdict(list)

for pf in pdf_files:
    try:
        doc = fitz.open(pf)
        text = ""
        for page in doc:
            text += page.get_text("text") + "\n"
            
        test_type = "Unknown"
        # Determine test type
        if re.search(r"Scan RDI|ScanRDI", text, re.IGNORECASE):
            test_type = "Scan RDI"
        elif re.search(r"Celsis", text, re.IGNORECASE):
            test_type = "Celsis"
        elif re.search(r"USP <71>|USP 71", text, re.IGNORECASE):
            test_type = "USP <71>"
            
        # Extract Justification text: look for paragraphs containing keywords
        justifications = []
        sentences = re.split(r'\.\s+', text.replace('\n', ' '))
        
        for s in sentences:
            if re.search(r"although microbial growth|not identical|background room|viable transfer pathway|layered disinfection|unlikely that the failing results", s, re.IGNORECASE):
                justifications.append(s.strip() + ".")
                
        if justifications:
            results[test_type].append(f"--- File: {os.path.basename(pf)} ---\n" + "\n".join(justifications))
            
    except Exception as e:
        print(f"Error reading {os.path.basename(pf)}: {e}")

output_file = "oos_analysis_by_test.txt"
with open(output_file, "w", encoding="utf-8") as f:
    for t_type, justs in results.items():
        f.write(f"================ {t_type} ================\n")
        f.write("\n\n".join(justs))
        f.write("\n\n")

print(f"Extraction complete. Grouped into {list(results.keys())}")
