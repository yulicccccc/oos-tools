import fitz
import os
import glob
import re

pdf_dir = r"G:\CRO\OOS\2026"
output_file = "oos_analysis_results.txt"

pdf_files = glob.glob(os.path.join(pdf_dir, "**", "*.pdf"), recursive=True)

results = []

for pf in pdf_files:
    try:
        doc = fitz.open(pf)
        text = ""
        for page in doc:
            text += page.get_text("text") + "\n"
        
        # Look for the sentence about growth
        growth_match = re.search(r"microbial growth was observed.*?(?=\n\n|\n[A-Z])", text, re.IGNORECASE | re.DOTALL)
        growth_desc = growth_match.group(0).replace("\n", " ") if growth_match else "No 'microbial growth was observed' found"
        
        # Look for 'Specifically, '
        spec_match = re.search(r"Specifically, .*?(?=\n\n)", text, re.IGNORECASE | re.DOTALL)
        spec_desc = spec_match.group(0).replace("\n", " ") if spec_match else ""
        
        # Look for Justification or Conclusion
        # Usually it's "Consequently, " or "Therefore, " or "Based on "
        # Let's extract the last 2000 characters to capture the concluding thoughts
        end_text = text[-2000:].replace("\n", " ")
        
        results.append(f"--- File: {os.path.basename(pf)} ---")
        results.append(f"Growth: {growth_desc}")
        if spec_desc:
            results.append(f"Details: {spec_desc}")
        
        # Try to find the exact justification sentence
        justification_match = re.search(r"(?:Based on the observations|Consequently|Therefore|In conclusion).*?(?=\.)", text, re.IGNORECASE | re.DOTALL)
        if justification_match:
            results.append(f"Justification: {justification_match.group(0).replace(chr(10), ' ')}.")
        else:
            # Fallback
            results.append(f"End Text Snippet: {end_text[-500:]}")
            
        results.append("\n")
    except Exception as e:
        results.append(f"Error reading {os.path.basename(pf)}: {e}\n")

with open(output_file, "w", encoding="utf-8") as f:
    f.write("\n".join(results))
print(f"Extracted {len(pdf_files)} PDFs. Results saved to {output_file}")
