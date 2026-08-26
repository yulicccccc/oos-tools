import re
from datetime import datetime

sample_text = """OOS-261187	ScanC/O CGS E001309 S1 11MAY2026	ETX-260518-0273 Test Results
 Total CFU Count on Plate
0
 Number of Organisms Identified
0
 Colony Description (Optional)
White, shiny, hardened, smooth area embeded on agar
 Microbial Identification (Optional)
N/A
 Stain Results (Optional)"""

def parse_em_text(text):
    data = {}
    
    # 1. OOS ID
    oos_match = re.search(r"(OOS-\d+)", text)
    if oos_match:
        data["oos_id"] = oos_match.group(1).strip()
        
    # 2. ETX / Event ID
    etx_match = re.search(r"(ETX-\d{6}-\d{4})", text)
    if etx_match:
        data["event_number"] = etx_match.group(1).strip()
        
    # 3. Plate / Sample Name (e.g. ScanC/O CGS E001309 S1 11MAY2026)
    # Match strings starting with Scan or Sterility or EM before tab/newline
    plate_match = re.search(r"((?:Scan|Sterility|EM)[^\t\r\n]+)", text)
    if plate_match:
        p_name = plate_match.group(1).strip()
        data["sample_name"] = p_name
        
        # Extract date from plate name (e.g. 11MAY2026 or 17FEB2026)
        date_in_plate = re.search(r"(\d{1,2}\s*[A-Za-z]{3}\s*\d{2,4})", p_name)
        if date_in_plate:
            raw_d = date_in_plate.group(1).replace(" ", "")
            try:
                if len(raw_d) >= 9:
                    d_obj = datetime.strptime(raw_d, "%d%b%Y")
                else:
                    d_obj = datetime.strptime(raw_d, "%d%b%y")
                data["test_date"] = d_obj.strftime("%d%b%y")
            except: pass
            
        # Extract BSC / Equipment ID (e.g. BSC1309, E001309, 1309)
        bsc_match = re.search(r"(?:BSC|E00)?(\d{4})", p_name, re.IGNORECASE)
        if bsc_match:
            data["bsc_id"] = f"BSC {bsc_match.group(1)}"
            
        # Infer sampling type
        if "sett" in p_name.lower():
            data["sampling_type"] = "Settling Sampling"
        elif "s1" in p_name.lower() or "s2" in p_name.lower() or "s3" in p_name.lower() or "c/o" in p_name.lower() or "surf" in p_name.lower():
            data["sampling_type"] = "Surface Sampling"
        elif "glove" in p_name.lower() or "pers" in p_name.lower():
            data["sampling_type"] = "Personnel Sampling (Glove)"

    # 4. CFU Count
    cfu_match = re.search(r"Total CFU Count on Plate\s*\n?\s*(\d+)", text, re.IGNORECASE)
    if cfu_match:
        data["cfu_count"] = cfu_match.group(1).strip()
        
    # 5. Colony Description / Organism
    desc_match = re.search(r"Colony Description \(Optional\)\s*\n?\s*([^\n\r]+)", text, re.IGNORECASE)
    if desc_match and desc_match.group(1).strip().upper() != "N/A":
        data["manual_org"] = desc_match.group(1).strip()
        
    org_match = re.search(r"Microbial Identification \(Optional\)\s*\n?\s*([^\n\r]+)", text, re.IGNORECASE)
    if org_match and org_match.group(1).strip().upper() != "N/A":
        if "manual_org" in data:
            data["manual_org"] += f" ({org_match.group(1).strip()})"
        else:
            data["manual_org"] = org_match.group(1).strip()

    return data

result = parse_em_text(sample_text)
print("Parsed Result:")
for k, v in result.items():
    print(f"  {k}: {v}")
