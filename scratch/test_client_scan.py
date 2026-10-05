import re

def test_client_name_scan():
    case1 = """
Hi @Eagle Client Care Team,

After review, please inform the client that the following sample has been found to yield incompatible results after Scan RDI testing:

AnazaoHealth Corporation - Tampa

ETX-260908-0123

Sample: Morphine Sulfate 10 mg/mL
"""
    lines = [l.strip() for l in case1.split('\n') if l.strip()]
    etx_idx = None
    for idx, l in enumerate(lines):
        if 'ETX-260908-0123' in l:
            etx_idx = idx
            break
            
    ignore = ["after review", "hi @", "client care", "attention", "attn:", "subject:", "to:", "cc:", "scan rdi testing:"]
    client_name = ""
    for k in range(etx_idx - 1, max(-1, etx_idx - 5), -1):
        cand = lines[k].strip()
        if not cand: continue
        cand_lower = cand.lower()
        if any(ig in cand_lower for ig in ignore) or cand_lower.startswith('conclusion:') or cand_lower.startswith('---') or re.search(r'ETX-\d{6}-\d{4}', cand, re.I):
            continue
        client_name = cand
        break
        
    print("Detected client name:", repr(client_name))

test_client_name_scan()
