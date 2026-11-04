import os
import re
import sys
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

def find_oos_email(oos_id: str, download_dir: str = None):
    """
    Search Outlook Inbox for an OOS assignment email by OOS ID (e.g. 'OOS-262329', '262329').
    Automatically downloads official CORP-FORM-21 / CORP-FORM-22 attachments and extracts metadata.
    """
    clean_id = re.search(r'\d{6}', oos_id)
    if not clean_id:
        print(f"[ERROR] Invalid OOS ID format: {oos_id}")
        return None
    oos_num = clean_id.group(0)
    
    if not download_dir:
        download_dir = os.path.join("scratch", f"OOS-{oos_num}_incoming")
    os.makedirs(download_dir, exist_ok=True)
    
    print(f"Connecting to Outlook to locate OOS-{oos_num}...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    inbox = mapi.GetDefaultFolder(6)
    
    # DASL search query
    query = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{oos_num}%'"
    items = inbox.Items.Restrict(query)
    items.Sort("[ReceivedTime]", True)
    
    print(f"Found {items.Count} email(s) matching OOS-{oos_num}.")
    
    results = []
    for i in range(1, items.Count + 1):
        msg = items.Item(i)
        subj = msg.Subject
        sender = getattr(msg, "SenderName", "")
        recv = str(msg.ReceivedTime)
        
        saved_atts = []
        for att in msg.Attachments:
            fn = att.FileName
            if fn.lower().endswith((".pdf", ".docx", ".xlsx", ".csv")) and not fn.lower().startswith("image") and not fn.lower().startswith("outlook"):
                save_path = os.path.join(download_dir, fn)
                att.SaveAsFile(os.path.abspath(save_path))
                saved_atts.append(save_path)
                
        results.append({
            "Subject": subj,
            "Sender": sender,
            "ReceivedTime": recv,
            "Attachments": saved_atts
        })
        print(f" - [{recv[:10]}] {sender}: {subj[:60]} (Saved {len(saved_atts)} atts)")
        
    return results

if __name__ == "__main__":
    if len(sys.argv) > 1:
        find_oos_email(sys.argv[1])
    else:
        # Test with OOS-262335
        find_oos_email("OOS-262335")
