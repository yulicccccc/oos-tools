import os
from docx import Document
from docx.shared import Pt
import win32com.client as win32

desktop_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
docx_path = os.path.join(desktop_dir, "OOS-261242 Review Email.docx")
html_path = os.path.join(desktop_dir, "OOS-261242_Review_Email_Robin_Format.html")

new_closing = "@Simin Mohammad please review, sign, and route back."

# 1. Update DOCX
doc = Document(docx_path)
for p in doc.paragraphs:
    if "please review and sign" in p.text or "please review, sign" in p.text:
        p.text = new_closing
        p.runs[0].font.name = "Calibri"
        p.runs[0].font.size = Pt(11)

doc.save(docx_path)
print("Updated DOCX closing to: " + new_closing)

# 2. Update HTML
if os.path.exists(html_path):
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()
    import re
    html = re.sub(
        r'<p style="margin: 14px 0 6px 0;[^>]*>[^<]*please review[^<]*</p>',
        f'<p style="margin: 14px 0 6px 0; font-family: Calibri, sans-serif; font-size: 11pt;">{new_closing}</p>',
        html
    )
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print("Updated HTML closing.")

# 3. Copy via Word COM to clipboard
word_app = win32.Dispatch("Word.Application")
word_app.Visible = False
try:
    d = word_app.Documents.Open(os.path.abspath(docx_path))
    d.Content.Copy()
    d.Close(False)
    print("SUCCESS: Copied updated email to Windows Clipboard via Word COM!")
finally:
    word_app.Quit()
