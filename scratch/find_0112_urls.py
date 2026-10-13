import os
import re

pat = re.compile(r'260901-0112')
base_dir = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc"

for root, dirs, files in os.walk(base_dir):
    if '.git' in root or 'node_modules' in root:
        continue
    for f in files:
        if f.endswith(('.txt', '.html', '.msg', '.eml', '.py', '.json')):
            p = os.path.join(root, f)
            try:
                with open(p, 'r', encoding='utf-8', errors='ignore') as fp:
                    content = fp.read()
                    if pat.search(content):
                        urls = re.findall(r'https?://[^\s"\'<>]+', content)
                        print(f"File: {p}")
                        for u in urls:
                            if 'etrax' in u or 'eagle' in u:
                                print(f"  Eagle URL: {u}")
            except Exception:
                pass
