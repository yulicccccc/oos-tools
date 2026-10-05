import re
content = open(r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\Email\public\index.html', encoding='utf-8').read()
for m in re.finditer(r'id=["\']([^"\']*(?:step|excel|btn-generate|generatebtn)[^"\']*)["\']', content, re.I):
    print(m.group(1))
