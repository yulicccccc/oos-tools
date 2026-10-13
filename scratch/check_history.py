import sqlite3
import os
import shutil

src = os.path.join(os.path.expanduser('~'), r'AppData\Local\Google\Chrome\User Data\Default\History')
dst = 'scratch/ChromeHistory_temp.db'
shutil.copy2(src, dst)

conn = sqlite3.connect(dst)
c = conn.cursor()
q = "SELECT url, title FROM urls WHERE url LIKE '%eagleanalytical%' ORDER BY last_visit_time DESC LIMIT 25"
for url, title in c.execute(q):
    print(f"{url} | {title}")
conn.close()
