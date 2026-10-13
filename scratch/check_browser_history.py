import os, sqlite3, shutil

edge_hist = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\History")
tmp_path = "scratch/temp_Edge_hist.db"
shutil.copy2(edge_hist, tmp_path)
conn = sqlite3.connect(tmp_path)
cursor = conn.cursor()
cursor.execute("SELECT url, title FROM urls WHERE url LIKE '%eagleanalytical.com%' ORDER BY last_visit_time DESC LIMIT 40")
for r in cursor.fetchall():
    print(r[0])
    print("  Title:", r[1])
conn.close()
