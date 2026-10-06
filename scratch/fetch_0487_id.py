import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
print(f"Using Session Directory: {session_dir}")

target_url = 'https://etrax.eagleanalytical.com/SubmissionTest/Details/fd3G2StZClcy1TP2ES6BLw__#TestDetails'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        session_dir,
        headless=True
    )
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    print(f"Navigating to: {target_url}")
    page.goto(target_url, wait_until='domcontentloaded')
    time.sleep(2)
    
    # Auto fill username if on login page
    if "/Account/Login" in page.url:
        try:
            inp = page.locator("input[name='Username'], input[id='Username']").first
            if inp.count() > 0:
                inp.fill("qchen")
                btn = page.locator("input[type='submit'], button[type='submit']").first
                if btn.count() > 0:
                    btn.click()
                    print("Auto-filled 'qchen' and clicked Continue.")
        except Exception as e:
            print(f"Info: {e}")
            
    time.sleep(3)
    
    # Click Test Details tab if available
    tab = page.locator("a:has-text('Test Details'), button:has-text('Test Details'), li:has-text('Test Details') a")
    if tab.count() > 0:
        try:
            tab.first.click()
            time.sleep(2)
        except Exception as e:
            print("Tab click error:", e)
            
    page.screenshot(path="scratch/etrax_0487_details.png")
    print("Saved screenshot to scratch/etrax_0487_details.png")
    
    body_text = page.locator("body").inner_text()
    with open("scratch/etrax_0487_text.txt", "w", encoding="utf-8") as f:
        f.write(body_text)
        
    print("\n=== PAGE TEXT SUMMARY ===")
    for line in body_text.splitlines():
        line_s = line.strip()
        if any(k in line_s for k in ['Total CFU', 'Organisms', 'Colony', 'Microbial', 'Stain', 'Result', 'Approved', 'Status', 'Completed']):
            print(f"  {line_s}")
            
    ctx.close()
