import os
import time
import re
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
print(f"Using Session Directory: {session_dir}")

target_url = 'https://etrax.eagleanalytical.com/SubmissionTest/Details/qnwcLQO5BWeJhWCBG7jl8Q__#TestDetails'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        session_dir,
        headless=False
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
            
    last_num = None
    start = time.time()
    logged_in = False
    
    while time.time() - start < 60:
        time.sleep(2)
        url = page.url
        
        # Detect Authenticator challenge number
        try:
            elem = page.locator("//*[@id='richId-number' or contains(@class, 'display-sign-in-large-text') or contains(@class, 'number')]")
            for i in range(elem.count()):
                txt = elem.nth(i).inner_text().strip()
                if txt.isdigit() and len(txt) == 2:
                    if txt != last_num:
                        last_num = txt
                        print(f"\n==========================================", flush=True)
                        print(f"  👉 微软 Authenticator 验证码: 【 {txt} 】", flush=True)
                        print(f"==========================================\n", flush=True)
                    break
        except Exception:
            pass
            
        # Auto click Stay signed in (Yes)
        try:
            kmsi = page.locator("#KmsiCheckboxField").first
            if kmsi.count() > 0 and not kmsi.is_checked():
                kmsi.check()
            yes_btn = page.locator("#idSIButton9, input[type='submit'][value='Yes'], button:has-text('Yes')").first
            if yes_btn.count() > 0:
                yes_btn.click()
                print("Clicked 'Stay signed in' (Yes).")
        except Exception:
            pass
            
        if "Details" in url and "/Account/Login" not in url and "microsoft" not in url:
            logged_in = True
            break
            
    if logged_in:
        print("\n[SUCCESS] Reached Test Details page!")
        time.sleep(2)
        tab = page.locator("a:has-text('Test Details'), button:has-text('Test Details'), li:has-text('Test Details') a")
        if tab.count() > 0:
            try:
                tab.first.click()
                time.sleep(2)
            except Exception as e:
                print("Tab click error:", e)
                
        page.screenshot(path="scratch/etrax_0520_success.png")
        body_text = page.locator("body").inner_text()
        with open("scratch/etrax_0520_success.txt", "w", encoding="utf-8") as f:
            f.write(body_text)
            
        # Print all visible text and input values
        print("\n=== PAGE TEXT SUMMARY ===")
        for line in body_text.splitlines():
            line_s = line.strip()
            if any(k in line_s for k in ['Total CFU', 'Organisms', 'Colony', 'Microbial', 'Stain', 'Result', 'Approved', 'Status']):
                print(f"  {line_s}")
    else:
        print("\nCould not reach page within 60s.")
        
    ctx.close()
