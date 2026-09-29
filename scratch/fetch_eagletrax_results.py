import os
import time
import re
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
print(f"Using Session Directory: {session_dir}")

urls = {
    '0584': 'https://etrax.eagleanalytical.com/SubmissionTest/Details/sPDE26htF1Ujkzz%2433Fmdw__',
    '0580': 'https://etrax.eagleanalytical.com/SubmissionTest/Details/9sbaVz3diyMXByV5bAGJ5A__'
}

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        session_dir,
        headless=False
    )
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/', wait_until='domcontentloaded')
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
    
    while time.time() - start < 180:
        time.sleep(2)
        url = page.url
        
        # Save screenshot
        try:
            page.screenshot(path="scratch/login_live.png")
        except Exception:
            pass
            
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
            
        # Check if logged in
        sidebar_items = page.locator("a:has-text('Submissions'), a:has-text('Operations'), a[href*='LogOff'], a[href*='Submission'], .sidebar-nav, .navbar")
        if sidebar_items.count() > 0 and "/Account/Login" not in url and "microsoft" not in url and "sts." not in url:
            print(f"\n[SUCCESS!] Genuinely logged in! Verified element count: {sidebar_items.count()}")
            logged_in = True
            break
            
        if "/Account/Login" not in url and "microsoft" not in url and "sts." not in url and "Details" in url:
            logged_in = True
            break
            
    if logged_in:
        print("\n=== Fetching Test Results for ETX-260908-0584 & 0580 ===")
        for key, u in urls.items():
            print(f"\nNavigating to {key}: {u}")
            page.goto(u, wait_until='domcontentloaded')
            time.sleep(3)
            
            # Click Test Details tab
            tab = page.locator("a:has-text('Test Details'), button:has-text('Test Details'), li:has-text('Test Details') a")
            if tab.count() > 0:
                try:
                    tab.first.click()
                    time.sleep(2)
                except Exception as e:
                    print("Tab click error:", e)
                    
            page.screenshot(path=f"scratch/details_{key}.png")
            print(f"Saved screenshot to scratch/details_{key}.png")
            
            # Extract dropdown or select values
            selects = page.locator("select")
            for i in range(selects.count()):
                sel = selects.nth(i)
                sel_id = sel.get_attribute("id") or sel.get_attribute("name")
                val = page.evaluate("(el) => el.options[el.selectedIndex] ? el.options[el.selectedIndex].text : ''", sel.element_handle())
                print(f"  Select [{sel_id}] = '{val}'")
                
            # Extract inputs
            inputs = page.locator("input[type='text'], input[id*='Result']")
            for i in range(inputs.count()):
                inp = inputs.nth(i)
                val = inp.get_attribute("value")
                inp_id = inp.get_attribute("id")
                if val:
                    print(f"  Input [{inp_id}] = '{val}'")
    else:
        print("\nLogin not completed within timeout.")
        
    ctx.close()
