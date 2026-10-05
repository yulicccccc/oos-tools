import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
target_url = 'https://etrax.eagleanalytical.com/SubmissionTest/Details/OCvPHc7TodycYzvim2KtjQ__#TestDetails'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto(target_url, wait_until='domcontentloaded')
    time.sleep(2)
    
    if '/Account/Login' in page.url:
        try:
            inp = page.locator("input[name='Username'], input[id='Username']").first
            if inp.count() > 0:
                inp.fill('qchen')
                btn = page.locator("input[type='submit'], button[type='submit']").first
                if btn.count() > 0:
                    btn.click()
                    print('Clicked login submit.')
        except Exception as e:
            print('Login fill err:', e)
            
    time.sleep(4)
    
    tab = page.locator("a:has-text('Test Details'), button:has-text('Test Details')")
    if tab.count() > 0:
        tab.first.click()
        time.sleep(2)
        
    page.evaluate('window.scrollBy(0, 400)')
    time.sleep(1)
    page.screenshot(path='scratch/etrax_262017_micro_lower.png')
    
    # Extract all visible text and input values
    rows = page.locator("div.form-group, tr, div[class*='row']")
    for i in range(rows.count()):
        r = rows.nth(i)
        t = r.inner_text().strip()
        if any(k in t.lower() for k in ['total cfu', 'organisms', 'colony description', 'microbial identification', 'stain results']):
            # print single line
            clean_t = ' | '.join(t.splitlines())
            if len(clean_t) < 200:
                print(f"  {clean_t}")
                
    ctx.close()
