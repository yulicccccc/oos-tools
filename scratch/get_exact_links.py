import os, time, sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')
session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='domcontentloaded')
    time.sleep(3)
    page.screenshot(path='scratch/login_or_sub.png')
    print(f"Initial URL: {page.url} | Title: {page.title()}")
    
    if 'Login' in page.title() or '/Account/Login' in page.url:
        print('On login page, logging in...')
        page.locator('input[name="Username"]').fill('qchen')
        page.locator('input[type="submit"]').click()
        time.sleep(5)
        print(f"After submit URL: {page.url} | Title: {page.title()}")
        page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='domcontentloaded')
        time.sleep(3)
        page.screenshot(path='scratch/after_login.png')

        
    for s_id in ['ETX-260921-0148', 'ETX-261005-0773', 'ETX-260929-0335', 'ETX-260929-0341', 'ETX-260929-0344']:
        page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='domcontentloaded')
        time.sleep(1)
        
        # Click Clear to reset all filters
        clr = page.locator("#ClearButton, button:has-text('Clear')").first
        if clr.count() > 0:
            clr.click()
            time.sleep(1)
            
        page.locator('#SearchFieldListDropdown').select_option('1')
        page.locator('#SearchMatchListDropdown').select_option('2')
        page.locator('#srchCriteria').fill(s_id)
        page.locator('#FindButton').click()
        time.sleep(3)
        
        links = page.locator(f"a:has-text('{s_id}')").all()
        print(f"\nResults for {s_id}: found {len(links)}")
        for idx, l in enumerate(links):
            txt = l.text_content().strip()
            href = l.get_attribute("href")
            print(f"  [{idx+1}] Text: {txt} | Href: {href}")
            
    ctx.close()
