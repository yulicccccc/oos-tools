import os
import time
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
scanrdi_url = 'https://etrax.eagleanalytical.com/SubmissionTest/Details/zFdCKB28Sd0gPid31SfEnA__'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    page.goto(scanrdi_url, wait_until='domcontentloaded')
    time.sleep(2)
    
    if "/Account/Login" in page.url or "Login" in page.title():
        inp = page.locator("input[name='Username'], input[id='Username']").first
        if inp.count() > 0:
            inp.fill('qchen')
            btn = page.locator("input[type='submit'], button[type='submit']").first
            btn.click()
            time.sleep(5)
            page.goto(scanrdi_url, wait_until='domcontentloaded')
            time.sleep(3)
            
    print(f"Current page: {page.url} ({page.title()})")
    page.screenshot(path='scratch/scanrdi_test_main.png', full_page=True)
    
    # Click Test Details tab
    td_tab = page.locator("a[href='#TestDetails'], a:has-text('Test Details')").first
    if td_tab.count() > 0:
        print("Clicking Test Details tab...")
        td_tab.click()
        time.sleep(2)
        page.screenshot(path='scratch/scanrdi_test_details.png', full_page=True)
        print("Saved scratch/scanrdi_test_details.png")
        
    # Also check if there is an Event History or Documents tab
    for tab_name in ['#Events', '#Documents', '#Notes', '#Samples']:
        tab = page.locator(f"a[href='{tab_name}']").first
        if tab.count() > 0:
            tab.click()
            time.sleep(1)
            page.screenshot(path=f"scratch/scanrdi_tab_{tab_name.replace('#','')}.png", full_page=True)
            
    # Save full html of test details
    with open('scratch/scanrdi_test_page.html', 'w', encoding='utf-8') as f:
        f.write(page.content())
        
    ctx.close()
