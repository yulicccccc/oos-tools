import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
scanrdi_url = 'https://etrax.eagleanalytical.com/SubmissionTest/Details/zFdCKB28Sd0gPid31SfEnA__'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto(scanrdi_url, wait_until='domcontentloaded')
    time.sleep(2)
    
    # Click Worksheets tab
    ws_tab = page.locator("a[href='#Worksheets'], a:has-text('Worksheets')").first
    if ws_tab.count() > 0:
        print("Clicking Worksheets tab...")
        ws_tab.click()
        time.sleep(2)
        page.screenshot(path='scratch/scanrdi_worksheets_tab.png', full_page=True)
        print("Saved scratch/scanrdi_worksheets_tab.png")
        
    ctx.close()
