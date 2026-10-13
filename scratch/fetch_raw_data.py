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
    
    # Listen for popup or download
    btn = page.locator("#SubmissionTestRawDataButton").first
    print("Found button:", btn.count())
    if btn.count() > 0:
        with page.expect_popup() as popup_info:
            btn.click()
        popup = popup_info.value
        popup.wait_for_load_state('domcontentloaded')
        time.sleep(3)
        print("Popup URL:", popup.url)
        print("Popup Title:", popup.title())
        popup.screenshot(path='scratch/raw_data_report.png', full_page=True)
        with open('scratch/raw_data_report.html', 'w', encoding='utf-8') as f:
            f.write(popup.content())
        print("Saved raw data report!")
        
    ctx.close()
