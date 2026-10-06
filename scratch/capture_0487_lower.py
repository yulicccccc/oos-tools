import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

target_url = 'https://etrax.eagleanalytical.com/SubmissionTest/Details/fd3G2StZClcy1TP2ES6BLw__#TestDetails'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto(target_url, wait_until='domcontentloaded')
    time.sleep(2)
    
    tab = page.locator("a:has-text('Test Details'), button:has-text('Test Details'), li:has-text('Test Details') a")
    if tab.count() > 0:
        tab.first.click()
        time.sleep(1)
        
    # Scroll down 300px
    page.evaluate("window.scrollBy(0, 350)")
    time.sleep(1)
    page.screenshot(path="scratch/etrax_0487_scroll.png")
    print("Saved screenshot to scratch/etrax_0487_scroll.png")
    ctx.close()
