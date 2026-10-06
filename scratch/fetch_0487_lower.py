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
    
    if "/Account/Login" in page.url:
        inp = page.locator("input[name='Username'], input[id='Username']").first
        if inp.count() > 0:
            inp.fill("qchen")
            page.locator("input[type='submit'], button[type='submit']").first.click()
            time.sleep(3)
            
    # Click Test Details
    tab = page.locator("a:has-text('Test Details'), button:has-text('Test Details'), li:has-text('Test Details') a")
    if tab.count() > 0:
        tab.first.click()
        time.sleep(2)
        
    page.evaluate("window.scrollBy(0, 300)")
    time.sleep(1)
    page.screenshot(path="scratch/etrax_0487_lower_full.png")
    
    # Extract all text in the form
    text_content = page.locator(".tab-content, #TestDetails, form").inner_text()
    with open("scratch/etrax_0487_form_text.txt", "w", encoding="utf-8") as f:
        f.write(text_content)
    print("Saved form text to scratch/etrax_0487_form_text.txt")
    ctx.close()
