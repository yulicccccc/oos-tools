import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Account/Login?ReturnUrl=%2FSubmission%2FIndex', wait_until='domcontentloaded')
    time.sleep(2)
    inp = page.locator("input[name='Username']").first
    if inp.count() > 0:
        inp.fill('qchen')
        btn = page.locator("input[type='submit']").first
        if btn.count() > 0:
            btn.click()
            print('Submitted qchen')
    time.sleep(6)
    print('Current URL:', page.url)
    print('Current Title:', page.title())
    ctx.close()
