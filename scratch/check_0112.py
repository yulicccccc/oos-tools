import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Submissions', wait_until='domcontentloaded')
    time.sleep(2)
    search_input = page.locator("input[name='TrackingNumber'], input[type='search'], input[name='search'], input[id='TrackingNumber']").first
    if search_input.count() > 0:
        search_input.fill('ETX-260901-0112')
        search_input.press('Enter')
        time.sleep(3)
        print('Page title:', page.title())
        links = page.locator("a:has-text('ETX-260901-0112')").all()
        for l in links:
            print('Found link:', l.get_attribute('href'), l.inner_text())
    page.screenshot(path='scratch/search_0112.png')
    ctx.close()
