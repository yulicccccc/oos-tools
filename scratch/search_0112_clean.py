import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='domcontentloaded')
    time.sleep(3)
    print('Current URL:', page.url)
    print('Title:', page.title())
    if '/Account/Login' in page.url:
        inp = page.locator("input[name='Username'], input[id='Username']").first
        if inp.count() > 0:
            inp.fill('qchen')
            btn = page.locator("input[type='submit'], button[type='submit']").first
            btn.click()
            time.sleep(4)
    # Search for 260901-0112
    search_inp = page.locator("input[type='search'], input[name='search'], input[id*='search'], input[class*='search'], input[placeholder*='Search']").first
    if search_inp.count() > 0:
        search_inp.fill('260901-0112')
        search_inp.press('Enter')
        time.sleep(3)
        print('Search submitted')
        rows = page.locator('tr').all()
        for r in rows[:10]:
            print('Row:', r.inner_text().strip().replace('\n', ' | '))
    page.screenshot(path='scratch/search_submission_0112.png')
    ctx.close()
