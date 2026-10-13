import os, time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Account/Login', wait_until='domcontentloaded')
    time.sleep(2)
    inp = page.locator("input[name='Username'], input[id='Username']").first
    if inp.count() > 0:
        inp.fill('qchen')
        btn = page.locator("input[type='submit'], button[type='submit']").first
        btn.click()
        print('Filled qchen and clicked submit')
    time.sleep(5)
    print('Current URL:', page.url)
    page.screenshot(path='scratch/login_step2.png')
    ctx.close()
