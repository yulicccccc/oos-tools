from playwright.sync_api import sync_playwright
import os

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/SubmissionTest/Details/qnwcLQO5BWeJhWCBG7jl8Q__#TestDetails')
    page.wait_for_timeout(2000)
    
    # Click Test Details
    tab = page.locator("a:has-text('Test Details'), button:has-text('Test Details'), li:has-text('Test Details') a")
    if tab.count() > 0:
        tab.first.click()
        page.wait_for_timeout(1000)
        
    print('All inputs:')
    inputs = page.locator('input')
    for i in range(inputs.count()):
        inp = inputs.nth(i)
        v = inp.input_value()
        name = inp.get_attribute('name') or inp.get_attribute('id')
        if v:
            print(f'  {name} = {repr(v)}')
    ctx.close()
