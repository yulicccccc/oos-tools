import os, time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Submission/Details/e-zI5w7ZUuRgQ9-SJ2n7DA__')
    time.sleep(2)
    
    doc_tab = page.locator("a[href='#Documents']").first
    if doc_tab.count() > 0:
        doc_tab.click()
        time.sleep(2)
        page.screenshot(path='scratch/docs_0773.png', full_page=True)
        
    ctx.close()
