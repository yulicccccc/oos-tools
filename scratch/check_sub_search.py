import os, time
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='domcontentloaded')
    time.sleep(2)
    if '/Account/Login' in page.url:
        page.locator("input[name='Username']").fill('qchen')
        page.locator("input[type='submit']").click()
        time.sleep(5)
        page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='domcontentloaded')
        time.sleep(2)
        
    print("Submission URL:", page.url)
    page.screenshot(path='scratch/submission_search_page.png', full_page=True)
    
    # Save HTML to inspect search form
    with open('scratch/sub_search.html', 'w', encoding='utf-8') as f:
        f.write(page.content())
        
    ctx.close()
