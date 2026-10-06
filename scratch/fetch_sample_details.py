import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
target_url = 'https://etrax.eagleanalytical.com/Submission/Details/OIX27OGNKb64xLa70p0RRQ__'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto(target_url, wait_until='domcontentloaded')
    time.sleep(2)
    if '/Account/Login' in page.url:
        inp = page.locator("input[name='Username'], input[id='Username']").first
        if inp.count() > 0:
            inp.fill('qchen')
            btn = page.locator("input[type='submit'], button[type='submit']").first
            if btn.count() > 0:
                btn.click()
                print('Clicked login submit.')
                time.sleep(5)
    print('URL after login:', page.url)
    page.goto(target_url, wait_until='domcontentloaded')
    time.sleep(3)
    body_text = page.locator('body').inner_text()
    with open('scratch/etrax_sample_details.txt', 'w', encoding='utf-8') as f:
        f.write(body_text)
    print('Length:', len(body_text))
    for line in body_text.splitlines():
        if any(k in line.lower() for k in ['analyst', 'assigned', 'sterility', 'celsis', 'es', 'aliquot']):
            print(' ', line.strip())
    ctx.close()
