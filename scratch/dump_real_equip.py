import os, time
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Equipment', wait_until='domcontentloaded')
    time.sleep(2)
    if '/Account/Login' in page.url:
        page.locator("input[name='Username']").fill('qchen')
        page.locator("input[type='submit']").click()
        time.sleep(5)
        page.goto('https://etrax.eagleanalytical.com/Equipment', wait_until='domcontentloaded')
        time.sleep(3)
        
    content = page.content()
    with open('scratch/equip_real_page.html', 'w', encoding='utf-8') as f:
        f.write(content)
        
    soup = BeautifulSoup(content, 'html.parser')
    for form in soup.find_all('form'):
        print("FORM:", form.get('action'), form.get('id'), form.get('class'))
        for inp in form.find_all(['input', 'select', 'button']):
            print("  INP:", inp.name, inp.get('id'), inp.get('name'), inp.get('type'), inp.get('value'))
    ctx.close()
