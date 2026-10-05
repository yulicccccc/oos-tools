import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
target_url = 'https://etrax.eagleanalytical.com/SubmissionTest/Details/OCvPHc7TodycYzvim2KtjQ__#TestDetails'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto(target_url, wait_until='domcontentloaded')
    time.sleep(2)
    
    if '/Account/Login' in page.url:
        try:
            inp = page.locator("input[name='Username'], input[id='Username']").first
            if inp.count() > 0:
                inp.fill('qchen')
                btn = page.locator("input[type='submit'], button[type='submit']").first
                if btn.count() > 0:
                    btn.click()
                    print('Clicked login submit.')
        except Exception as e:
            print('Login fill err:', e)
    
    time.sleep(4)
    print('Current URL:', page.url)
    print('Title:', page.title())
    
    # Click Test Details tab
    tab = page.locator("a:has-text('Test Details'), button:has-text('Test Details')")
    if tab.count() > 0:
        try:
            tab.first.click()
            time.sleep(2)
        except Exception as e:
            print('Tab click err:', e)
            
    page.screenshot(path='scratch/etrax_262017_micro.png')
    body_text = page.locator('body').inner_text()
    with open('scratch/etrax_262017_text.txt', 'w', encoding='utf-8') as f:
        f.write(body_text)
        
    print('=== Text excerpt ===')
    for line in body_text.splitlines():
        l = line.strip()
        if any(k in l.lower() for k in ['total cfu', 'organisms', 'colony', 'microbial identification', 'stain', 'result', 'bacillus', 'cocci', 'rods', 'gram', 'staphylococcus', 'cutibacterium', 'pseudomonas', 'micrococcus']):
            print(' ', l)
    ctx.close()
