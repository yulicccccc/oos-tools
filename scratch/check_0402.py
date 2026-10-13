import os, sys, time
from playwright.sync_api import sync_playwright

profile_dir = os.path.expandvars(r'%LOCALAPPDATA%\pastdue_playwright_session')

with sync_playwright() as p:
    browser = p.chromium.launch_persistent_context(user_data_dir=profile_dir, headless=True)
    page = browser.new_page()
    print("Navigating to Submission search...")
    page.goto('https://etrax.eagleanalytical.com/Submission', wait_until='networkidle')
    
    # Search ETX-260923-0402
    search_inp = page.locator('input[type="search"], input[placeholder*="Search"]').first
    if search_inp.count() > 0:
        search_inp.fill('ETX-260923-0402')
        search_inp.press('Enter')
        time.sleep(3)
        
        rows = page.locator('tr').all()
        print(f"Rows count: {len(rows)}")
        sub_link = None
        for r in rows:
            txt = r.inner_text().strip().replace('\n', ' | ')
            if '260923-0402' in txt:
                print("Found row:", txt)
                for a in r.locator('a').all():
                    href = a.get_attribute('href')
                    print("  Link:", href, a.inner_text())
                    if href and '/Submission/Details/' in href and not sub_link:
                        sub_link = href
        
        if sub_link:
            target_url = 'https://etrax.eagleanalytical.com' + sub_link
            print("Navigating to submission details:", target_url)
            page.goto(target_url, wait_until='networkidle')
            time.sleep(3)
            page.screenshot(path='scratch/sub_0402_details.png', full_page=True)
            with open('scratch/sub_0402_page.html', 'w', encoding='utf-8') as f:
                f.write(page.content())
            print("Page title:", page.title())
            body_txt = page.locator('body').inner_text()
            print("Submission details body text snippet:")
            print('\n'.join([line for line in body_txt.split('\n') if any(k in line for k in ['Client', 'Sample', 'Lot', 'Test', 'Date', '115', '114', 'Air', 'Surface', 'Analyst', 'Status'])][:30]))
            
    browser.close()
