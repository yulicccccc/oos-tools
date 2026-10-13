import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    # Go to Equipment page
    page.goto('https://etrax.eagleanalytical.com/Equipment', wait_until='domcontentloaded')
    time.sleep(2)
    
    if "/Account/Login" in page.url:
        page.locator("input[name='Username']").first.fill('qchen')
        page.locator("input[type='submit']").first.click()
        time.sleep(5)
        page.goto('https://etrax.eagleanalytical.com/Equipment', wait_until='domcontentloaded')
        time.sleep(2)
        
    print(f"Equipment Page URL: {page.url} ({page.title()})")
    
    # Search for 1937
    search_box = page.locator("input[type='search'], input[id*='Search'], #EquipmentSearch, input.form-control").first
    if search_box.count() > 0:
        search_box.fill('1937')
        search_box.press('Enter')
        time.sleep(2)
        page.screenshot(path='scratch/equip_1937.png', full_page=True)
        
        # Print table rows
        rows = page.locator("tr").all()
        for r in rows:
            txt = r.inner_text().strip().replace('\n', ' | ')
            if '1937' in txt or 'BSC' in txt:
                print("Row 1937:", txt)
                
        # Also search for 1312
        search_box.fill('1312')
        search_box.press('Enter')
        time.sleep(2)
        page.screenshot(path='scratch/equip_1312.png', full_page=True)
        rows = page.locator("tr").all()
        for r in rows:
            txt = r.inner_text().strip().replace('\n', ' | ')
            if '1312' in txt or 'BSC' in txt:
                print("Row 1312:", txt)
                
    ctx.close()
