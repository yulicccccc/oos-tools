import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Equipment', wait_until='domcontentloaded')
    time.sleep(2)
    
    # Click Search header to expand if collapsed
    search_hdr = page.locator("a:has-text('Search'), div:has-text('Search')").first
    search_hdr.click()
    time.sleep(1)
    
    # Fill Equipment ID
    eq_inp = page.locator("input[id='EquipmentId'], input[name='EquipmentId']").first
    print("Found eq_inp:", eq_inp.count(), "visible:", eq_inp.is_visible())
    if eq_inp.count() > 0:
        eq_inp.fill("1937")
        # Click Search button
        btn = page.locator("input[value='Search'], button:has-text('Search')").first
        btn.click()
        time.sleep(2)
        page.screenshot(path='scratch/equip_1937_result.png', full_page=True)
        
        # Print table
        rows = page.locator("tr").all()
        for r in rows:
            print("ROW:", r.inner_text().strip().replace('\n', ' | '))
            
    ctx.close()
