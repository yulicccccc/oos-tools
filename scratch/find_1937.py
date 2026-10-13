import os, time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/Equipment')
    time.sleep(2)
    if '/Account/Login' in page.url:
        page.locator("input[name='Username']").fill('qchen')
        page.locator("input[type='submit']").click()
        time.sleep(5)
        page.goto('https://etrax.eagleanalytical.com/Equipment')
        time.sleep(2)
        
    # Click the Search button/link
    btn = page.locator("a:has-text('Search'), span:has-text('Search')").first
    print("Search button count:", btn.count())
    btn.click()
    time.sleep(1)
    page.screenshot(path='scratch/equip_search_expanded.png', full_page=True)
    
    # Fill Equipment ID
    eq_input = page.locator("#EquipmentId, input[name='EquipmentId']").first
    print("EquipmentId input visible:", eq_input.is_visible())
    if eq_input.count() > 0:
        eq_input.fill("1937")
        submit_btn = page.locator("input[type='submit'][value='Search'], button:has-text('Search')").first
        submit_btn.click()
        time.sleep(3)
        page.screenshot(path='scratch/equip_1937_final.png', full_page=True)
        rows = page.locator("tr").all()
        for r in rows:
            print("ROW:", r.inner_text().strip().replace('\n', ' | '))
    ctx.close()
