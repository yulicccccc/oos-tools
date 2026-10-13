import os
import time
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
    print('URL:', page.url)
    page.screenshot(path='scratch/equip_logged_in.png', full_page=True)
    
    # Fill search or filter
    # On EagleTrax equipment page, look at all inputs
    inputs = page.locator("input")
    print(f"Total inputs: {inputs.count()}")
    for i in range(inputs.count()):
        inp = inputs.nth(i)
        if inp.is_visible():
            print(f"Visible Input {i}: id={inp.get_attribute('id')}, name={inp.get_attribute('name')}, type={inp.get_attribute('type')}")
            
    # Search table for 1937
    page.goto('https://etrax.eagleanalytical.com/Equipment?EquipmentId=E001937')
    time.sleep(2)
    page.screenshot(path='scratch/equip_1937_query.png', full_page=True)
    
    ctx.close()
