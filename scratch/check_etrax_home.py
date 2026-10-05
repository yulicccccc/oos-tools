import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'pastdue_playwright_session')

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=session_dir,
        headless=True
    )
    page = context.pages[0] if context.pages else context.new_page()
    page.set_viewport_size({"width": 1920, "height": 1080})
    
    page.goto('https://etrax.eagleanalytical.com/')
    page.wait_for_timeout(4000)
    print("URL:", page.url)
    print("Title:", page.title())
    page.screenshot(path='scratch/etrax_home.png')
    
    links = page.locator('a').all()
    print("Top navigation links:")
    for a in links[:25]:
        text = a.inner_text().strip()
        href = a.get_attribute('href')
        if text and href:
            print(f"  {text} -> {href}")
            
    context.close()
