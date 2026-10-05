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
    
    # Let's open EagleTrax Submission Search or Test Search
    print("Navigating to EagleTrax...")
    page.goto('https://etrax.eagleanalytical.com/Submission/Search')
    page.wait_for_timeout(5000)
    print("URL:", page.url)
    print("Title:", page.title())
    page.screenshot(path='scratch/etrax_submission_search.png')
    
    # Check if we can search by Account Number E19193 or Client Name
    # Print visible inputs
    inputs = page.locator('input').all()
    print(f"Found {len(inputs)} inputs")
    for inp in inputs[:15]:
        print(f"Input: id='{inp.get_attribute('id')}', name='{inp.get_attribute('name')}', placeholder='{inp.get_attribute('placeholder')}'")
        
    context.close()
