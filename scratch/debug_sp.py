import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'pastdue_playwright_session')
url = 'https://eagleanalytical-my.sharepoint.com/:x:/r/personal/rseymour_eagleanalytical_com/_layouts/15/Doc.aspx?sourcedoc=%7B8A575D4F-67AB-4253-B1E2-08A540577BCB%7D&file=Sterile%20Lab%20-%20OOS%20Tracking%20Log.xlsx&fromShare=true&action=default&mobileredirect=true'

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=session_dir,
        headless=True
    )
    page = context.pages[0] if context.pages else context.new_page()
    page.set_viewport_size({"width": 1920, "height": 1080})
    print("Navigating...")
    page.goto(url)
    print("Waiting 15 seconds...")
    page.wait_for_timeout(15000)
    page.screenshot(path='scratch/current_page.png')
    print("Title:", page.title())
    print("URL:", page.url)
    
    # Print all frames
    for i, frame in enumerate(page.frames):
        print(f"Frame {i}: name='{frame.name}', url='{frame.url[:80]}'")
        
    context.close()
