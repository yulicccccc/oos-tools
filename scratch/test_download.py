import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'pastdue_playwright_session')
url = 'https://eagleanalytical-my.sharepoint.com/:x:/r/personal/rseymour_eagleanalytical_com/_layouts/15/Doc.aspx?sourcedoc=%7B8A575D4F-67AB-4253-B1E2-08A540577BCB%7D&file=Sterile%20Lab%20-%20OOS%20Tracking%20Log.xlsx&fromShare=true&action=default&mobileredirect=true'

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=session_dir,
        headless=True,
        accept_downloads=True,
        downloads_path='scratch'
    )
    page = context.pages[0] if context.pages else context.new_page()
    page.set_viewport_size({"width": 1920, "height": 1080})
    page.goto(url)
    page.wait_for_timeout(15000)
    
    frame = page.frame(name='WacFrame_Excel_0')
    if not frame:
        print("WacFrame_Excel_0 not found")
        context.close()
        exit(1)
        
    # Click File
    file_btn = frame.locator('#FileMenuLauncher, button:has-text("File"), [aria-label="File"]').first
    file_btn.click()
    time.sleep(1)
    
    copy_item = frame.get_by_text("Create a Copy", exact=True)
    if copy_item.is_visible():
        print("Found Create a Copy, clicking...")
        copy_item.click()
        time.sleep(1)
        page.screenshot(path='scratch/sub_create_copy.png')
        
        # Now check what text is visible in the flyout menu
        dl_item = frame.get_by_text("Download a Copy")
        if dl_item.is_visible():
            print("Found Download a Copy! Clicking...")
            with page.expect_download(timeout=30000) as dl_info:
                dl_item.click()
            dl = dl_info.value
            dest = os.path.join('scratch', 'Sterile Lab - OOS Tracking Log.xlsx')
            dl.save_as(dest)
            print(f"SUCCESSFULLY DOWNLOADED: {dest}, size = {os.path.getsize(dest)} bytes")
        else:
            print("Download a Copy not found inside Create a Copy menu.")
    else:
        print("Create a Copy not visible.")

    context.close()
