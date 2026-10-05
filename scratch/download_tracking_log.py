import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'pastdue_playwright_session')
url = 'https://eagleanalytical-my.sharepoint.com/:x:/r/personal/rseymour_eagleanalytical_com/_layouts/15/Doc.aspx?sourcedoc=%7B8A575D4F-67AB-4253-B1E2-08A540577BCB%7D&file=Sterile%20Lab%20-%20OOS%20Tracking%20Log.xlsx&fromShare=true&action=default&mobileredirect=true'

print("Launching browser with persistent session...")
with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=session_dir,
        headless=True,
        downloads_path='scratch'
    )
    page = context.pages[0] if context.pages else context.new_page()
    page.set_viewport_size({"width": 1920, "height": 1080})
    print("Navigating to SharePoint Excel...")
    page.goto(url, wait_until='domcontentloaded')
    
    # Wait for Excel Online to load
    print("Waiting for File button...")
    page.wait_for_selector('button:has-text("File"), [aria-label="File"]', timeout=45000)
    time.sleep(3)
    
    file_button = page.locator('button:has-text("File"), [aria-label="File"]').first
    print("Clicking File button...")
    file_button.click()
    time.sleep(2)
    page.screenshot(path='scratch/step1_file_menu.png')
    
    # Check "Export" or "Create a Copy"
    export_item = page.locator('div[role="menuitem"]:has-text("Export"), button:has-text("Export"), span:has-text("Export")').first
    create_copy_item = page.locator('div[role="menuitem"]:has-text("Create a Copy"), button:has-text("Create a Copy"), span:has-text("Create a Copy")').first
    
    if export_item.is_visible():
        print("Clicking Export...")
        export_item.click()
        time.sleep(2)
        page.screenshot(path='scratch/step2_export.png')
    elif create_copy_item.is_visible():
        print("Clicking Create a Copy...")
        create_copy_item.click()
        time.sleep(2)
        page.screenshot(path='scratch/step2_create_copy.png')
        
    # Check for "Download a Copy"
    download_btn = page.locator('text="Download a Copy", button:has-text("Download a Copy"), div[role="menuitem"]:has-text("Download a Copy")').first
    if download_btn.is_visible():
        print("Found 'Download a Copy'! Triggering download...")
        with page.expect_download(timeout=30000) as download_info:
            download_btn.click()
        download = download_info.value
        target = os.path.join('scratch', 'Sterile Lab - OOS Tracking Log.xlsx')
        download.save_as(target)
        print(f"SUCCESS! Downloaded to {target}, size = {os.path.getsize(target)} bytes")
    else:
        print("Could not find 'Download a Copy' yet. Checking all visible text...")
        page.screenshot(path='scratch/step3_state.png')
        
    context.close()
