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
    print("Navigating...")
    page.goto(url)
    print("Waiting 15 seconds for Excel iframe...")
    page.wait_for_timeout(15000)
    
    frame = page.frame(name='WacFrame_Excel_0')
    if not frame:
        print("Error: WacFrame_Excel_0 not found!")
        context.close()
        exit(1)
        
    print("Found WacFrame_Excel_0!")
    
    # 1. First, click on "Celsis Sterility OOS" tab to inspect it
    celsis_tab = frame.locator('text="Celsis Sterility OOS"').first
    if celsis_tab.is_visible():
        print("Found Celsis Sterility OOS tab! Clicking it...")
        celsis_tab.click()
        time.sleep(3)
        page.screenshot(path='scratch/celsis_tab_view.png')
        print("Captured celsis_tab_view.png")

    # 2. Try File -> Export / Create a Copy -> Download a Copy
    file_btn = frame.locator('#FileMenuLauncher, button:has-text("File"), [aria-label="File"]').first
    if file_btn.is_visible():
        print("Clicking File menu inside iframe...")
        file_btn.click()
        time.sleep(2)
        page.screenshot(path='scratch/file_menu_inside_frame.png')
        
        # Check Export or Create a Copy
        export_btn = frame.locator('text="Export", button:has-text("Export"), div[role="menuitem"]:has-text("Export")').first
        create_copy_btn = frame.locator('text="Create a Copy", button:has-text("Create a Copy"), div[role="menuitem"]:has-text("Create a Copy")').first
        
        target_menu = None
        if export_btn.is_visible():
            print("Found Export button!")
            target_menu = export_btn
        elif create_copy_btn.is_visible():
            print("Found Create a Copy button!")
            target_menu = create_copy_btn
            
        if target_menu:
            target_menu.click()
            time.sleep(2)
            page.screenshot(path='scratch/export_suboptions.png')
            
            # Now find "Download a Copy"
            dl_btn = frame.locator('text="Download a Copy", button:has-text("Download a Copy"), div:has-text("Download a Copy")').first
            if dl_btn.is_visible():
                print("Found 'Download a Copy'! Expecting download...")
                try:
                    with page.expect_download(timeout=20000) as dl_info:
                        dl_btn.click()
                    dl = dl_info.value
                    save_path = os.path.join('scratch', 'Sterile Lab - OOS Tracking Log.xlsx')
                    dl.save_as(save_path)
                    print(f"SUCCESSFULLY DOWNLOADED to {save_path}, size={os.path.getsize(save_path)} bytes")
                except Exception as e:
                    print(f"Download trigger failed or timed out: {e}")
            else:
                print("Could not find 'Download a Copy' inside sub-menu.")
        else:
            print("Neither Export nor Create a Copy was visible.")
    else:
        print("File button not visible inside iframe.")
        
    context.close()
    print("Done!")
