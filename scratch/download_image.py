import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
img_url = 'https://etrax.eagleanalytical.com/Resource?id=SxVGhsA6pcJv6lxGqeYN1Kv0Z7-2epfDss3ygfZumDWxFuNJgC1QpGZ0UGuWMXafJOw2gNapRb-7X77yXk7fwquR9MKDreIaxjJZNdlzxPc_'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/SubmissionTest/Details/zFdCKB28Sd0gPid31SfEnA__', wait_until='domcontentloaded')
    time.sleep(2)
    
    # Download image
    with page.expect_download() as download_info:
        page.goto(img_url)
    download = download_info.value
    download.save_as('scratch/image_4_isolated_rod.jpeg')
    print("Saved scratch/image_4_isolated_rod.jpeg!")
    ctx.close()
