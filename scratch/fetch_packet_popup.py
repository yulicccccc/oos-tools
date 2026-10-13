import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')
session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto('https://etrax.eagleanalytical.com/SubmissionTest/Details/RV99MDkG5h5q830GJT5lSA__', wait_until='domcontentloaded')
    time.sleep(2)
    
    if '/Account/Login' in page.url:
        page.fill('input[name="Username"]', 'qchen')
        page.click('input[type="submit"]')
        time.sleep(5)
        page.goto('https://etrax.eagleanalytical.com/SubmissionTest/Details/RV99MDkG5h5q830GJT5lSA__', wait_until='domcontentloaded')
        time.sleep(2)

    print('On test details page:', page.url)
    link = page.locator('a:has-text("data packet")').first
    print('Found data packet link count:', link.count())
    href = link.get_attribute("href")
    full_url = f"https://etrax.eagleanalytical.com{href}"
    print(f"Resource URL: {full_url}")
    
    # Use page.evaluate to fetch binary data using browser's fetch API
    js_fetch = """
    async (url) => {
        const response = await fetch(url);
        const buffer = await response.arrayBuffer();
        return Array.from(new Uint8Array(buffer));
    }
    """
    print("Fetching resource bytes in browser context...")
    data_bytes = page.evaluate(js_fetch, full_url)
    byte_array = bytes(data_bytes)
    print(f"Fetched {len(byte_array)} bytes!")
    with open('scratch/16_sep_26_packet_2_updated.pdf', 'wb') as f:
        f.write(byte_array)
    print("Saved to scratch/16_sep_26_packet_2_updated.pdf")

    # Also fetch microscope image
    img_link = page.locator('a:has-text("Positive")').first
    if img_link.count() > 0:
        img_href = img_link.get_attribute("href")
        full_img_url = f"https://etrax.eagleanalytical.com{img_href}"
        img_bytes = page.evaluate(js_fetch, full_img_url)
        with open('scratch/img_6546.jpeg', 'wb') as f:
            f.write(bytes(img_bytes))
        print(f"Saved scratch/img_6546.jpeg ({len(img_bytes)} bytes)")

    ctx.close()
