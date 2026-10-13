import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
scan_url = "https://etrax.eagleanalytical.com/SubmissionTest/Details/RV99MDkG5h5q830GJT5lSA__"

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    print(f"Navigating to Scan RDI test URL: {scan_url}")
    page.goto(scan_url, wait_until='domcontentloaded')
    time.sleep(3)
    
    # 1. Click Test Details tab
    test_details_tab = page.locator("a:has-text('Test Details')").first
    if test_details_tab.count() > 0:
        print("Clicking Test Details tab...")
        test_details_tab.click()
        time.sleep(2)
        page.screenshot(path='scratch/sample_0306_tab_testdetails.png', full_page=True)
        print("Saved scratch/sample_0306_tab_testdetails.png")
        with open('scratch/sample_0306_tab_testdetails.html', 'w', encoding='utf-8') as f:
            f.write(page.content())
            
    # 2. Click Raw Data Report button if available
    raw_btn = page.locator("a:has-text('Raw Data Report'), button:has-text('Raw Data Report')").first
    if raw_btn.count() > 0:
        print("Raw Data Report button found, clicking to see target...")
        href = raw_btn.get_attribute("href")
        print(f"Raw data report href: {href}")
        # If it's a popup or new page:
        with page.expect_popup() as popup_info:
            raw_btn.click()
        popup = popup_info.value
        popup.wait_for_load_state("domcontentloaded")
        time.sleep(3)
        popup.screenshot(path='scratch/sample_0306_raw_report.png', full_page=True)
        print("Saved scratch/sample_0306_raw_report.png")
        with open('scratch/sample_0306_raw_report.html', 'w', encoding='utf-8') as f:
            f.write(popup.content())
        popup.close()

    # 3. Download attached documents
    doc_links = page.locator("table tr td a[href*='Resource']")
    print(f"\nFound {doc_links.count()} resource links in document table:")
    for i in range(doc_links.count()):
        dl = doc_links.nth(i)
        txt = dl.inner_text().strip()
        href = dl.get_attribute("href")
        print(f"Document {i+1}: '{txt}' -> {href}")
        
    # Download each document via page.expect_download or request
    for i in range(doc_links.count()):
        dl = doc_links.nth(i)
        txt = dl.inner_text().strip()
        if not txt:
            # maybe it's on the row
            row = dl.locator("xpath=ancestor::tr")
            txt = row.locator("td").first.inner_text().strip()
        print(f"\nDownloading document {i+1}: {txt} ...")
        try:
            with page.expect_download(timeout=10000) as download_info:
                dl.click()
            download = download_info.value
            save_path = os.path.join(os.getcwd(), 'scratch', f"doc_{i+1}_{txt}")
            download.save_as(save_path)
            print(f"Downloaded to: {save_path}")
        except Exception as e:
            print(f"Could not download directly via click: {e}")
            # Try navigating to href
            href = dl.get_attribute("href")
            if href:
                full_href = f"https://etrax.eagleanalytical.com{href}" if href.startswith("/") else href
                print(f"Trying direct fetch: {full_href}")
                r = page.request.get(full_href)
                save_path = os.path.join(os.getcwd(), 'scratch', f"doc_{i+1}_{txt}")
                with open(save_path, 'wb') as df:
                    df.write(r.body())
                print(f"Saved via request: {save_path}")

    ctx.close()
    print("Done fetching details and documents.")
