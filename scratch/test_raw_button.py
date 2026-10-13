import os
import time
from playwright.sync_api import sync_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
scanrdi_url = 'https://etrax.eagleanalytical.com/SubmissionTest/Details/zFdCKB28Sd0gPid31SfEnA__'

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    page.on("request", lambda r: print(f"REQ: {r.method} {r.url}") if "Report" in r.url or "Raw" in r.url or "PDF" in r.url else None)
    page.on("response", lambda r: print(f"RESP: {r.status} {r.url}") if "Report" in r.url or "Raw" in r.url or "PDF" in r.url else None)
    
    page.goto(scanrdi_url, wait_until='domcontentloaded')
    time.sleep(2)
    
    if "/Account/Login" in page.url or "Login" in page.title():
        inp = page.locator("input[name='Username']").first
        if inp.count() > 0:
            inp.fill('qchen')
            page.locator("input[type='submit']").first.click()
            time.sleep(5)
            page.goto(scanrdi_url, wait_until='domcontentloaded')
            time.sleep(2)
            
    td_tab = page.locator("a[href='#TestDetails']").first
    if td_tab.count() > 0:
        td_tab.click()
        time.sleep(1)
        
    btn = page.locator("#SubmissionTestRawDataButton").first
    if btn.count() > 0:
        print("Clicking #SubmissionTestRawDataButton...")
        btn.click()
        time.sleep(5)
        
    print("Open pages count:", len(ctx.pages))
    for i, p_item in enumerate(ctx.pages):
        print(f"Page {i}: {p_item.url} ({p_item.title()})")
        p_item.screenshot(path=f"scratch/page_{i}.png", full_page=True)
        
    ctx.close()
