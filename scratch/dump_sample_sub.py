import os
from playwright.sync_api import sync_playwright

user_data_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')
url_sample = "https://etrax.eagleanalytical.com/Submission/Details/fZlFl0vWDpqR-VWBsez7IA__"

with sync_playwright() as p:
    browser = p.chromium.launch_persistent_context(user_data_dir=user_data_dir, headless=True)
    page = browser.new_page()
    page.goto(url_sample, wait_until='domcontentloaded')
    page.wait_for_timeout(3000)
    
    if "Log in" in page.title():
        page.fill("input[name='Username']", "qchen")
        page.click("input[type='submit'], button[type='submit']")
        page.wait_for_timeout(3000)
        
    print("Page URL:", page.url)
    print("Page Title:", page.title())
    
    text = page.inner_text("body")
    with open("scratch/sample_submission_261002_0288.txt", "w", encoding="utf-8") as f:
        f.write(text)
    page.screenshot(path="scratch/sample_submission_261002_0288.png", full_page=True)
    
    # Also find all test links in this submission
    links = page.evaluate('''() => {
        return Array.from(document.querySelectorAll('a')).map(a => ({text: a.innerText.trim(), href: a.href})).filter(a => a.href.includes('SubmissionTest'));
    }''')
    print("SubmissionTest links:", links)
    
    browser.close()
