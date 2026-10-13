import os, time, sys
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')
session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

links = {
    'ETX-260921-0148': 'https://etrax.eagleanalytical.com/Submission/Details/08VuIoFsBaXGehnOxJ3PMQ__',
    'ETX-261005-0773': 'https://etrax.eagleanalytical.com/Submission/Details/e-zI5w7ZUuRgQ9-SJ2n7DA__',
    'ETX-260929-0335': 'https://etrax.eagleanalytical.com/Submission/Details/QchKBV1D2VrN7y3VezYRkg__',
    'ETX-260929-0341': 'https://etrax.eagleanalytical.com/Submission/Details/gKex8LPhpMsQyRuqi%24ua4g__',
    'ETX-260929-0344': 'https://etrax.eagleanalytical.com/Submission/Details/iCnj0gFFyEtIFNZbfpq5qg__',
}

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    for s_id, url in links.items():
        page.goto(url, wait_until='domcontentloaded')
        time.sleep(2)
        # Check title and displayed submission #
        sub_text = page.locator("h2, .submission-id, a:has-text('ETX-')").all_text_contents()
        sub_no = [t for t in sub_text if 'ETX-' in t]
        print(f"{s_id}: URL={url}")
        print(f"   Title: {page.title()}")
        print(f"   Found ETX texts: {sub_no[:3]}")
    ctx.close()
