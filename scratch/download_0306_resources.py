import os
import sys
import time
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding='utf-8')

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')), 'pastdue_playwright_session')

resources = {
    "packet_2.pdf": "https://etrax.eagleanalytical.com/Resource?id=SxVGhsA6pcJv6lxGqeYN1Afp-cyDkpPHJ5LwLb-S2bH-Mg-9ngLxYlW$OnjJPs6cCA0g8SUvLvdpRN-2tLQKN-qIYnqOnRim84-Gh4yIk6rOd6boaGijZZ-QlRAp6Leb",
    "microscope_0306.jpeg": "https://etrax.eagleanalytical.com/Resource?id=SxVGhsA6pcJv6lxGqeYN1Afp-cyDkpPHJ5LwLb-S2bH-Mg-9ngLxYlW$OnjJPs6c1xaBs$FL$9Pog6vDLKFpFJp305mWTs1tbqNUSVYlFfw_"
}

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(session_dir, headless=True)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    
    # Ensure logged in
    page.goto("https://etrax.eagleanalytical.com/Submission/Details/j$B8MDDTkRSGXhr3At81EQ__", wait_until='domcontentloaded')
    time.sleep(2)
    
    for fname, url in resources.items():
        print(f"Downloading {fname} from {url} ...")
        resp = page.request.get(url)
        print(f"Status: {resp.status}")
        save_path = os.path.join(os.getcwd(), "scratch", fname)
        with open(save_path, "wb") as f:
            f.write(resp.body())
        print(f"Saved: {save_path} ({len(resp.body())} bytes)")

    ctx.close()
    print("Done downloading resources.")
