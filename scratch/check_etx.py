import os, asyncio
from playwright.async_api import async_playwright

session_dir = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'pastdue_playwright_session')

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch_persistent_context(
            user_data_dir=session_dir,
            headless=True
        )
        page = await browser.new_page()
        for etx in ['ETX-260921-0510', 'ETX-260921-0520']:
            url = f'https://etrax.eagleanalytical.com/Submission/Index?search={etx}'
            print('Navigating to:', url)
            await page.goto(url, timeout=30000)
            await page.wait_for_timeout(3000)
            links = await page.eval_on_selector_all('a', 'elements => elements.map(e => ({href: e.href, text: e.innerText.trim()}))')
            matched = [l for l in links if etx in l['text'] or etx in l['href']]
            print(f'Matched links for {etx}:', matched)
            if matched:
                full_url = matched[0]['href']
                print(f'Visiting {full_url}')
                await page.goto(full_url)
                await page.wait_for_timeout(3000)
                text = await page.inner_text('body')
                lines = [l.strip() for l in text.split('\n') if l.strip()]
                for l in lines[:40]:
                    print('   ', l)

        await browser.close()

asyncio.run(main())
