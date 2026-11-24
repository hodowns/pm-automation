#!/usr/bin/env python3
"""
Quick SDC revenue reader — opens the project, finds Financial Summary fields,
and prints revenue to-date vs total revenue.

Usage:
    python read_project_revenue.py F59190
"""

import asyncio
import os
import re
import sys

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

SDC_URL = "https://csone.lightning.force.com/"
PROFILE_DIR = os.path.normpath(
    os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "sdc_revenue_profile",
    )
)


async def launch_browser(playwright):
    os.makedirs(PROFILE_DIR, exist_ok=True)
    context = await playwright.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        channel="chrome",
        headless=False,
        slow_mo=200,
    )
    page = await context.new_page()
    await page.goto(SDC_URL)

    print("Waiting for SDC login (up to 3 min)...")
    deadline = asyncio.get_event_loop().time() + 180
    while asyncio.get_event_loop().time() < deadline:
        await asyncio.sleep(2)
        url = page.url
        if ("force.com" in url or "salesforce.com" in url) and "login" not in url:
            print("  Logged in.")
            return context, page
    raise RuntimeError("Login timed out.")


async def find_and_open_project(page, pid: str) -> bool:
    print(f"Searching for {pid}...")
    try:
        search = page.locator(".slds-global-header input[type='search']")
        await search.wait_for(state="visible", timeout=8000)
        await search.click()
    except Exception:
        vp = page.viewport_size or {"width": 1280, "height": 720}
        x = int(vp["width"] * 0.48)
        for _ in range(5):
            await page.mouse.click(x, 25)
            await asyncio.sleep(1)
            tag = await page.evaluate("document.activeElement.tagName")
            if tag not in ("BODY", "HTML"):
                break
            await asyncio.sleep(3)

    await asyncio.sleep(0.3)
    await page.keyboard.type(pid)
    await asyncio.sleep(0.5)
    await page.keyboard.press("Enter")
    await asyncio.sleep(6)

    context = page.context
    project_link = None
    deadline = asyncio.get_event_loop().time() + 20
    while asyncio.get_event_loop().time() < deadline:
        await asyncio.sleep(2)
        for p in context.pages:
            try:
                links = p.locator(f"a:has-text('{pid}')")
                count = await links.count()
                for i in range(count):
                    link = links.nth(i)
                    try:
                        text = (await link.inner_text(timeout=500)).strip()
                        if text.startswith(f"{pid} - "):
                            continue
                        if await link.is_visible(timeout=500):
                            project_link = link
                            print(f"  Found link: '{text}'")
                            break
                    except Exception:
                        continue
                if project_link:
                    break
            except Exception:
                continue
        if project_link:
            break

    if project_link is None:
        print(f"  ERROR: No project link found for {pid}.")
        return False

    await project_link.click()
    await asyncio.sleep(5)
    print(f"  Page URL: {page.url[:100]}")
    return True


async def click_tab(page, tab_name: str) -> bool:
    """Click a named tab in the project record. Returns True if successful."""
    selectors = [
        f"a[title='{tab_name}']",
        f"a:has-text('{tab_name}')",
        f"span:has-text('{tab_name}')",
    ]
    for sel in selectors:
        try:
            el = page.locator(sel).first
            if await el.is_visible(timeout=3000):
                await el.click()
                await asyncio.sleep(3)
                print(f"  Clicked tab: {tab_name}")
                return True
        except Exception:
            continue
    print(f"  Tab not found: {tab_name}")
    return False


async def extract_revenue(page) -> dict:
    """Scrape financial fields from the current page. Returns a dict of label->value."""
    results = {}

    # Strategy 1: Salesforce standard field layout  (dt/dd or label+value pairs)
    raw = await page.evaluate("""() => {
        const pairs = {};
        // slds-form__item rows
        document.querySelectorAll('.slds-form__item').forEach(item => {
            const label = item.querySelector('.slds-form-element__label, label');
            const value = item.querySelector('.slds-form-element__static, .slds-form-element__control');
            if (label && value) {
                pairs[label.innerText.trim()] = value.innerText.trim();
            }
        });
        // generic dt/dd
        document.querySelectorAll('dt').forEach(dt => {
            const dd = dt.nextElementSibling;
            if (dd && dd.tagName === 'DD') {
                pairs[dt.innerText.trim()] = dd.innerText.trim();
            }
        });
        return pairs;
    }""")

    if raw:
        results.update(raw)

    # Strategy 2: Any element whose text contains "revenue" or "earned"
    revenue_els = await page.evaluate("""() => {
        const hits = {};
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
        let node;
        while ((node = walker.nextNode())) {
            const text = node.textContent.trim();
            if (/revenue|earned|billed|invoiced/i.test(text) && text.length < 120) {
                // grab the next sibling text as the value
                const parent = node.parentElement;
                const sib = parent ? parent.nextElementSibling : null;
                const val = sib ? sib.innerText.trim() : '';
                if (text && val) hits[text] = val;
            }
        }
        return hits;
    }""")
    if revenue_els:
        results.update(revenue_els)

    return results


def filter_revenue_fields(data: dict) -> dict:
    """Keep only fields that look revenue-related."""
    keywords = re.compile(r"revenue|earned|billed|invoiced|recognized|forecast|total|ytd|mtd|to.date", re.I)
    return {k: v for k, v in data.items() if keywords.search(k) and v.strip()}


async def main(pid: str):
    async with async_playwright() as pw:
        context, page = await launch_browser(pw)
        try:
            ok = await find_and_open_project(page, pid)
            if not ok:
                print("Could not open project.")
                return

            await asyncio.sleep(3)

            # Take a screenshot so we can see what loaded
            ss_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{pid}_revenue_check.png")

            # Try Financial Summary tab first, then Financials, then Details
            found_tab = False
            for tab in ["Financial Summary", "Financials", "Financial", "Details", "Summary"]:
                if await click_tab(page, tab):
                    found_tab = True
                    await asyncio.sleep(2)
                    break

            if not found_tab:
                print("  Could not find a Financials tab — reading current page.")

            # Take screenshot for reference
            await page.screenshot(path=ss_path)
            print(f"\nScreenshot saved: {ss_path}")

            # Extract all field data
            all_fields = await extract_revenue(page)
            revenue_fields = filter_revenue_fields(all_fields)

            print(f"\n{'='*60}")
            print(f"  Revenue data for {pid}")
            print(f"{'='*60}")

            if revenue_fields:
                max_key = max(len(k) for k in revenue_fields)
                for label, value in sorted(revenue_fields.items()):
                    print(f"  {label:<{max_key}}  |  {value}")
            else:
                print("  No revenue fields found on this tab.")
                print(f"\n  All fields found ({len(all_fields)} total):")
                for label, value in list(all_fields.items())[:30]:
                    print(f"    {label}: {value}")

            print(f"{'='*60}\n")

        finally:
            input("Press Enter to close the browser...")
            await context.close()


if __name__ == "__main__":
    pid = sys.argv[1] if len(sys.argv) > 1 else "F59190"
    asyncio.run(main(pid))
