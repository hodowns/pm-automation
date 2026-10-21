#!/usr/bin/env python3
"""
My Performance / Team Space Check-In Automation
Reads the JSON payload from a Markdown file and fills the weekly check-in form.

Usage:
    python my_performance_checkin.py
    python my_performance_checkin.py "C:\\path\\to\\My_Performance.md"
"""

import asyncio
import json
import os
import re
import sys
from pathlib import Path

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

TEAMSPACE_URL  = "https://myperformance.cisco.com"
CISCO_EMAIL    = "hodowns@cisco.com"
SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR    = os.path.join(SCRIPT_DIR, "teamspace_profile")
DEFAULT_MD     = r"C:\Users\hodowns\Downloads\My_Performance.md"
CDP_URL        = "http://localhost:9222"  # attach to existing Chrome with SSO session

# ── Label text patterns for each form field ───────────────────────────────────
FIELD_MAP = {
    "weekly_accomplishments": [
        "What were your key accomplishments and wins this week",
        "accomplishments and wins",
        "key accomplishments",
        "accomplishments",
        "wins this week",
    ],
    "top_priorities_next_week": [
        "What are your top focus areas and priorities for next week",
        "top focus areas and priorities",
        "priorities for next week",
        "focus areas",
        "next week",
    ],
    "customer_impact_value": [
        "How did your work drive customer value and business impact",
        "customer value and business impact",
        "drive customer value",
        "business impact",
    ],
    "blockers_and_support_needed": [
        "What blockers, risks, or support do you need from leadership",
        "blockers, risks",
        "support do you need from leadership",
        "blockers",
        "risks",
    ],
    "internal_leadership_contributions": [
        "Innovation, Internal Champions",
        "PMO Contributions",
        "Internal Champions",
        "innovation",
        "contributions",
    ],
}

# ── Parse the Markdown file ────────────────────────────────────────────────────
def load_fields(md_path: str) -> dict:
    text = Path(md_path).read_text(encoding="utf-8")
    match = re.search(r"```json\s*(\{.*?\})\s*```", text, re.DOTALL)
    if not match:
        raise ValueError(f"No JSON block found in {md_path}")
    data = json.loads(match.group(1))
    return data.get("fields", data)


# ── Try to fill one field by searching for its label ─────────────────────────
async def fill_field(page, field_key: str, text: str) -> bool:
    for label_text in FIELD_MAP[field_key]:

        # Strategy 1: get_by_label (aria-label / label[for])
        try:
            el = page.get_by_label(re.compile(re.escape(label_text), re.IGNORECASE))
            cnt = await el.count()
            if cnt > 0 and await el.first.is_visible(timeout=800):
                await el.first.click()
                await el.first.fill(text)
                print(f"  ✓ {field_key} via aria-label")
                return True
        except Exception:
            pass

        # Strategy 2: find the label element and walk up to find input
        label_selectors = [
            f"label:has-text('{label_text}')",
            f"h3:has-text('{label_text}')",
            f"h4:has-text('{label_text}')",
            f"p:has-text('{label_text}')",
            f"span:has-text('{label_text}')",
            f"div[class*='label']:has-text('{label_text}')",
            f"[class*='question']:has-text('{label_text}')",
        ]
        for lsel in label_selectors:
            try:
                label_el = page.locator(lsel).first
                if not await label_el.is_visible(timeout=600):
                    continue

                # If label has a [for], jump straight to the target
                label_for = await label_el.get_attribute("for")
                if label_for:
                    tgt = page.locator(f"#{label_for}").first
                    if await tgt.is_visible(timeout=500):
                        await tgt.click()
                        try:
                            await tgt.fill(text)
                        except Exception:
                            await page.keyboard.press("Control+a")
                            await page.keyboard.type(text)
                        print(f"  ✓ {field_key} via label[for='{label_for}']")
                        return True

                # Walk up the DOM looking for an input within the same section
                container = label_el
                for depth in range(6):
                    container = container.locator("xpath=..")

                    for inp_sel in [
                        "textarea",
                        "[contenteditable='true']",
                        "input[type='text']",
                    ]:
                        try:
                            inp = container.locator(inp_sel).first
                            if await inp.count() > 0 and await inp.is_visible(timeout=300):
                                await inp.click()
                                await asyncio.sleep(0.15)
                                try:
                                    await inp.fill(text)
                                except Exception:
                                    await page.keyboard.press("Control+a")
                                    await asyncio.sleep(0.1)
                                    await page.keyboard.type(text, delay=3)
                                print(f"  ✓ {field_key} via {inp_sel} (depth {depth+1})")
                                return True
                        except Exception:
                            pass
            except Exception:
                continue

    print(f"  ✗ Could not auto-fill '{field_key}'")
    return False


# ── Shadow DOM helpers ────────────────────────────────────────────────────────
async def fill_shadow_textarea(page, text: str, label: str = "") -> bool:
    """Fill a textarea that may live inside an hbr-textarea shadow root."""
    # Strategy 1: JS native setter — bypasses visibility checks and works for Web Components
    ok = await page.evaluate("""(text) => {
        for (const h of document.querySelectorAll('hbr-textarea')) {
            const ta = h.shadowRoot?.querySelector('textarea');
            if (ta) {
                const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set;
                setter.call(ta, text);
                ta.dispatchEvent(new Event('input', {bubbles: true, composed: true}));
                ta.dispatchEvent(new Event('change', {bubbles: true, composed: true}));
                return true;
            }
        }
        const ta = document.querySelector('textarea');
        if (ta) {
            ta.value = text;
            ta.dispatchEvent(new Event('input', {bubbles: true}));
            ta.dispatchEvent(new Event('change', {bubbles: true}));
            return true;
        }
        return false;
    }""", text)
    if ok:
        return True

    # Strategy 2: Playwright locator with force click (bypasses visibility check)
    for loc in [
        page.locator("hbr-textarea").locator("textarea").first,
        page.locator("textarea").first,
    ]:
        try:
            if await loc.count() > 0:
                await loc.wait_for(state="attached", timeout=3000)
                await loc.click(force=True)
                await asyncio.sleep(0.2)
                await loc.press_sequentially(text, delay=5)
                return True
        except Exception:
            pass

    return False


async def click_save_continue(page) -> bool:
    """Click 'Save & continue' through shadow DOM."""
    # Strategy 1: hbr-button light-DOM slot text filter + force click
    for loc in [
        page.locator("hbr-button").filter(has_text=re.compile(r"save.*continue", re.I)).first,
        page.get_by_role("button", name=re.compile(r"save.*continue", re.I)).first,
    ]:
        try:
            if await loc.count() > 0:
                await loc.scroll_into_view_if_needed(timeout=3000)
                await loc.click(force=True)
                await asyncio.sleep(4)
                return True
        except Exception:
            pass

    # Strategy 2: JS — find hbr-button by text content, click inner shadow button
    ok = await page.evaluate("""() => {
        for (const hbr of document.querySelectorAll('hbr-button')) {
            if (/save.*continue/i.test(hbr.textContent)) {
                const inner = hbr.shadowRoot?.querySelector('button') || hbr;
                inner.click();
                return true;
            }
        }
        function findDeep(root) {
            for (const el of root.querySelectorAll('button, [role="button"]')) {
                if (/save.*continue/i.test((el.textContent || '').trim()) && el.offsetParent) {
                    el.click();
                    return true;
                }
            }
            for (const el of root.querySelectorAll('*')) {
                if (el.shadowRoot && findDeep(el.shadowRoot)) return true;
            }
            return false;
        }
        return findDeep(document);
    }""")
    if ok:
        await asyncio.sleep(4)
    return bool(ok)


# ── Navigate from the landing page to the check-in form ──────────────────────
async def open_checkin_form(page) -> bool:
    await asyncio.sleep(2)

    # Step 1: Ensure we're on the Check-ins tab (coordinate click at known position)
    print("  Clicking Check-ins tab (~254, 93)...")
    await page.mouse.click(254, 93)
    await asyncio.sleep(2)

    # Step 2: Click "Start / Resume Weekly check-in" button
    print("  Waiting for check-in launch button (up to 10s)...")
    clicked = False

    # Strategy A: Playwright get_by_role with regex — waits for element to appear
    for label_pattern in [
        re.compile(r"resume weekly check.in", re.I),
        re.compile(r"start weekly check.in", re.I),
        re.compile(r"resume check.in", re.I),
        re.compile(r"start check.in", re.I),
    ]:
        try:
            btn = page.get_by_role("button", name=label_pattern)
            await btn.wait_for(state="visible", timeout=5000)
            print(f"  Clicking button matching: {label_pattern.pattern}")
            await btn.click()
            await asyncio.sleep(5)
            clicked = True
            break
        except Exception:
            pass

    # Strategy B: coordinate click — button is at ~(692, 726) in 1440x900 viewport
    if not clicked:
        print("  Trying coordinate click at (692, 726)...")
        await page.mouse.click(692, 726)
        await asyncio.sleep(5)
        clicked = True

    if clicked:
        ss = os.path.join(SCRIPT_DIR, "debug_after_start_click.png")
        await page.screenshot(path=ss, full_page=True)
        print(f"  Screenshot after launch click: {ss}")

    tab_clicked = True  # We attempted — proceed to wait for form

    # Step 3: dismiss any "Check-in Guide" or welcome modal
    for dismiss_sel in [
        "button:has-text('Got it')",
        "button:has-text('Got It')",
        "button:has-text('Dismiss')",
        "[aria-label='Close']",
        "button[class*='close']",
        ".modal button",
        "[role='dialog'] button:last-child",
    ]:
        try:
            dismiss = page.locator(dismiss_sel).first
            if await dismiss.is_visible(timeout=2000):
                print(f"  Dismissing modal: {dismiss_sel}")
                await dismiss.click()
                await asyncio.sleep(2)
                break
        except Exception:
            pass

    ss_after_modal = os.path.join(SCRIPT_DIR, "debug_after_modal.png")
    await page.screenshot(path=ss_after_modal)
    print(f"  Screenshot after modal: {ss_after_modal}")

    # Dump all visible button/link texts to help diagnose
    try:
        btns = await page.locator("button, a").all()
        btn_texts = []
        for b in btns[:30]:
            try:
                t = (await b.inner_text(timeout=300)).strip()
                if t:
                    btn_texts.append(repr(t))
            except Exception:
                pass
        print(f"  Visible buttons/links: {', '.join(btn_texts[:20])}")
    except Exception as e:
        print(f"  Could not list buttons: {e}")

    # Step 4: Wait for actual form textareas to appear (need 3+ to confirm it's the real form)
    print("  Waiting for form fields to render (up to 30s)...")
    for _ in range(15):
        await asyncio.sleep(2)
        count = await page.locator("textarea").count()
        if count >= 3:
            print(f"  Form ready — found {count} textarea(s).")
            ss = os.path.join(SCRIPT_DIR, "debug_form_ready.png")
            await page.screenshot(path=ss, full_page=True)
            print(f"  Form screenshot: {ss}")
            return True
        elif count > 0:
            print(f"  Found {count} textarea(s) — waiting for more...")

    # Fallback — dump page state for debugging
    ss = os.path.join(SCRIPT_DIR, "teamspace_checkins_tab.png")
    await page.screenshot(path=ss, full_page=True)
    print(f"  Form not found after 30s. Screenshot: {ss}")
    try:
        all_texts = await page.evaluate("""() => {
            return [...document.querySelectorAll('button, a, [role=button], [tabindex]')]
                .map(e => e.textContent.trim())
                .filter(t => t.length > 0 && t.length < 60);
        }""")
        print(f"  Clickable texts: {all_texts[:30]}")
    except Exception:
        pass
    return False


# ── Click Save as Draft ────────────────────────────────────────────────────────
async def save_draft(page) -> bool:
    for sel in [
        "button:has-text('Save as Draft')",
        "button:has-text('Save Draft')",
        "button:has-text('Save as draft')",
        "button:has-text('Submit')",
        "[data-testid*='save-draft']",
        "[data-testid*='saveDraft']",
        "a:has-text('Save as Draft')",
    ]:
        try:
            btn = page.locator(sel).first
            if await btn.is_visible(timeout=1500):
                print(f"  Clicking: {sel}")
                await btn.click()
                await asyncio.sleep(3)
                return True
        except Exception:
            pass
    return False


# ── Main ──────────────────────────────────────────────────────────────────────
async def main():
    md_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_MD
    print(f"Loading fields from: {md_path}")
    fields = load_fields(md_path)
    print(f"  {len(fields)} fields loaded: {list(fields.keys())}\n")

    os.makedirs(PROFILE_DIR, exist_ok=True)

    async with async_playwright() as p:
        # ── Try CDP first; fall back to persistent profile ────────────────────
        context = None
        try:
            browser = await p.chromium.connect_over_cdp(CDP_URL, timeout=4000)
            print(f"  CDP connected.")
            context = browser.contexts[0] if browser.contexts else await browser.new_context()
            page = await context.new_page()
        except Exception:
            print(f"  CDP not available — launching persistent profile instead...")
            context = await p.chromium.launch_persistent_context(
                user_data_dir=PROFILE_DIR,
                channel="chrome",
                headless=False,
                slow_mo=250,
                viewport={"width": 1440, "height": 900},
            )
            page = await context.new_page()

        await page.set_viewport_size({"width": 1440, "height": 900})
        print(f"Opening {TEAMSPACE_URL} ...")
        await page.goto(TEAMSPACE_URL, wait_until="domcontentloaded", timeout=45000)

        # ── Wait for SSO / Duo if needed ──────────────────────────────────────
        try:
            email_inp = page.locator("input[type='email'], input[name='email'], input[placeholder*='email' i], input#email").first
            if await email_inp.is_visible(timeout=4000):
                await email_inp.fill(CISCO_EMAIL)
                await email_inp.press("Enter")
                print(f"  Email entered — complete Duo in the browser window if prompted...")
        except Exception:
            pass

        print("Waiting for TeamSpace to load (up to 3 min if Duo is needed)...")
        deadline = asyncio.get_event_loop().time() + 180
        while asyncio.get_event_loop().time() < deadline:
            await asyncio.sleep(2)
            url = page.url.lower()
            if "teamspace.cisco.com" in url and not any(kw in url for kw in ("login", "sso", "duo", "auth", "idp")):
                print(f"  Loaded: {page.url[:80]}")
                break
        await asyncio.sleep(3)

        # ── Navigate to the Check-In form ─────────────────────────────────────
        print("Looking for Check-In form entry point...")
        found = await open_checkin_form(page)
        if not found:
            print("  Form not auto-detected. Please open the check-in form manually in the browser window.")
            print("  Waiting up to 60s for form fields to appear...")
            for _ in range(30):
                await asyncio.sleep(2)
                count = await page.locator("textarea, [contenteditable='true']").count()
                if count > 0:
                    print(f"  Form detected manually — {count} fields visible.")
                    found = True
                    break

        # Screenshot of the form as it loaded
        ss_before = os.path.join(SCRIPT_DIR, "checkin_01_before.png")
        await page.screenshot(path=ss_before, full_page=True)
        print(f"Screenshot (before): {ss_before}")

        # ── 3-Step Wizard Fill ────────────────────────────────────────────────
        print("\n── Step 1: Reflect ──")

        # Manager connection: click "Meeting"
        mtg_clicked = False
        for mtg_loc in [
            page.locator("hbr-button").filter(has_text=re.compile(r"^meeting$", re.I)).first,
            page.get_by_role("button", name=re.compile(r"^meeting$", re.I)).first,
        ]:
            try:
                if await mtg_loc.count() > 0:
                    await mtg_loc.wait_for(state="attached", timeout=3000)
                    await mtg_loc.click(force=True)
                    print("  ✓ Clicked 'Meeting'")
                    await asyncio.sleep(1)
                    mtg_clicked = True
                    break
            except Exception:
                pass
        if not mtg_clicked:
            print("  ✗ Meeting button not found")

        # Improvement textarea (500 char limit) — use shortened accomplishments
        improvement_text = (fields.get("weekly_accomplishments", "")[:490]).rstrip()
        ok = await fill_shadow_textarea(page, improvement_text, "improvement")
        print(f"  {'✓' if ok else '✗'} Improvement textarea ({len(improvement_text)} chars)")
        await asyncio.sleep(1)

        # Take Step 1 bottom screenshot before navigating
        await asyncio.sleep(1)
        ss_step1_bottom = os.path.join(SCRIPT_DIR, "checkin_step1_bottom.png")
        await page.screenshot(path=ss_step1_bottom, full_page=True)
        print(f"  Step 1 screenshot: {ss_step1_bottom}")

        # Click "Save & continue" to advance to Step 2
        step1_advanced = await click_save_continue(page)
        print(f"  {'✓' if step1_advanced else '✗'} Save & continue (Step 1 → Step 2)")

        ss_step2 = os.path.join(SCRIPT_DIR, "checkin_step2.png")
        await page.screenshot(path=ss_step2, full_page=True)
        print(f"  After Step 1 nav screenshot: {ss_step2}")

        print("\n── Step 2: Plan ──")

        # Wait for Step 2 heading to appear
        try:
            await page.locator("text=Step 2, text=Plan").first.wait_for(state="visible", timeout=8000)
            print("  Step 2 loaded")
        except Exception:
            print("  (Could not confirm Step 2 loaded — proceeding anyway)")

        # Fill priorities textarea — wait for Step 2 form to render
        priorities_text = fields.get("top_priorities_next_week", "").strip()
        await asyncio.sleep(2)
        ok = await fill_shadow_textarea(page, priorities_text, "priorities")
        print(f"  {'✓' if ok else '✗'} Priorities textarea ({len(priorities_text)} chars)")
        await asyncio.sleep(1)

        # Click "Save & continue" to advance to Step 3
        step2_advanced = await click_save_continue(page)
        print(f"  {'✓' if step2_advanced else '✗'} Save & continue (Step 2 → Step 3)")

        ss_step3 = os.path.join(SCRIPT_DIR, "checkin_step3.png")
        await page.screenshot(path=ss_step3, full_page=True)
        print(f"  Step 3 screenshot: {ss_step3}")

        # Screenshot of filled form
        ss_filled = os.path.join(SCRIPT_DIR, "checkin_02_filled.png")
        await page.screenshot(path=ss_filled, full_page=True)
        print(f"Screenshot (filled):  {ss_filled}")

        # ── Save as Draft ─────────────────────────────────────────────────────
        print("\nSaving as Draft...")
        ok = await save_draft(page)
        if not ok:
            print("  'Save as Draft' button not found automatically.")
            print("  Please click it in the browser.")
            print("  The script will wait up to 3 min for confirmation (URL change or success banner)...")
            url_before = page.url
            deadline_save = asyncio.get_event_loop().time() + 180
            while asyncio.get_event_loop().time() < deadline_save:
                await asyncio.sleep(2)
                # Check for a success indicator
                url_now = page.url
                success_text = await page.locator(
                    "text=saved, text=draft, [class*='success'], [class*='toast'], [role='alert']"
                ).count()
                if url_now != url_before or success_text > 0:
                    print("  Detected save action.")
                    break

        # Final screenshot
        ss_final = os.path.join(SCRIPT_DIR, "checkin_03_saved.png")
        await page.screenshot(path=ss_final, full_page=True)
        print(f"Screenshot (final):   {ss_final}")
        print("\n✓ Done. Keeping browser open for 60 seconds for review...")
        await asyncio.sleep(60)
        await context.close()


if __name__ == "__main__":
    asyncio.run(main())
