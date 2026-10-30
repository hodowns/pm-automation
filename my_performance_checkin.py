#!/usr/bin/env python3
"""
My Performance / TeamSpace Check-In Automation
Reads performance_data.json (or My_Performance.md fallback) and fills the weekly check-in via CDP.

Prerequisites:
    chrome.exe --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\\chrome-dev-profile"
    Log into teamspace.cisco.com in that window, then run this script.

Usage:
    python my_performance_checkin.py
    python my_performance_checkin.py "C:\\path\\to\\performance_data.json"
"""

import asyncio
import json
import os
import re
import sys
from pathlib import Path

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

TEAMSPACE_URL   = "https://teamspace.cisco.com"
CISCO_EMAIL     = "hodowns@cisco.com"
SCRIPT_DIR      = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR     = os.path.join(SCRIPT_DIR, "teamspace_profile")
CDP_URL         = "http://localhost:9222"
# Chrome must be launched first: chrome.exe --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\chrome-dev-profile"
DEFAULT_PAYLOAD = r"C:\Users\hodowns\Downloads\performance_data.json"
DEFAULT_MD      = r"C:\Users\hodowns\Downloads\My_Performance.md"  # fallback (md with embedded json)

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
def load_fields(path: str) -> dict:
    text = Path(path).read_text(encoding="utf-8")
    if path.endswith(".json"):
        data = json.loads(text)
    else:
        match = re.search(r"```json\s*(\{.*?\})\s*```", text, re.DOTALL)
        if not match:
            raise ValueError(f"No JSON block found in {path}")
        data = json.loads(match.group(1))
    # Merge top-level keys (new_goals, etc.) into the returned fields dict
    fields = data.get("fields", data)
    if "fields" in data:
        for k, v in data.items():
            if k != "fields":
                fields.setdefault(k, v)
    return fields


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

    # Step 0: Navigate directly to the check-ins URL (most reliable for SPA router)
    current_url = page.url.lower()
    if "check-in" not in current_url:
        checkin_url = re.sub(r"/goals.*$|/summary.*$", "/check-ins", page.url)
        if checkin_url == page.url:
            checkin_url = re.sub(r"myperformance\.cisco\.com.*", "myperformance.cisco.com/check-ins", page.url)
        print(f"  Navigating directly to: {checkin_url}")
        try:
            await page.goto(checkin_url, wait_until="domcontentloaded", timeout=20000)
            await asyncio.sleep(3)
            print(f"  Landed: {page.url}")
        except Exception as e:
            print(f"  Direct nav failed: {e}")

    # Step 1: Click the Check-ins tab by text (coordinate fallback if still needed)
    tab_clicked = "check-in" in page.url.lower()
    if not tab_clicked:
        for tab_sel in [
            "text=Check-ins",
            "text=Check-In",
            "text=Check-Ins",
            "[role='tab']:has-text('Check-in')",
            "[role='tab']:has-text('Check-ins')",
            "nav >> text=Check-ins",
            "nav >> text=Check-In",
            "a:has-text('Check-ins')",
            "a:has-text('Check-In')",
            "a:has-text('Check-Ins')",
        ]:
            try:
                tab = page.locator(tab_sel).first
                if await tab.count() > 0:
                    await tab.click(force=True)
                    print(f"  ✓ Clicked Check-ins tab via: {tab_sel}")
                    tab_clicked = True
                    await asyncio.sleep(3)
                    break
            except Exception:
                pass
    if not tab_clicked:
        print("  Tab not found by text — coordinate fallback (~254, 93)...")
        await page.mouse.click(254, 93)
        await asyncio.sleep(3)

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

    # Step 4: Wait for the check-in wizard to appear.
    # We require EITHER a "Start" / "Resume" button OR a textarea that is NOT on the Goals page.
    print("  Waiting for form fields to render (up to 30s)...")
    for _ in range(15):
        await asyncio.sleep(2)
        # If we're still on Goals, keep waiting
        heading = await page.locator("h1, h2").first.inner_text(timeout=500) if await page.locator("h1, h2").count() > 0 else ""
        if "goal" in heading.lower():
            continue
        count = await page.locator("textarea").count()
        if count >= 1:
            print(f"  Form ready — found {count} textarea(s).")
            ss = os.path.join(SCRIPT_DIR, "debug_form_ready.png")
            await page.screenshot(path=ss, full_page=True)
            print(f"  Form screenshot: {ss}")
            return True
        # Also accept if Start/Resume check-in button is visible (pre-wizard landing)
        for btn_text in ["Start weekly check-in", "Resume weekly check-in", "Start Check-In", "Resume Check-In"]:
            btn = page.get_by_role("button", name=re.compile(btn_text, re.I))
            if await btn.count() > 0:
                await btn.first.click(force=True)
                await asyncio.sleep(4)
                break

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
    if len(sys.argv) > 1:
        payload_path = sys.argv[1]
    elif Path(DEFAULT_PAYLOAD).exists():
        payload_path = DEFAULT_PAYLOAD
    else:
        payload_path = DEFAULT_MD
    print(f"Loading fields from: {payload_path}")
    fields = load_fields(payload_path)
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
            if any(h in url for h in ("teamspace.cisco.com", "myperformance.cisco.com")) and \
               not any(kw in url for kw in ("login", "sso", "duo", "auth", "idp")):
                print(f"  Loaded: {page.url[:80]}")
                break
        await asyncio.sleep(3)

        # ── Navigate to Goals and add new strategic goals ─────────────────────
        new_goals = fields.get("new_goals", [])
        if new_goals:
            print(f"\n── Goals: adding {len(new_goals)} new strategic goal(s) ──")
            goals_nav_clicked = False
            for goals_sel in [
                "text=Goals",
                "[role='tab']:has-text('Goals')",
                "nav >> text=Goals",
                "a:has-text('Goals')",
            ]:
                try:
                    tab = page.locator(goals_sel).first
                    if await tab.count() > 0:
                        await tab.click(force=True)
                        print(f"  ✓ Clicked Goals tab via: {goals_sel}")
                        goals_nav_clicked = True
                        await asyncio.sleep(2)
                        break
                except Exception:
                    pass
            if not goals_nav_clicked:
                print("  ✗ Goals tab not found — skipping goal creation")
                new_goals = []

            for i, goal in enumerate(new_goals):
                g_title = goal.get("title", "").strip()
                g_desc  = goal.get("description", "").strip()
                if not g_title:
                    continue

                # Click "Create goal" button on Goals page → opens the modal
                for add_sel in [
                    page.get_by_role("button", name=re.compile(r"create goal", re.I)).first,
                    page.locator("hbr-button").filter(has_text=re.compile(r"create goal", re.I)).first,
                    page.locator("button").filter(has_text=re.compile(r"create goal", re.I)).first,
                ]:
                    try:
                        if await add_sel.count() > 0:
                            await add_sel.scroll_into_view_if_needed(timeout=3000)
                            await add_sel.click(force=True)
                            await asyncio.sleep(1.5)
                            break
                    except Exception:
                        pass

                # Select "Business Outcome" goal type (first card in the modal)
                await page.evaluate("""() => {
                    for (const el of document.querySelectorAll('button, [role="button"], [class*="card"], [class*="option"]')) {
                        if (/business outcome/i.test(el.textContent || '')) {
                            el.click(); return true;
                        }
                    }
                    return false;
                }""")
                await asyncio.sleep(0.5)

                # Fill goal text — the modal has ONE textarea (the goal statement = title)
                ok_t = await fill_shadow_textarea(page, g_title[:490], f"goal_title_{i+1}")

                # Click "Create goal" button in the modal
                created = False
                for save_sel in [
                    page.get_by_role("button", name=re.compile(r"create goal", re.I)).first,
                    page.locator("hbr-button").filter(has_text=re.compile(r"create goal", re.I)).first,
                    page.locator("button").filter(has_text=re.compile(r"create goal", re.I)).first,
                ]:
                    try:
                        if await save_sel.count() > 0:
                            await save_sel.click(force=True)
                            await asyncio.sleep(2)
                            print(f"  ✓ Created goal {i+1}: {g_title[:60]}")
                            created = True
                            break
                    except Exception:
                        pass
                if not created:
                    ok_c = await page.evaluate("""() => {
                        for (const el of document.querySelectorAll('hbr-button, button, [role="button"]')) {
                            if (/create goal/i.test(el.textContent || '')) {
                                (el.shadowRoot?.querySelector('button') || el).click(); return true;
                            }
                        }
                        return false;
                    }""")
                    if ok_c:
                        await asyncio.sleep(2)
                        print(f"  ✓ Created goal {i+1} via JS: {g_title[:60]}")

            ss_goals = os.path.join(SCRIPT_DIR, "checkin_goals.png")
            await page.screenshot(path=ss_goals, full_page=True)
            print(f"  Goals screenshot: {ss_goals}")

        # ── Navigate to the Check-In form ─────────────────────────────────────
        print("\nLooking for Check-In form entry point...")
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

        # Leader communication: select "Chat/Email"
        comm_clicked = False
        for comm_loc in [
            page.locator("hbr-button").filter(has_text=re.compile(r"^chat[\s/]*email$", re.I)).first,
            page.get_by_role("button", name=re.compile(r"chat.*email|email.*chat", re.I)).first,
            page.locator("button").filter(has_text=re.compile(r"chat[\s/]*email", re.I)).first,
        ]:
            try:
                if await comm_loc.count() > 0:
                    await comm_loc.wait_for(state="attached", timeout=3000)
                    await comm_loc.click(force=True)
                    print("  ✓ Clicked 'Chat/Email'")
                    await asyncio.sleep(1)
                    comm_clicked = True
                    break
            except Exception:
                pass
        if not comm_clicked:
            print("  ✗ Chat/Email button not found — trying JS..."  )
            ok = await page.evaluate("""() => {
                for (const el of document.querySelectorAll('button, hbr-button, [role="button"]')) {
                    const t = (el.textContent || '').trim().toLowerCase();
                    if (t.includes('chat') || t.includes('email')) {
                        el.click(); return true;
                    }
                    const inner = el.shadowRoot?.querySelector('button');
                    if (inner && (inner.textContent || '').toLowerCase().includes('chat')) {
                        inner.click(); return true;
                    }
                }
                return false;
            }""")
            print(f"  {'✓' if ok else '✗'} Chat/Email via JS")

        # Improvement textarea — sourced from transcripts (improvement_insight field)
        improvement_text = (fields.get("improvement_insight", fields.get("weekly_accomplishments", ""))[:490]).rstrip()
        ok = await fill_shadow_textarea(page, improvement_text, "improvement")
        print(f"  {'✓' if ok else '✗'} Improvement textarea ({len(improvement_text)} chars)")
        await asyncio.sleep(1)

        # Take Step 1 bottom screenshot before navigating
        await asyncio.sleep(1)
        ss_step1_bottom = os.path.join(SCRIPT_DIR, "checkin_step1_bottom.png")
        await page.screenshot(path=ss_step1_bottom, full_page=True)
        print(f"  Step 1 screenshot: {ss_step1_bottom}")

        # Resolve last week's priorities:
        # - Cards 1 & 2 (oldest, top of list): click checkmark ✓ (done)
        # - Cards 3+: click rollover ↻ (carry forward to next week)
        print("  Resolving last week's priorities...")
        resolved = 0
        rolled = 0
        try:
            await page.evaluate("() => window.scrollTo(0, 0)")
            await asyncio.sleep(1)

            # Collect all priority cards by finding pairs of icon buttons in main content area
            # Strategy: get ALL buttons with y > 120 (skip nav), width < 60 (icon-sized)
            # Group them by row (similar y-coordinate) → each row = one card's button pair
            all_btns = page.locator("button")
            btn_count = await all_btns.count()
            card_buttons = []  # list of (y, [btn_left, btn_right]) per card
            raw = []
            for bi in range(btn_count):
                btn = all_btns.nth(bi)
                try:
                    if not await btn.is_visible(timeout=200):
                        continue
                    box = await btn.bounding_box()
                    if not box or box["y"] < 120 or box["width"] >= 60:
                        continue
                    raw.append((box["y"], box["x"], btn))
                except Exception:
                    continue

            # Group into rows (buttons within 30px y of each other = same card)
            raw.sort(key=lambda r: (r[0], r[1]))
            rows = []
            for y, x, btn in raw:
                if rows and abs(y - rows[-1][0]) < 30:
                    rows[-1][1].append((x, btn))
                else:
                    rows.append([y, [(x, btn)]])

            # Sort each row's buttons left→right; rows are top→bottom = oldest→newest
            card_pairs = []
            for row_y, btns_in_row in rows:
                btns_in_row.sort(key=lambda b: b[0])
                card_pairs.append((row_y, [b for _, b in btns_in_row]))

            print(f"  Found {len(card_pairs)} priority card(s) to resolve")

            for idx, (row_y, btns) in enumerate(card_pairs):
                if len(btns) < 1:
                    continue
                if idx < 2:
                    # Oldest 2: click checkmark (first/left button)
                    btn_to_click = btns[0]
                    action = "checkmark ✓"
                else:
                    # Rest: click rollover (second/right button), fall back to first
                    btn_to_click = btns[1] if len(btns) > 1 else btns[0]
                    action = "rollover ↻"
                try:
                    await btn_to_click.scroll_into_view_if_needed(timeout=2000)
                    await btn_to_click.click(force=True)
                    await asyncio.sleep(1.5)
                    if idx < 2:
                        resolved += 1
                    else:
                        rolled += 1
                    print(f"    card {idx+1}: {action}")
                except Exception as ce:
                    print(f"    card {idx+1} click failed: {ce}")

            print(f"  ✓ Resolved {resolved}, rolled over {rolled} prior priorit{'y' if (resolved+rolled)==1 else 'ies'}")
            await asyncio.sleep(2)
        except Exception as e:
            print(f"  ✗ Could not resolve priorities: {e}")
            await asyncio.sleep(1)

        # Click "Save & continue" to advance to Step 2
        step1_advanced = await click_save_continue(page)
        print(f"  {'✓' if step1_advanced else '✗'} Save & continue (Step 1 → Step 2)")

        ss_step2 = os.path.join(SCRIPT_DIR, "checkin_step2.png")
        await page.screenshot(path=ss_step2, full_page=True)
        print(f"  After Step 1 nav screenshot: {ss_step2}")

        print("\n── Step 2: Plan ──")
        await asyncio.sleep(3)

        # Add new priorities via "Add Priority" button (4–6 items from new_priorities field)
        new_priorities = fields.get("new_priorities", [])
        if not new_priorities:
            # Fallback: split top_priorities_next_week into individual items
            raw = fields.get("top_priorities_next_week", "")
            new_priorities = [
                {"text": line.lstrip("0123456789.-• ").strip(), "goal": ""}
                for line in raw.splitlines() if line.strip()
            ][:6]

        # Diagnostic: dump all button texts on Step 2 page to diagnose "Add Priority" failures
        try:
            step2_btn_texts = await page.evaluate("""() => {
                return [...document.querySelectorAll('button, hbr-button, [role="button"]')]
                    .map(el => (el.textContent || '').trim())
                    .filter(t => t.length > 0 && t.length < 80);
            }""")
            print(f"  Step 2 buttons visible: {step2_btn_texts[:20]}")
        except Exception:
            pass

        print(f"  Adding {len(new_priorities)} new priorit{'y' if len(new_priorities) == 1 else 'ies'}...")
        for i, priority in enumerate(new_priorities):
            p_text = priority.get("text", "").strip()
            p_goal = priority.get("goal", "").strip()
            if not p_text:
                continue

            # Click "Add Priority" button
            add_btn_clicked = False
            for add_sel in [
                page.locator("hbr-button").filter(has_text=re.compile(r"add priority", re.I)).first,
                page.get_by_role("button", name=re.compile(r"add priority", re.I)).first,
                page.locator("button").filter(has_text=re.compile(r"add priority", re.I)).first,
            ]:
                try:
                    if await add_sel.count() > 0:
                        await add_sel.scroll_into_view_if_needed(timeout=3000)
                        await add_sel.click(force=True)
                        await asyncio.sleep(1.5)
                        add_btn_clicked = True
                        break
                except Exception:
                    pass
            if not add_btn_clicked:
                ok = await page.evaluate("""() => {
                    for (const el of document.querySelectorAll('hbr-button, button, [role="button"]')) {
                        if (/add priority/i.test(el.textContent)) {
                            const inner = el.shadowRoot?.querySelector('button') || el;
                            inner.click(); return true;
                        }
                    }
                    return false;
                }""")
                await asyncio.sleep(1.5)
                if not ok:
                    print(f"  ✗ 'Add Priority' button not found for item {i+1}")
                    continue

            # "Add priority" opens a MODAL: Goal combobox → Priority textarea → "Save priority"
            # Use Playwright locators scoped to the modal for proper async handling.
            await asyncio.sleep(1.5)  # give modal time to fully render

            # --- Goal: open combobox via Playwright, wait for options, click first ---
            try:
                cb_loc = page.locator("[role='combobox']").first
                if await cb_loc.count() > 0:
                    await cb_loc.scroll_into_view_if_needed(timeout=3000)
                    await cb_loc.click(force=True)
                    # Wait for dropdown options to appear in the DOM
                    try:
                        await page.locator("[role='option']").first.wait_for(state="visible", timeout=4000)
                        first_opt = page.locator("[role='option']").first
                        await first_opt.click(force=True)
                        print(f"    ✓ Goal selected")
                        await asyncio.sleep(0.5)
                    except PlaywrightTimeout:
                        # Options didn't appear — try JS approach
                        ok_g = await page.evaluate("""() => {
                            const lb = document.querySelector('[role="listbox"]');
                            if (lb) {
                                const opts = [...lb.querySelectorAll('[role="option"], li')];
                                if (opts.length > 0) { opts[0].click(); return true; }
                            }
                            for (const el of document.querySelectorAll('[role="option"]')) {
                                el.click(); return true;
                            }
                            return false;
                        }""")
                        print(f"    {'✓' if ok_g else '✗'} Goal via JS listbox")
                        await asyncio.sleep(0.3)
            except Exception as e:
                print(f"    ✗ Goal combobox error: {e}")

            # --- Priority text: focus the modal textarea then keyboard.type() ---
            # The native-setter approach doesn't trigger framework input handlers.
            # Focus via JS, then use page.keyboard.type() for real key events.
            focused = await page.evaluate("""() => {
                for (const modal of document.querySelectorAll('[role="dialog"], dialog, [aria-modal="true"]')) {
                    const h = modal.querySelector('hbr-textarea');
                    if (h) {
                        const ta = h.shadowRoot?.querySelector('textarea');
                        if (ta) { ta.focus(); ta.click(); return true; }
                    }
                }
                // Fallback: any hbr-textarea whose shadow textarea is empty
                for (const h of document.querySelectorAll('hbr-textarea')) {
                    const ta = h.shadowRoot?.querySelector('textarea');
                    if (ta && !ta.value) { ta.focus(); ta.click(); return true; }
                }
                return false;
            }""")
            await asyncio.sleep(0.3)
            if focused:
                await page.keyboard.type(p_text[:490], delay=5)
                filled = True
            else:
                filled = False
            print(f"  {'✓' if filled else '✗'} Priority {i+1}: {p_text[:60]}...")
            await asyncio.sleep(0.5)

            # --- Save priority button ---
            saved = False
            try:
                modal_loc = page.get_by_role("dialog")
                if await modal_loc.count() > 0:
                    save_btn = modal_loc.get_by_role("button", name=re.compile(r"save priority", re.I))
                    if await save_btn.count() > 0:
                        await save_btn.click(force=True)
                        await asyncio.sleep(1.5)
                        saved = True
                        print(f"    ✓ Saved priority {i+1}")
            except Exception:
                pass
            if not saved:
                for save_loc in [
                    page.get_by_role("button", name=re.compile(r"save priority", re.I)).first,
                    page.locator("hbr-button").filter(has_text=re.compile(r"save priority", re.I)).first,
                ]:
                    try:
                        if await save_loc.count() > 0:
                            await save_loc.click(force=True)
                            await asyncio.sleep(1.5)
                            saved = True
                            print(f"    ✓ Saved priority {i+1} (Playwright fallback)")
                            break
                    except Exception:
                        pass
            if not saved:
                ok_s = await page.evaluate("""() => {
                    for (const el of document.querySelectorAll('hbr-button, button, [role="button"]')) {
                        if (/save priority/i.test(el.textContent||'')) {
                            (el.shadowRoot?.querySelector('button')||el).click(); return true;
                        }
                    }
                    return false;
                }""")
                if ok_s:
                    await asyncio.sleep(1.5)
                    print(f"    ✓ Saved priority {i+1} via JS")
                else:
                    print(f"    ✗ 'Save priority' not found for priority {i+1}")

            await asyncio.sleep(0.5)

            # Debug screenshot after first priority only
            if i == 0:
                ss_p1 = os.path.join(SCRIPT_DIR, "debug_after_priority1.png")
                await page.screenshot(path=ss_p1, full_page=True)
                print(f"  Debug screenshot after priority 1: {ss_p1}")

        # Fill "What do you need from your leader?" textarea with blockers field
        leader_text = fields.get("blockers_and_support_needed", "").strip()[:490]
        if leader_text:
            ok_lt = await fill_shadow_textarea(page, leader_text, "leader_needs")
            print(f"  {'✓' if ok_lt else '✗'} Leader needs textarea ({len(leader_text)} chars)")
            await asyncio.sleep(0.5)

        # Click "Save & continue" to advance to Step 3
        step2_advanced = await click_save_continue(page)
        print(f"  {'✓' if step2_advanced else '✗'} Save & continue (Step 2 → Step 3)")

        ss_step3 = os.path.join(SCRIPT_DIR, "checkin_step3.png")
        await page.screenshot(path=ss_step3, full_page=True)
        print(f"  Step 3 screenshot: {ss_step3}")

        ss_filled = os.path.join(SCRIPT_DIR, "checkin_02_filled.png")
        await page.screenshot(path=ss_filled, full_page=True)
        print(f"Screenshot (filled):  {ss_filled}")

        # ── Step 3: Review & Submit ───────────────────────────────────────────
        print("\nSubmitting check-in...")
        submitted = False
        for sub_loc in [
            page.locator("hbr-button").filter(has_text=re.compile(r"^submit$", re.I)).first,
            page.get_by_role("button", name=re.compile(r"^submit$", re.I)).first,
            page.locator("button").filter(has_text=re.compile(r"^submit$", re.I)).first,
        ]:
            try:
                if await sub_loc.count() > 0:
                    await sub_loc.scroll_into_view_if_needed(timeout=3000)
                    await sub_loc.click(force=True)
                    await asyncio.sleep(4)
                    submitted = True
                    print("  ✓ Submitted check-in")
                    break
            except Exception:
                pass
        if not submitted:
            ok = await page.evaluate("""() => {
                for (const el of document.querySelectorAll('hbr-button, button, [role="button"]')) {
                    const t = (el.textContent || '').trim().toLowerCase();
                    if (t === 'submit') {
                        const inner = el.shadowRoot?.querySelector('button') || el;
                        inner.click(); return true;
                    }
                }
                return false;
            }""")
            if ok:
                await asyncio.sleep(4)
                submitted = True
                print("  ✓ Submitted via JS")
            else:
                print("  ✗ Submit button not found — please click manually in the browser.")

        # Final screenshot
        ss_final = os.path.join(SCRIPT_DIR, "checkin_03_saved.png")
        await page.screenshot(path=ss_final, full_page=True)
        print(f"Screenshot (final):   {ss_final}")
        print("\n✓ Done. Keeping browser open for 60 seconds for review...")
        await asyncio.sleep(60)
        await context.close()


if __name__ == "__main__":
    asyncio.run(main())
