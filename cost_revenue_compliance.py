#!/usr/bin/env python3
"""
SDC Cost & Revenue Compliance — Discovery / Automation
Navigates to each active project in SDC and performs cost/revenue compliance steps.

DISCOVERY MODE (default — run first to map the UI):
    python -u -X utf8 cost_revenue_compliance.py F51119 F48852 F40696
    Generates: Desktop\SDC Compliance Discovery\discovery_report.html

AUTOMATION MODE (run after confirming selectors from discovery report):
    python -u -X utf8 cost_revenue_compliance.py --run F51119 F48852 F40696

Claude fetches active PIDs via sfdc_get_my_projects() and passes them as CLI args.
"""

import asyncio
import base64
import os
import sys
from datetime import datetime
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

SDC_URL = "https://csone.lightning.force.com/"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Reuse the persistent Chrome profile from close_project.py (SSO session already active)
PROFILE_DIR = os.path.normpath(
    os.path.join(
        SCRIPT_DIR,
        "Claude How To",
        "Close a project workflow Gen 1",
        "sdc_profile",
    )
)

OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "SDC Compliance Discovery")
SCREENSHOT_DIR = os.path.join(OUTPUT_DIR, "screenshots")

COMPLIANCE_KEYWORDS = [
    "forecast", "compliance", "cost", "revenue", "financial", "budget",
    "submit", "approve", "commit", "publish", "accept",
]

# Tabs to explore on each project record
TABS_TO_EXPLORE = [
    "Details", "Status", "Financials", "Financial Summary",
    "Forecast", "Revenue", "Cost", "Compliance", "Budget",
]


# ─────────────────────────────────────────────
#  BROWSER / LOGIN SETUP
# ─────────────────────────────────────────────

async def launch_and_login(playwright):
    os.makedirs(PROFILE_DIR, exist_ok=True)
    context = await playwright.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        channel="chrome",
        headless=False,
        slow_mo=300,
    )
    page = await context.new_page()

    print("Opening SDC...")
    await page.goto(SDC_URL)

    print("Waiting for login (up to 3 min)...")
    deadline = asyncio.get_event_loop().time() + 180
    logged_in = False
    while asyncio.get_event_loop().time() < deadline:
        await asyncio.sleep(2)
        url = page.url
        if ("force.com" in url or "salesforce.com" in url) and "login" not in url:
            logged_in = True
            print("  Logged in.")
            break

    if not logged_in:
        print("ERROR: Login timed out.")
        await context.close()
        return None, None

    if "lightning.force.com" not in page.url:
        try:
            await page.goto(SDC_URL, wait_until="commit", timeout=30000)
        except Exception:
            pass

    # Wait for PM Console search bar
    print("  Waiting for PM Console...")
    for _ in range(40):
        await asyncio.sleep(3)
        try:
            search_ok = await page.locator(
                "input[type='search'], .slds-global-header input"
            ).first.is_visible(timeout=1000)
            if search_ok:
                print("  PM Console ready.")
                break
        except Exception:
            pass
    await asyncio.sleep(2)

    return context, page


# ─────────────────────────────────────────────
#  SEARCH & OPEN PROJECT  (same logic as close_project.py)
# ─────────────────────────────────────────────

async def find_and_open_project(page, pid: str) -> bool:
    print(f"  Searching for {pid}...")

    clicked = False
    try:
        search = page.locator(".slds-global-header input[type='search']")
        await search.wait_for(state="visible", timeout=8000)
        await search.click()
        clicked = True
    except Exception:
        pass

    if not clicked:
        vp = page.viewport_size or {"width": 1280, "height": 720}
        x = int(vp["width"] * 0.48)
        for attempt in range(5):
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
    target_page = None
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
                        if text == f"{pid} - Search" or text.startswith(f"{pid} - "):
                            continue
                        if await link.is_visible(timeout=500):
                            target_page = p
                            project_link = link
                            print(f"  Found: '{text}'")
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
        print(f"  ERROR: Could not find project link for {pid}.")
        return False

    await project_link.click()
    await asyncio.sleep(5)

    if "pse__Proj__c" in page.url:
        print(f"  Project opened: {page.url[:80]}")
        return True

    if target_page != page:
        for p in context.pages:
            if "pse__Proj__c" in p.url:
                await page.goto(p.url, wait_until="domcontentloaded", timeout=30000)
                await asyncio.sleep(3)
                print(f"  Project opened (from other tab): {page.url[:80]}")
                return True

    print("  Waiting for project page to load in workspace tab...")
    for _ in range(6):
        await asyncio.sleep(3)
        try:
            pid_visible = await page.locator(f"text={pid}").first.is_visible(timeout=1000)
            if pid_visible:
                print(f"  Project page confirmed.")
                return True
        except Exception:
            pass

    # Work Planner fallback
    try:
        page_text = await page.evaluate("() => document.body.innerText")
        if "Work Planner" in page_text:
            print("  Detected Work Planner — navigating to project record sub-tab...")
            await page.evaluate(
                """(pid) => {
                    const sels = ['li.slds-sub-tabs__item a', '[class*="subTab"] a',
                                  '[class*="sub-tab"] a', '[role="tablist"] li a'];
                    for (const sel of sels) {
                        for (const a of document.querySelectorAll(sel)) {
                            const t = (a.textContent || '').trim();
                            if (t && !t.includes('Work Planner') &&
                                (t.includes(pid) || t.includes('COH') || t.includes('TX-T-'))) {
                                a.click(); return true;
                            }
                        }
                    }
                    return false;
                }""",
                pid,
            )
            await asyncio.sleep(3)
    except Exception:
        pass

    print("  Warning: Project page may not be fully loaded. Proceeding anyway.")
    return True


# ─────────────────────────────────────────────
#  DISCOVERY HELPERS
# ─────────────────────────────────────────────

async def take_screenshot(page, pid: str, label: str) -> str:
    filename = f"{pid}_{label}_{datetime.now().strftime('%H%M%S')}.png"
    path = os.path.join(SCREENSHOT_DIR, filename)
    await page.screenshot(path=path, full_page=False)
    return path


async def enumerate_page_elements(page) -> dict:
    """Return all tabs, visible buttons, and compliance-keyword elements."""
    return await page.evaluate(
        """(keywords) => {
            const lower = s => (s || '').toLowerCase();
            const getText = el =>
                (el.innerText || el.textContent || el.value ||
                 el.getAttribute('aria-label') || el.getAttribute('title') || '').trim();

            const tabs = [];
            document.querySelectorAll(
                '[role="tab"], .slds-tabs__item a, .slds-sub-tabs__item a'
            ).forEach(el => {
                const t = getText(el);
                if (t && !tabs.includes(t)) tabs.push(t);
            });

            const buttons = [];
            document.querySelectorAll(
                'button, input[type="button"], input[type="submit"], a[role="button"]'
            ).forEach(el => {
                if (el.offsetParent !== null) {
                    const t = getText(el);
                    if (t && t.length < 80 && !buttons.includes(t)) buttons.push(t);
                }
            });

            const compliance = [];
            document.querySelectorAll(
                'button, a, h1, h2, h3, h4, span, label, td, th, li, div[class*="title"]'
            ).forEach(el => {
                if (el.offsetParent !== null) {
                    const t = getText(el);
                    if (t && t.length < 100) {
                        for (const kw of keywords) {
                            if (lower(t).includes(kw)) {
                                const entry = '[' + el.tagName.toLowerCase() + '] ' + t;
                                if (!compliance.includes(entry)) compliance.push(entry);
                                break;
                            }
                        }
                    }
                }
            });

            return {
                tabs,
                buttons: buttons.slice(0, 60),
                compliance: compliance.slice(0, 120),
            };
        }""",
        COMPLIANCE_KEYWORDS,
    )


# ─────────────────────────────────────────────
#  DISCOVERY MODE — per PID
# ─────────────────────────────────────────────

async def discover_project(page, pid: str) -> dict:
    discovery = {
        "pid": pid,
        "url": "",
        "tabs_found": [],
        "tab_data": {},
        "screenshots": [],
        "error": None,
    }

    print(f"\n{'='*55}")
    print(f"  DISCOVERING: {pid}")
    print(f"{'='*55}")

    if not await find_and_open_project(page, pid):
        discovery["error"] = "Project not found in SDC search"
        return discovery

    discovery["url"] = page.url

    ss = await take_screenshot(page, pid, "01_main")
    discovery["screenshots"].append(ss)

    initial = await enumerate_page_elements(page)
    discovery["tabs_found"] = initial.get("tabs", [])
    print(f"  Tabs: {discovery['tabs_found']}")
    print(f"  Buttons (sample): {initial.get('buttons', [])[:8]}")

    comp = initial.get("compliance", [])
    if comp:
        print(f"  Compliance elements: {comp[:6]}")

    discovery["tab_data"]["(main view)"] = {
        "buttons": initial.get("buttons", []),
        "compliance": comp,
        "screenshot": ss,
    }

    project_url = page.url

    for tab_name in TABS_TO_EXPLORE:
        try:
            # Try exact role=tab first
            tab = page.get_by_role("tab", name=tab_name)
            if await tab.count() == 0:
                tab = page.locator(f"[role='tab']:has-text('{tab_name}')")
            if await tab.count() == 0:
                continue

            if not await tab.first.is_visible(timeout=1500):
                continue

            print(f"  Exploring tab: {tab_name}")
            await tab.first.click()
            await asyncio.sleep(3)

            ss = await take_screenshot(page, pid, f"tab_{tab_name.lower().replace(' ', '_')}")
            discovery["screenshots"].append(ss)

            elements = await enumerate_page_elements(page)
            discovery["tab_data"][tab_name] = {
                "buttons": elements.get("buttons", []),
                "compliance": elements.get("compliance", []),
                "screenshot": ss,
            }

            found = elements.get("compliance", [])
            if found:
                print(f"    -> {len(found)} compliance-keyword elements found")
                for item in found[:4]:
                    print(f"       {item}")

        except Exception as e:
            print(f"  Tab '{tab_name}' error: {e}")

    # Return to project record
    if "pse__Proj__c" in project_url:
        await page.goto(project_url, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(2)

    return discovery


# ─────────────────────────────────────────────
#  HTML REPORT GENERATOR
# ─────────────────────────────────────────────

def img_tag(path: str) -> str:
    if path and os.path.exists(path):
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        return f'<img src="data:image/png;base64,{b64}" class="screenshot">'
    return '<p class="no-img">No screenshot</p>'


def generate_report(discoveries: list) -> str:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    sections = ""

    for d in discoveries:
        pid = d["pid"]

        if d.get("error"):
            sections += f"""
            <div class="pid-card error">
                <h2>{pid}</h2>
                <p class="err">{d["error"]}</p>
            </div>"""
            continue

        tabs_str = ", ".join(d.get("tabs_found", [])) or "<i>none detected</i>"
        tab_blocks = ""

        for tab_name, data in d.get("tab_data", {}).items():
            comp_items = "".join(
                f'<li class="hit">{c}</li>' for c in data.get("compliance", [])
            ) or "<li><i>None</i></li>"
            btn_items = "".join(
                f"<li>{b}</li>" for b in data.get("buttons", [])[:25]
            ) or "<li><i>None</i></li>"

            tab_blocks += f"""
            <div class="tab-card">
                <h3>{tab_name}</h3>
                <div class="two-col">
                    <div>{img_tag(data.get("screenshot", ""))}</div>
                    <div>
                        <h4>Compliance-keyword elements</h4>
                        <ul>{comp_items}</ul>
                        <h4>Buttons visible</h4>
                        <ul class="btn-list">{btn_items}</ul>
                    </div>
                </div>
            </div>"""

        sections += f"""
        <div class="pid-card">
            <h2>{pid} <span class="url">{d.get("url", "")[:70]}</span></h2>
            <p><strong>Tabs on record:</strong> {tabs_str}</p>
            {tab_blocks}
        </div>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>SDC Compliance Discovery — {ts}</title>
<style>
  body {{ font-family: Aptos, 'Segoe UI', sans-serif; font-size: 12pt;
          background: #f0f2f5; margin: 24px; color: #222; }}
  h1 {{ color: #1565A0; border-bottom: 3px solid #1565A0; padding-bottom: 6px; }}
  h2 {{ color: #1565A0; border-bottom: 1px solid #ccc; margin-top: 28px; font-size: 14pt; }}
  h3 {{ color: #333; margin: 14px 0 6px; font-size: 12pt; }}
  h4 {{ color: #555; margin: 8px 0 4px; font-size: 10pt; text-transform: uppercase;
        letter-spacing: .5px; }}
  .pid-card {{ background: #fff; border: 1px solid #ddd; border-radius: 8px;
               padding: 20px 24px; margin: 20px 0; box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
  .pid-card.error {{ border-left: 4px solid #c00; }}
  .err {{ color: #c00; font-weight: bold; }}
  .tab-card {{ border: 1px solid #e4e6ea; border-radius: 6px; padding: 12px 16px;
               margin: 12px 0; }}
  .two-col {{ display: flex; gap: 20px; align-items: flex-start; margin-top: 8px; }}
  .screenshot {{ max-width: 520px; border: 1px solid #ccc; border-radius: 4px; }}
  .no-img {{ color: #999; font-style: italic; }}
  ul {{ margin: 4px 0 8px 18px; padding: 0; }}
  li {{ margin: 2px 0; font-size: 11pt; }}
  li.hit {{ color: #1565A0; font-weight: 600; }}
  .btn-list li {{ color: #444; }}
  .url {{ font-size: 9pt; color: #888; font-weight: normal; word-break: break-all; }}
</style>
</head>
<body>
<h1>SDC Cost &amp; Revenue Compliance — Discovery Report</h1>
<p>Generated: {ts} &nbsp;|&nbsp;
   <strong>Blue items</strong> matched compliance keywords.
   Use these to build the automation selectors.</p>
{sections}
</body>
</html>"""


# ─────────────────────────────────────────────
#  AUTOMATION MODE STUB
#  (Fill in after discovery run confirms selectors)
# ─────────────────────────────────────────────

async def run_compliance(page, pid: str) -> dict:
    """
    Placeholder for the full automation.
    Run discovery mode first, then replace this body with the
    confirmed tab/button selectors from discovery_report.html.
    """
    result = {"pid": pid, "status": "not_implemented", "notes": []}
    print(f"  AUTOMATION MODE not yet configured for {pid}.")
    print("  Run discovery mode first: omit --run flag.")
    return result


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

async def main():
    args = sys.argv[1:]
    automation_mode = "--run" in args
    pids = [a.strip() for a in args if a.strip() and not a.startswith("-")]

    if not pids:
        print("Usage:")
        print("  python cost_revenue_compliance.py F51119 F48852     # discovery mode")
        print("  python cost_revenue_compliance.py --run F51119       # automation mode")
        print("")
        print("Claude fetches your active PIDs from sfdc_get_my_projects() first.")
        sys.exit(1)

    os.makedirs(SCREENSHOT_DIR, exist_ok=True)

    mode_label = "AUTOMATION" if automation_mode else "DISCOVERY"
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Starting {mode_label} MODE — {len(pids)} PID(s)")

    async with async_playwright() as playwright:
        context, page = await launch_and_login(playwright)
        if page is None:
            return

        if automation_mode:
            results = []
            for pid in pids:
                r = await run_compliance(page, pid)
                results.append(r)

            print(f"\n{'='*55}")
            print("  AUTOMATION RESULTS")
            print(f"{'='*55}")
            for r in results:
                icon = "OK " if r["status"] == "success" else "---"
                print(f"  [{icon}] {r['pid']}: {r['status']}")
                for n in r.get("notes", []):
                    print(f"          {n}")

        else:
            discoveries = []
            for pid in pids:
                d = await discover_project(page, pid)
                discoveries.append(d)

            report_html = generate_report(discoveries)
            report_path = os.path.join(OUTPUT_DIR, "discovery_report.html")
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(report_html)

            print(f"\n{'='*55}")
            print("  DISCOVERY COMPLETE")
            print(f"  Report: {report_path}")
            print(f"{'='*55}")

            import subprocess
            subprocess.Popen(["start", "", report_path], shell=True)

        print("\n  Browser closes in 20 seconds...")
        await asyncio.sleep(20)
        await context.close()


if __name__ == "__main__":
    asyncio.run(main())
