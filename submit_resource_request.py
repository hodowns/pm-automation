#!/usr/bin/env python3
"""
SDC Submit Resource Request Automation
Submits a New Resource Request in Cisco Service Delivery Consoles.

Usage:
    python submit_resource_request.py ^
        --pid F43673 ^
        --role Engineer ^
        --suggested-resource "Aryan Sawhney" ^
        --work-description "Assisting with Implementation" ^
        --start-date 2026-10-29 ^
        --end-date 2027-02-15 ^
        --hours 30 ^
        --architecture Collaboration ^
        --technology "Packaged CCE" ^
        --timezone Americas ^
        --onsite No ^
        --grade 10 ^
        --milestone "Implementation Execution"

Requirements:
    pip install playwright
    python -m playwright install chromium
"""

import argparse
import asyncio
import os
import sys
from datetime import datetime
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

SDC_URL = "https://csone.lightning.force.com/"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.normpath(os.path.join(
    SCRIPT_DIR, "rr_sdc_profile"
))


def iso_to_mmddyyyy(iso_date: str) -> str:
    """Convert 2026-10-29 -> 10/29/2026."""
    d = datetime.strptime(iso_date, "%Y-%m-%d")
    return d.strftime("%m/%d/%Y")


async def screenshot(page, name: str):
    path = os.path.join(SCRIPT_DIR, f"debug_rr_{name}.png")
    try:
        await page.screenshot(path=path)
        print(f"  Screenshot: debug_rr_{name}.png")
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
#  SEARCH & OPEN PROJECT  (same pattern as close_project.py)
# ─────────────────────────────────────────────────────────────────────────────

async def dismiss_any_modal(page):
    """Close any open modal/dialog before navigating."""
    try:
        modal_visible = await page.locator(".slds-modal.slds-fade-in-open").is_visible(timeout=1500)
        if modal_visible:
            print("  Dismissing open modal...")
            await page.keyboard.press("Escape")
            await asyncio.sleep(1)
            # Also try Cancel button
            try:
                cancel = page.get_by_role("button", name="Cancel")
                if await cancel.is_visible(timeout=1000):
                    await cancel.click()
                    await asyncio.sleep(1)
            except Exception:
                pass
    except Exception:
        pass


async def find_and_open_project(page, pid: str) -> bool:
    # Dismiss any leftover modal from a previous run
    await dismiss_any_modal(page)

    print(f"  Searching for {pid}...")

    try:
        search = page.locator(".slds-global-header input[type='search']")
        await search.wait_for(state="visible", timeout=8000)
        await search.click()
    except Exception:
        vp = page.viewport_size or {"width": 1280, "height": 720}
        for attempt in range(5):
            await page.mouse.click(int(vp["width"] * 0.48), 25)
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
    target_page = None
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
        print(f"  ERROR: Could not find project link for {pid}")
        return False

    await project_link.click()
    await asyncio.sleep(5)
    return True


# ─────────────────────────────────────────────────────────────────────────────
#  JAVASCRIPT SHADOW-DOM HELPERS
#  Salesforce LWC uses native shadow DOM — CSS selectors can't pierce it.
#  These JS functions traverse shadow roots to find and interact with elements.
# ─────────────────────────────────────────────────────────────────────────────

# JS utility injected once — recursive shadow-DOM traversal
_JS_UTIL = """
window._sdcFindByLabel = function(labelHint, elementTypes) {
    function walk(root) {
        // Search .slds-form-element containers in this root
        const containers = root.querySelectorAll(
            '.slds-form-element, lightning-input, lightning-textarea, lightning-combobox'
        );
        for (const c of containers) {
            // Check label text in this container OR its shadow root
            const roots = [c];
            if (c.shadowRoot) roots.push(c.shadowRoot);
            for (const r of roots) {
                const labelEls = r.querySelectorAll(
                    'label, span.slds-form-element__label, abbr.slds-required'
                );
                for (const lbl of labelEls) {
                    const txt = lbl.textContent.replace(/[*ⓘ\\s]+/g,' ').trim();
                    if (txt.includes(labelHint)) {
                        // Found container matching this label — look for target element
                        for (const et of elementTypes) {
                            let el = c.querySelector(et);
                            if (!el && c.shadowRoot) el = c.shadowRoot.querySelector(et);
                            if (el) return el;
                        }
                    }
                }
            }
        }
        // Recurse into shadow roots of all children
        for (const el of root.querySelectorAll('*')) {
            if (el.shadowRoot) {
                const found = walk(el.shadowRoot);
                if (found) return found;
            }
        }
        return null;
    }
    return walk(document.body);
};
"""

async def inject_js_util(page):
    """Inject the shadow-DOM traversal utility once per page."""
    await page.evaluate(_JS_UTIL)


async def js_open_combobox(page, label_hint: str) -> dict:
    """
    Find a lightning-combobox by label text (via shadow DOM traversal) and
    click its dropdown trigger. Returns {ok, label} or {error, ...}.
    """
    return await page.evaluate("""
        (hint) => {
            function walk(root) {
                const cbs = root.querySelectorAll('lightning-combobox');
                for (const cb of cbs) {
                    if (!cb.shadowRoot) continue;
                    const lbl = cb.shadowRoot.querySelector(
                        'label, span.slds-form-element__label'
                    );
                    if (lbl) {
                        const txt = lbl.textContent.replace(/[*ⓘ\\s]+/g,' ').trim();
                        if (txt.includes(hint)) {
                            const btn = cb.shadowRoot.querySelector(
                                'button[aria-haspopup="listbox"], div[aria-haspopup="listbox"]'
                            );
                            if (btn) { btn.click(); return {ok:true, label:txt}; }
                            return {error:'no trigger', label:txt};
                        }
                    }
                }
                for (const el of root.querySelectorAll('*')) {
                    if (el.shadowRoot) {
                        const r = walk(el.shadowRoot);
                        if (r) return r;
                    }
                }
                return null;
            }
            return walk(document.body) || {error:'not found', hint:hint};
        }
    """, label_hint)


async def js_fill_input(page, label_hint: str, value: str) -> dict:
    """
    Find an input/textarea by label via shadow DOM and set its value natively,
    dispatching the events LWC components listen to.
    """
    return await page.evaluate("""
        ([hint, val]) => {
            function walk(root) {
                const containers = root.querySelectorAll(
                    '.slds-form-element, lightning-input, lightning-textarea'
                );
                for (const c of containers) {
                    const roots = [c];
                    if (c.shadowRoot) roots.push(c.shadowRoot);
                    for (const r of roots) {
                        const lbls = r.querySelectorAll(
                            'label, span.slds-form-element__label'
                        );
                        for (const lbl of lbls) {
                            const txt = lbl.textContent.replace(/[*ⓘ\\s]+/g,' ').trim();
                            if (txt.includes(hint)) {
                                // Try input first, then textarea
                                let el = c.querySelector(
                                    'input:not([type=hidden]):not([type=checkbox]):not([type=radio])'
                                );
                                if (!el && c.shadowRoot) el = c.shadowRoot.querySelector(
                                    'input:not([type=hidden]):not([type=checkbox]):not([type=radio])'
                                );
                                if (!el) el = c.querySelector('textarea');
                                if (!el && c.shadowRoot) el = c.shadowRoot.querySelector('textarea');
                                if (el) {
                                    // Use native prototype setter to bypass LWC tracking
                                    const proto = el.tagName === 'TEXTAREA'
                                        ? HTMLTextAreaElement.prototype
                                        : HTMLInputElement.prototype;
                                    const setter = Object.getOwnPropertyDescriptor(proto, 'value');
                                    if (setter) setter.set.call(el, val);
                                    else el.value = val;
                                    el.dispatchEvent(new InputEvent('input',
                                        {bubbles:true, composed:true, data:val}));
                                    el.dispatchEvent(new Event('change',
                                        {bubbles:true, composed:true}));
                                    el.dispatchEvent(new Event('blur',
                                        {bubbles:true, composed:true}));
                                    return {ok:true, tag:el.tagName, label:hint};
                                }
                            }
                        }
                    }
                }
                for (const el of root.querySelectorAll('*')) {
                    if (el.shadowRoot) {
                        const r = walk(el.shadowRoot);
                        if (r) return r;
                    }
                }
                return null;
            }
            return walk(document.body) || {error:'not found', hint:hint};
        }
    """, [label_hint, value])


# ─────────────────────────────────────────────────────────────────────────────
#  HIGH-LEVEL FORM HELPERS
# ─────────────────────────────────────────────────────────────────────────────

async def set_combobox(page, label_hint: str, value: str) -> bool:
    """
    Open a lightning-combobox via JS shadow traversal, then pick an option
    using Playwright's ARIA role selector (which also pierces shadow DOM).
    """
    print(f"    [{label_hint}] -> {value}")
    result = await js_open_combobox(page, label_hint)
    if not result.get("ok"):
        print(f"    WARNING: combobox open failed: {result}")
        return False
    await asyncio.sleep(0.5)
    try:
        opt = page.get_by_role("option", name=value, exact=False).first
        await opt.wait_for(state="visible", timeout=5000)
        await opt.click()
        await asyncio.sleep(0.3)
        return True
    except Exception as e:
        print(f"    WARNING: option select failed: {e}")
        return False


async def set_date(page, label_hint: str, iso_date: str) -> bool:
    """
    Click a date input via JS shadow traversal, then keyboard-type the date.
    Date inputs in this form pre-populate; we must select-all and overwrite.
    """
    mmddyyyy = iso_to_mmddyyyy(iso_date)
    print(f"    [{label_hint}] -> {mmddyyyy}")

    # JS: find the input, focus it, select all text
    focused = await page.evaluate("""
        ([hint]) => {
            function walk(root) {
                const cands = root.querySelectorAll('lightning-input');
                for (const c of cands) {
                    const lbl = c.shadowRoot?.querySelector('label');
                    if (lbl && lbl.textContent.includes(hint)) {
                        const inp = c.shadowRoot?.querySelector('input');
                        if (inp) { inp.focus(); inp.select(); return {ok:true}; }
                    }
                }
                for (const el of root.querySelectorAll('*')) {
                    if (el.shadowRoot) {
                        const r = walk(el.shadowRoot);
                        if (r) return r;
                    }
                }
                return null;
            }
            return walk(document.body) || {error:'not found'};
        }
    """, [label_hint])

    if not focused or not focused.get("ok"):
        print(f"    WARNING: date input not found via JS: {focused}")
        return False

    await asyncio.sleep(0.2)
    await page.keyboard.press("Control+a")
    await page.keyboard.type(mmddyyyy)
    await page.keyboard.press("Tab")
    await asyncio.sleep(0.4)
    return True


async def set_text(page, label_hint: str, value: str) -> bool:
    """Fill a text input or textarea via JS shadow traversal."""
    print(f"    [{label_hint}] -> {value}")
    result = await js_fill_input(page, label_hint, value)
    if result and result.get("ok"):
        return True
    print(f"    WARNING: text fill failed: {result}")
    return False


async def add_technology(page, tech_text: str) -> bool:
    """
    After Architecture is set, find a matching technology in the dual-listbox
    source list (role="option") and move it to Selected.

    The dual-listbox items ARE visible as ARIA options — we use get_by_role
    which pierces shadow DOM, scoped to text that matches our keywords.
    """
    print(f"    [Technologies] -> {tech_text}")
    keywords = [k.strip().lower() for k in tech_text.replace("/", ",").split(",")]

    # Wait for the technology list to dynamically populate
    await asyncio.sleep(2.5)

    # Log available options to help with debugging
    try:
        all_opts = page.get_by_role("option")
        count = await all_opts.count()
        print(f"    Visible role=option items: {count}")
        for i in range(min(count, 10)):
            txt = await all_opts.nth(i).inner_text(timeout=300)
            print(f"      [{i}] {txt.strip()[:50]}")
    except Exception as e:
        print(f"    (could not enumerate options: {e})")

    # Find a matching option in the source list
    all_opts = page.get_by_role("option")
    count = await all_opts.count()
    matched = False
    for i in range(count):
        try:
            opt = all_opts.nth(i)
            opt_text = (await opt.inner_text(timeout=300)).strip().lower()
            if any(kw in opt_text for kw in keywords):
                await opt.click()
                print(f"    Clicked tech option: {opt_text}")
                matched = True
                break
        except Exception:
            continue

    if not matched:
        print(f"    WARNING: no matching tech option for '{tech_text}'")
        return False

    # Click the right-arrow "Move to Selected" button
    # Try by title first, then by position (it's the first button after the source listbox)
    for btn_sel in [
        'button[title="Move selection to Selected"]',
        'button[title*="Move selection"]',
        'button[title*="Move to Selected"]',
    ]:
        try:
            btn = page.locator(btn_sel).first
            if await btn.is_visible(timeout=1500):
                await btn.click()
                await asyncio.sleep(0.5)
                print(f"    Moved to Selected.")
                return True
        except Exception:
            pass

    # JS fallback: find the move-right button by traversing the dual-listbox shadow root
    moved = await page.evaluate("""
        () => {
            function walk(root) {
                const lbs = root.querySelectorAll('lightning-dual-listbox');
                for (const lb of lbs) {
                    if (!lb.shadowRoot) continue;
                    const btns = lb.shadowRoot.querySelectorAll('button');
                    for (const btn of btns) {
                        const t = (btn.getAttribute('title') || '').toLowerCase();
                        if (t.includes('move') || t.includes('selected')) {
                            btn.click();
                            return {ok:true, title:btn.getAttribute('title')};
                        }
                    }
                }
                for (const el of root.querySelectorAll('*')) {
                    if (el.shadowRoot) {
                        const r = walk(el.shadowRoot);
                        if (r) return r;
                    }
                }
                return null;
            }
            return walk(document.body) || {error:'move button not found'};
        }
    """)
    if moved and moved.get("ok"):
        print(f"    Moved via JS: {moved}")
        await asyncio.sleep(0.5)
        return True

    print(f"    WARNING: could not find move-right button")
    return False


async def set_suggested_resource(page, name: str) -> bool:
    """Type the resource name in the Suggested Resource lookup field."""
    print(f"    [Suggested Resource] -> {name}")
    first_name = name.split()[0]

    # JS: find the Suggested Resource input and focus it
    focused = await page.evaluate("""
        ([hint, firstName]) => {
            function walk(root) {
                const all = root.querySelectorAll('.slds-form-element');
                for (const c of all) {
                    const lbls = c.querySelectorAll('label, span.slds-form-element__label');
                    for (const lbl of lbls) {
                        if (lbl.textContent.includes(hint)) {
                            const inp = c.querySelector('input');
                            if (inp) { inp.focus(); return {ok:true}; }
                        }
                    }
                }
                for (const el of root.querySelectorAll('*')) {
                    if (el.shadowRoot) {
                        const r = walk(el.shadowRoot);
                        if (r) return r;
                    }
                }
                return null;
            }
            return walk(document.body) || {error:'not found'};
        }
    """, ["Suggested Resource", first_name])

    if not focused or not focused.get("ok"):
        print(f"    WARNING: Suggested Resource input not found: {focused}")
        return False

    await asyncio.sleep(0.2)
    await page.keyboard.type(first_name)
    await asyncio.sleep(2.5)

    # Autocomplete
    for sel in [
        f"[role='option']:has-text('{name}')",
        f"[role='option']:has-text('{first_name}')",
        f".slds-listbox__item:has-text('{first_name}')",
    ]:
        try:
            opt = page.locator(sel).first
            if await opt.is_visible(timeout=2000):
                await opt.click()
                await asyncio.sleep(0.4)
                print(f"    Selected: {name}")
                return True
        except Exception:
            pass

    print(f"    WARNING: autocomplete didn't appear — name typed, no selection made")
    return False


async def select_milestone(page, milestone_text: str) -> bool:
    """Select an Activity Milestone radio in the scrolled-down section."""
    print(f"    [Activity Milestone] -> {milestone_text}")
    keywords = milestone_text.lower().split()

    try:
        content = page.locator(".slds-modal__content").first
        await content.evaluate("el => el.scrollTop = el.scrollHeight")
        await asyncio.sleep(1)
    except Exception:
        pass

    for row_sel in ["tr", "li", "div[data-id]", ".slds-listbox__item"]:
        try:
            rows = page.locator(row_sel)
            count = await rows.count()
            for i in range(count):
                row = rows.nth(i)
                try:
                    row_text = (await row.inner_text(timeout=400)).lower()
                    if all(kw in row_text for kw in keywords[:2]):
                        radio = row.locator("input[type='radio'], input[type='checkbox']").first
                        if await radio.is_visible(timeout=400):
                            await radio.click()
                        else:
                            await row.click()
                        print(f"    Selected milestone.")
                        await asyncio.sleep(0.3)
                        return True
                except Exception:
                    continue
        except Exception:
            pass

    print(f"    WARNING: Milestone '{milestone_text}' not found")
    return False


# ─────────────────────────────────────────────────────────────────────────────
#  MODAL OPENERS — project-level vs task-level navigation paths
# ─────────────────────────────────────────────────────────────────────────────

def milestone_similarity(a: str, b: str) -> float:
    """Rough word-overlap similarity between two strings (case-insensitive)."""
    a_words = set(a.lower().split())
    b_words = set(b.lower().split())
    if not a_words:
        return 0.0
    return len(a_words & b_words) / len(a_words)


async def open_modal_project_level(page) -> bool:
    """Resources tab → New Resource Request button → modal open."""
    print("  [Project-level] Opening Resources tab...")
    for attempt in range(3):
        try:
            tab = page.get_by_role("tab", name="Resources")
            await tab.wait_for(state="visible", timeout=10000)
            await tab.click()
            await asyncio.sleep(3)
            break
        except PlaywrightTimeout:
            print(f"  Resources tab not ready (attempt {attempt+1})")
            await asyncio.sleep(3)
    else:
        print("  ERROR: Resources tab not found")
        return False

    print("  Clicking New Resource Request...")
    try:
        btn = page.get_by_role("button", name="New Resource Request")
        await btn.wait_for(state="visible", timeout=15000)
        await btn.click()
    except PlaywrightTimeout:
        print("  ERROR: 'New Resource Request' button not found")
        await screenshot(page, "error_no_button")
        return False

    await asyncio.sleep(4)
    return True


async def _scan_and_click_link(page, hint: str, label: str):
    """Scan all frames for a link matching hint, click it. Returns (frame, text) or (None, None)."""
    keywords = [w for w in hint.split() if len(w) > 3] or hint.split()
    best_link = None
    best_score = 0.0
    best_txt = ""
    best_frame = None

    for frame in page.frames:
        for kw in keywords:
            try:
                candidates = frame.locator(f"a:has-text('{kw}')")
                count = await candidates.count()
                for i in range(count):
                    try:
                        txt = (await candidates.nth(i).inner_text(timeout=300)).strip()
                        if not txt or len(txt) > 120:
                            continue
                        score = milestone_similarity(hint, txt)
                        if score > 0.3:
                            print(f"    Frame {frame.name!r}: {txt!r}  score={score:.2f}")
                        if score > best_score:
                            best_score = score
                            best_link = candidates.nth(i)
                            best_txt = txt
                            best_frame = frame
                    except Exception:
                        continue
            except Exception:
                continue
        if best_score >= 0.6:
            break

    if best_link is None or best_score < 0.4:
        print(f"  ERROR: No {label} matching {hint!r} (best={best_score:.2f})")
        await screenshot(page, f"error_no_{label.replace(' ', '_')}")
        return None, None

    print(f"  Clicking {label}: {best_txt!r} (score={best_score:.2f})")
    await best_link.click()
    return best_frame, best_txt


async def open_modal_task_level(page, task_path) -> bool:
    """
    Work Planner → navigate task path → click 'Create Resource Request' icon → modal open.
    task_path: str (single hint) or list of str (hierarchical path, e.g.
               ["Outcome Packages", "Solution Design Development", "ACI Architectural Review"]).
    Intermediate levels are clicked to expand the tree; the CRR icon is clicked at the last level.
    """
    if isinstance(task_path, str):
        task_path = [task_path] if task_path else []
    if not task_path:
        print("  ERROR: task_path is empty")
        return False

    print("  [Task-level] Opening Work Planner...")
    # Three buttons share "Work Planner" — use the neutral-style launch button specifically
    # If Work Planner tab is already open, just wait for it to be visible
    try:
        wp_btn = page.locator(
            "button.slds-button_neutral:has-text('Work Planner'),"
            "button[kx-scope='button-neutral']:has-text('Work Planner')"
        ).first
        if await wp_btn.is_visible(timeout=3000):
            await wp_btn.click()
            await asyncio.sleep(4)
        else:
            print("  Work Planner button not visible — may already be open")
    except Exception as e:
        print(f"  Work Planner button click error: {e}")

    # Wait for Work Planner to load — look for Billing Milestone or All Packages header
    print("  Work Planner loading...")
    try:
        await page.locator(
            "text=Billing Milestone, text=All Packages, text=All"
        ).first.wait_for(state="visible", timeout=15000)
        await asyncio.sleep(1)
    except Exception:
        await asyncio.sleep(3)

    # Navigate intermediate levels (all but the last) by clicking each link to expand the tree
    for i, level in enumerate(task_path[:-1]):
        print(f"  Expanding level {i+1}/{len(task_path)-1}: {level!r}")
        frame_ref, _ = await _scan_and_click_link(page, level, f"level-{i+1}")
        if frame_ref is None:
            return False
        await asyncio.sleep(4)
        await screenshot(page, f"task_level_{i+1}_opened")

    # Final level: find "Create Resource Requests" icon in the correct frame.
    # Re-scan all frames to locate the frame containing the target task after navigation.
    milestone_hint = task_path[-1]
    print(f"  Scanning for final task: {milestone_hint!r}")
    keywords = [w for w in milestone_hint.split() if len(w) > 3] or milestone_hint.split()
    best_frame = None

    for frame in page.frames:
        for kw in keywords:
            try:
                candidates = frame.locator(f"a:has-text('{kw}')")
                count = await candidates.count()
                for i in range(count):
                    try:
                        txt = (await candidates.nth(i).inner_text(timeout=300)).strip()
                        if txt and len(txt) <= 120:
                            score = milestone_similarity(milestone_hint, txt)
                            if score >= 0.4:
                                best_frame = frame
                                print(f"    Final task found in frame {frame.name!r}: {txt!r}")
                                break
                    except Exception:
                        continue
            except Exception:
                continue
            if best_frame:
                break
        if best_frame:
            break

    # Fallback: use any frame that has CRR icons
    if best_frame is None:
        print("  Final task link not found by name — falling back to frame with CRR icons")
        for frame in page.frames:
            try:
                crr_count = await frame.locator('lightning-icon[title="Create Resource Requests"]').count()
                if crr_count > 0:
                    best_frame = frame
                    print(f"    CRR icons found in frame {frame.name!r} ({crr_count} icon(s))")
                    break
            except Exception:
                continue

    if best_frame is None:
        print(f"  ERROR: Could not locate frame for {milestone_hint!r} after navigation")
        await screenshot(page, "error_no_final_frame")
        return False

    await screenshot(page, "task_milestone_opened")

    # Milestone detail loaded — sub-tasks appear with 'Create Resource Request' icons.
    # Find the sub-task row matching milestone_hint (~80% similarity) and click its icon.
    # Search in the same frame the milestone was found in, plus all frames.
    print("  Looking for sub-task 'Create Resource Request' icon...")

    # Diagnostic: traverse shadow DOM inside frame to find Create Resource Request button
    try:
        btn_info = await best_frame.evaluate("""
            () => {
                const results = [];
                function walk(root) {
                    root.querySelectorAll('[title], button, a').forEach(el => {
                        const t = (el.title || el.getAttribute('aria-label') || '').trim();
                        if (t && t.length < 60) results.push({tag: el.tagName, title: t});
                    });
                    root.querySelectorAll('*').forEach(el => {
                        if (el.shadowRoot) walk(el.shadowRoot);
                    });
                }
                walk(document.body);
                return [...new Map(results.map(r => [r.title, r])).values()].slice(0, 30);
            }
        """)
        print(f"  Shadow-traversal buttons: {btn_info}")
    except Exception as e:
        print(f"  Button dump error: {e}")

    # The icon is a LIGHTNING-ICON with title="Create Resource Requests" (plural).
    # Use Playwright's frame locator (pierces shadow DOM) for a real mouse click.
    # Scope to the sub-task row matching our hint when multiple icons exist.

    # Find all Create Resource Requests icons in the frame
    crr_icons = best_frame.locator('lightning-icon[title="Create Resource Requests"]')
    count = await crr_icons.count()
    print(f"  Found {count} 'Create Resource Requests' icon(s) in frame")

    if count == 0:
        print("  ERROR: 'Create Resource Requests' icon not found in frame")
        await screenshot(page, "error_no_crr_icon")
        return False

    if count == 1:
        target = crr_icons.first
    else:
        # Multiple icons — find the one on the sub-task row matching the hint
        best_icon_idx = 0
        best_icon_score = 0.0
        for j in range(count):
            try:
                row_txt = await crr_icons.nth(j).evaluate(
                    "el => (el.closest('tr, [class*=\"row\"], li') || el.parentElement?.parentElement || el.parentElement)?.textContent?.substring(0, 100) || ''"
                )
                s = milestone_similarity(milestone_hint, row_txt)
                print(f"    Icon[{j}] row text={row_txt[:50]!r} score={s:.2f}")
                if s > best_icon_score:
                    best_icon_score = s
                    best_icon_idx = j
            except Exception:
                pass
        target = crr_icons.nth(best_icon_idx)

    # Playwright click fires real mouse events, which propagates through LWC correctly
    await target.click(timeout=5000)
    print("  Clicked 'Create Resource Requests' icon.")
    await asyncio.sleep(4)
    return best_frame  # return the frame — modal opens inside it


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN FORM FLOW
# ─────────────────────────────────────────────────────────────────────────────

async def submit_resource_request(page, args) -> bool:
    """Navigate to the New Resource Request modal (project or task level), fill it, submit."""

    # Choose navigation path
    # task-level returns the VF frame (modal is inside it); project-level returns True
    if args.task_level:
        # Build task_path: prefer --task-path (list), else fall back to --milestone (str)
        task_path = getattr(args, "task_path", None) or ([args.milestone] if args.milestone else [])
        result = await open_modal_task_level(page, task_path)
        if not result:
            return False
        ctx = result  # frame containing the modal
    else:
        ok = await open_modal_project_level(page)
        if not ok:
            return False
        ctx = page  # modal is on the main page

    await asyncio.sleep(1)
    modal = ctx.locator(".slds-modal__container").first
    try:
        await modal.wait_for(state="visible", timeout=10000)
    except PlaywrightTimeout:
        print("  ERROR: Modal did not open")
        await screenshot(page, "error_no_modal")
        return False

    print("  Modal open — filling form...")
    await inject_js_util(ctx)
    await screenshot(page, "01_modal_open")

    # ── DIAGNOSTIC: enumerate every lightning-combobox and native select ──────
    print("\n  [DIAG] Enumerating form elements...")
    try:
        cb_count = await ctx.locator("lightning-combobox").count()
        print(f"    lightning-combobox count: {cb_count}")
        for i in range(cb_count):
            try:
                txt = await ctx.locator("lightning-combobox").nth(i).inner_text(timeout=800)
                print(f"      CB[{i}]: {repr(txt[:60])}")
            except Exception as e:
                print(f"      CB[{i}]: (no text — {type(e).__name__})")
    except Exception as e:
        print(f"    combobox enum error: {e}")

    try:
        sel_count = await ctx.locator("select").count()
        print(f"    <select> count: {sel_count}")
        for i in range(sel_count):
            try:
                opts = await ctx.locator("select").nth(i).evaluate(
                    "el => Array.from(el.options).slice(0,6).map(o=>o.text)"
                )
                print(f"      SEL[{i}]: {opts}")
            except Exception as e:
                print(f"      SEL[{i}]: error — {e}")
    except Exception as e:
        print(f"    select enum error: {e}")

    # Enumerate form-relevant inputs (skip range sliders and search)
    try:
        form_inps = ctx.locator(
            "input:not([type=range]):not([type=search])"
            ":not([type=hidden]):not([type=checkbox]):not([type=radio])"
        )
        inp_count = await form_inps.count()
        print(f"    form inputs (non-range): {inp_count}")
        for i in range(min(inp_count, 20)):
            try:
                info = await form_inps.nth(i).evaluate(
                    "el => ({type: el.type, val: el.value.substring(0,20), "
                    "placeholder: el.placeholder, "
                    "ariaLabel: el.getAttribute('aria-label') || ''})"
                )
                print(f"      INP[{i}]: {info}")
            except Exception as e:
                print(f"      INP[{i}]: error — {e}")
    except Exception as e:
        print(f"    input enum error: {e}")

    try:
        li_count = await ctx.locator("lightning-input").count()
        ta_count = await ctx.locator("lightning-textarea").count()
        print(f"    lightning-input: {li_count}  |  lightning-textarea: {ta_count}")
    except Exception as e:
        print(f"    lightning-input/textarea count error: {e}")
    print("  [DIAG] done.\n")

    # Scroll to top
    try:
        content = ctx.locator(".slds-modal__content").first
        await content.evaluate("el => el.scrollTop = 0")
        await asyncio.sleep(0.5)
    except Exception:
        pass

    # ── POSITIONAL COMBOBOX HELPER (uses nth index, bypasses text-filter issue) ─
    # ctx = page (project-level) OR VF frame (task-level) — all form ops use this
    async def pick_cb(idx: int, value: str, label: str):
        """Click the Nth lightning-combobox trigger and select value."""
        print(f"    [CB{idx} {label}] -> {value}")
        try:
            cb = ctx.locator("lightning-combobox").nth(idx)
            trigger = cb.locator("button[aria-haspopup='listbox']").first
            await trigger.click(timeout=4000)
            await asyncio.sleep(0.5)
            opt = ctx.get_by_role("option", name=value, exact=False).first
            await opt.wait_for(state="visible", timeout=5000)
            await opt.click()
            await asyncio.sleep(0.3)
            print(f"      ✓")
            return True
        except Exception as e:
            print(f"      FAILED: {e}")
            return False

    async def fill_nth_input(idx: int, value: str, label: str, is_date=False):
        """Fill the Nth form input (non-range) and fire change events."""
        filt = ("input:not([type=range]):not([type=search])"
                ":not([type=hidden]):not([type=checkbox]):not([type=radio])")
        print(f"    [INP{idx} {label}] -> {value}")
        try:
            inp = ctx.locator(filt).nth(idx)
            await inp.click(timeout=3000)
            await asyncio.sleep(0.2)
            await ctx.evaluate("() => {}") if False else None  # keep ctx reference
            await page.keyboard.press("Control+a")
            await page.keyboard.type(value)
            await page.keyboard.press("Tab")
            await asyncio.sleep(0.3)
            print(f"      ✓")
            return True
        except Exception as e:
            print(f"      FAILED: {e}")
            return False

    # ── COMMON REQUIREMENTS ──────────────────────────────────────────────────
    # CB indices confirmed from diagnostic:
    #   CB[0]  = Resource Role
    #   CB[1]  = Intended LOE Type (pre-set "Requested Hours", skip)
    #   CB[2]  = Start Date Flexibility → Ontime (hardcoded)
    #   CB[3]  = Architecture
    print("\n  -- Common Requirements --")
    await pick_cb(0, args.role, "Resource Role")

    # Dates — indices TBD from diagnostic output above; try 0 and 1 for form date inputs
    await fill_nth_input(0, iso_to_mmddyyyy(args.start_date), "Start Date")
    await fill_nth_input(1, iso_to_mmddyyyy(args.end_date), "End Date")

    # Start Date Flexibility (CB[2]) — always Ontime
    await pick_cb(2, "Ontime", "Start Date Flexibility")

    # Requested Hours — INP[2]; Work Description is a lightning-textarea
    await fill_nth_input(2, str(args.hours), "Requested Hours")

    print(f"    [Work Description] -> {args.work_description}")
    try:
        ta = ctx.locator("lightning-textarea").first.locator("textarea")
        await ta.click(timeout=3000)
        await page.keyboard.press("Control+a")
        await page.keyboard.type(args.work_description)
        await asyncio.sleep(0.3)
        print(f"      ✓")
    except Exception as e:
        print(f"      Work Description FAILED: {e}")

    # Lead time checkbox — use JS to set checked=true directly (bypasses LWC pointer issues)
    print(f"    [Lead time < 14 days checkbox]")
    result = await ctx.evaluate("""
        () => {
            function walk(root) {
                const cbs = root.querySelectorAll('input[name="leadTime"]');
                for (const cb of cbs) {
                    if (!cb.checked) {
                        const nativeSetter = Object.getOwnPropertyDescriptor(
                            HTMLInputElement.prototype, 'checked');
                        if (nativeSetter) nativeSetter.set.call(cb, true);
                        else cb.checked = true;
                        cb.dispatchEvent(new Event('change', {bubbles:true, composed:true}));
                        cb.dispatchEvent(new Event('input', {bubbles:true, composed:true}));
                    }
                    return {ok:true, checked:cb.checked};
                }
                for (const el of root.querySelectorAll('*')) {
                    if (el.shadowRoot) { const r = walk(el.shadowRoot); if (r) return r; }
                }
                return null;
            }
            return walk(document.body) || {error:'not found'};
        }
    """)
    print(f"      {result}")

    # Architecture & Technology — project-level only.
    # Task-level: these are pre-filled from the milestone (read-only), skip.
    if not args.task_level:
        print(f"    [Architecture CB3] -> {args.architecture}")
        try:
            cb = ctx.locator("lightning-combobox").filter(has_text="Architecture").first
            trigger = cb.locator("button[aria-haspopup='listbox']").first
            await trigger.click(timeout=4000)
            await asyncio.sleep(0.5)
            opt = ctx.get_by_role("option", name=args.architecture, exact=False).first
            await opt.wait_for(state="visible", timeout=6000)
            await opt.click()
            await asyncio.sleep(1.5)
            print(f"      ✓")
        except Exception as e:
            print(f"      Architecture FAILED: {e}")

        if args.technology:
            await add_technology(ctx, args.technology)
    else:
        print("    [Architecture/Technology] pre-filled from milestone — skipped")

    await screenshot(page, "02_common_filled")

    # ── RESOURCE 1 REQUIREMENTS ──────────────────────────────────────────────
    # CB indices from diagnostic (Resource 1 section):
    #   CB[4]  = Resource Allocation Reason  (hardcoded: Assign new resource)
    #   CB[5]  = Delivery Organization       (optional, skip)
    #   CB[6]  = Onsite Requirement          (hardcoded: No)
    #   CB[7]  = Time Zone Requirement       (hardcoded: Americas)
    #   CB[8]  = Citizenship Requirement     (hardcoded: No)
    #   CB[9]  = Time Zone actual            (optional, skip)
    #   CB[10] = Country                     (optional, skip)
    #   CB[11] = Resource Grade              (user-supplied)
    #   CB[12] = Security Clearance          (hardcoded: No)
    #   CB[14] = Scheduling Mode             (pre-set "Scheduled", skip)
    print("\n  -- Resource 1 Requirements --")

    # ── Confirmed CB mapping for Resource 1 (from systematic diagnostic scan):
    #   CB[4]  = Resource Allocation Reason  → Assign new resource
    #   CB[5]  = Delivery Organization       (optional, skip)
    #   CB[6]  = Onsite Requirement          → No  (Yes/No)
    #   CB[7]  = Citizenship Requirement     → No  (Yes/No)
    #   CB[8]  = Time Zone Requirement       → No  (Yes/No — NOT Americas/EMEA, form updated)
    #   CB[9]  = Time Zone actual            (optional, skip)
    #   CB[10] = Direct Assignment Reason    (optional, skip)
    #   CB[11] = Resource Grade              → user grade
    #   CB[12] = Country                     (optional, skip)
    #   CB[13] = Security Clearance          → No  (Yes/No)
    #   CB[14] = Scheduling Mode             (pre-set Scheduled, skip)
    # CB mapping differs between project-level (18 CBs) and task-level (15 CBs).
    # Task-level: Architecture/Technology are pre-filled, so Resource 1 fields shift left by 1.
    # Hardcoded defaults — change these if org differs (see README)
    if args.task_level:
        # Task-level CB mapping (15 total, Resource 1 starts at CB[3]):
        #   CB[3]  = Resource Allocation Reason
        #   CB[4]  = Delivery Organization (skip)
        #   CB[5]  = Onsite Requirement
        #   CB[6]  = Citizenship Requirement
        #   CB[7]  = Time Zone Requirement
        #   CB[8]  = Time Zone actual (skip)
        #   CB[9]  = Country (skip)
        #   CB[10] = Resource Grade
        #   CB[11] = Security Clearance (or Country — TBD from scan)
        #   CB[12] = Security Clearance (asterisk)
        #   CB[13] = Scheduling Mode (pre-set Scheduled — confirm it)
        await pick_cb(3,  "Assign new resource",      "Resource Allocation Reason")
        await pick_cb(5,  "No",                        "Onsite Requirement")
        await pick_cb(6,  "No",                        "Citizenship Requirement")
        await pick_cb(7,  "No",                        "Time Zone Requirement")
        await pick_cb(10, str(args.grade),             "Resource Grade")
        await pick_cb(12, "No",                        "Security Clearance")
        await pick_cb(13, "Scheduled",                 "Scheduling Mode")
    else:
        # Project-level CB mapping (18 total, Resource 1 starts at CB[4]):
        await pick_cb(4,  "Assign new resource",      "Resource Allocation Reason")
        await pick_cb(6,  "No",                        "Onsite Requirement")
        await pick_cb(7,  "No",                        "Citizenship Requirement")
        await pick_cb(8,  "No",                        "Time Zone Requirement")
        await pick_cb(11, str(args.grade),             "Resource Grade")
        await pick_cb(13, "No",                        "Security Clearance")
        # Direct Assignment Reason (CB[10]) — always "Resource on Project Team"
        await pick_cb(10, "Resource on Project Team",  "Direct Assignment Reason")

    # Suggested Resource — use set_suggested_resource (JS shadow traversal + autocomplete)
    if args.suggested_resource:
        await set_suggested_resource(ctx, args.suggested_resource)

    # Milestone is pre-selected for task-level (from the row we clicked)
    if args.milestone and not args.task_level:
        await select_milestone(ctx, args.milestone)

    await screenshot(page, "03_resource1_filled")

    # ── NEXT ────────────────────────────────────────────────────────────────
    print("\n  Clicking Next...")
    try:
        # Two "Next" buttons exist: a disabled nav arrow and the actual submit button
        # The submit button is a slds-button_brand; use .last to get the real one
        next_btn = ctx.locator("button.slds-button_brand:has-text('Next'), button[type='submit']:has-text('Next')").first
        if not await next_btn.is_visible(timeout=3000):
            next_btn = ctx.get_by_role("button", name="Next").last
        await next_btn.wait_for(state="visible", timeout=8000)
        await next_btn.click()
        await asyncio.sleep(5)
    except PlaywrightTimeout:
        print("  ERROR: Next button not found — check for validation errors")
        await screenshot(page, "error_next")
        return False

    await screenshot(page, "04_review_screen")

    # Detect actual step-2 Review page: it has "Submit" + "Previous" buttons, no Input Error
    try:
        submit_visible = await ctx.get_by_role("button", name="Submit").is_visible(timeout=2000)
        error_visible = await ctx.locator("text=Input Error").is_visible(timeout=500)
        on_review = submit_visible and not error_visible
    except Exception:
        on_review = False

    if not on_review:
        print("  Not on review page — checking for lead time error...")
        try:
            close_btn = ctx.get_by_role("button", name="Close").first
            if await close_btn.is_visible(timeout=2000):
                await close_btn.click(timeout=3000)
                await asyncio.sleep(1)
        except Exception:
            pass

        # Re-check lead time via JS
        result2 = await ctx.evaluate("""
            () => {
                function walk(root) {
                    const cbs = root.querySelectorAll('input[name="leadTime"]');
                    for (const cb of cbs) {
                        if (!cb.checked) {
                            const ns = Object.getOwnPropertyDescriptor(
                                HTMLInputElement.prototype, 'checked');
                            if (ns) ns.set.call(cb, true);
                            else cb.checked = true;
                            cb.dispatchEvent(new Event('change', {bubbles:true, composed:true}));
                        }
                        return {ok:true, checked:cb.checked};
                    }
                    for (const el of root.querySelectorAll('*')) {
                        if (el.shadowRoot) { const r = walk(el.shadowRoot); if (r) return r; }
                    }
                    return null;
                }
                return walk(document.body) || {error:'not found'};
            }
        """)
        print(f"  Lead time checkbox JS result: {result2} — retrying Next...")

        # Close the error dialog if still open, then click Next
        try:
            err_section = ctx.locator("section[data-id='error-modal-main']")
            if await err_section.is_visible(timeout=1000):
                close_x = err_section.locator("button[title='Close']").first
                if await close_x.is_visible(timeout=1000):
                    await close_x.click(timeout=2000)
                await asyncio.sleep(1)
        except Exception:
            pass

        try:
            next_btn2 = ctx.locator("button.slds-button_brand:has-text('Next')").first
            if not await next_btn2.is_visible(timeout=2000):
                next_btn2 = ctx.get_by_role("button", name="Next").last
            await next_btn2.click(timeout=10000)
            await asyncio.sleep(5)
            await screenshot(page, "04b_after_leadtime_retry")
        except Exception as e:
            print(f"  Next retry failed: {e}")

    print("  Review screen reached.")

    # ── SUBMIT ──────────────────────────────────────────────────────────────
    if args.dry_run:
        print("  DRY RUN — stopping before Submit. Review screenshot saved.")
        await screenshot(page, "dry_run_review")
        return True

    print("  Clicking Submit...")
    try:
        submit_btn = ctx.get_by_role("button", name="Submit")
        await submit_btn.wait_for(state="visible", timeout=10000)
        await submit_btn.click()
        await asyncio.sleep(6)
    except PlaywrightTimeout:
        print("  ERROR: Submit button not found")
        await screenshot(page, "error_submit")
        return False

    await screenshot(page, "05_post_submit")

    # Detect success
    try:
        page_text = await page.evaluate("() => document.body.innerText")
        if any(s in page_text for s in ["successfully created", "RR-", "Resource Request"]):
            print("  Resource request created successfully.")
    except Exception:
        pass

    return True


# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

async def main():
    parser = argparse.ArgumentParser(description="Submit SDC Resource Request")
    parser.add_argument("--pid", required=True, help="Project PID e.g. F43673")
    parser.add_argument("--role", default="Engineer", help="Resource role")
    parser.add_argument("--suggested-resource", default="", help="Full name of specific person")
    parser.add_argument("--work-description", default="", help="Work description text")
    parser.add_argument("--start-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--end-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--hours", type=int, required=True, help="Requested hours")
    parser.add_argument("--architecture", default="", help="e.g. Collaboration")
    parser.add_argument("--technology", default="", help="Partial tech name e.g. 'Packaged CCE'")
    parser.add_argument("--timezone", default="Americas", help="Americas / EMEA / APAC / No")
    parser.add_argument("--onsite", default="No", help="Yes / No")
    parser.add_argument("--grade", default="", help="Resource grade e.g. 10")
    parser.add_argument("--milestone", default="", help="Activity milestone name")
    parser.add_argument("--task-path", default="",
                        help="Comma-separated task path for task-level navigation, e.g. "
                             "'Outcome Packages,Solution Design Development,ACI Architectural Review'")
    parser.add_argument("--task-level", action="store_true",
                        help="Use Work Planner task-level navigation instead of Resources tab")
    parser.add_argument("--dry-run", action="store_true",
                        help="Fill the form and reach review page but do NOT click Submit")
    args = parser.parse_args()
    # Convert --task-path CSV string to list; strip whitespace around each component
    if args.task_path:
        args.task_path = [t.strip() for t in args.task_path.split(",") if t.strip()]
    else:
        args.task_path = []

    os.makedirs(PROFILE_DIR, exist_ok=True)

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False,
            slow_mo=300,
        )
        page = await context.new_page()

        print("Opening SDC...")
        try:
            await page.goto(SDC_URL)
        except Exception:
            pass  # network redirect / change is normal

        print("Waiting for SSO login (up to 5 min) — log in via the browser that just opened...")
        deadline = asyncio.get_event_loop().time() + 300
        logged_in = False
        last_url = ""
        while asyncio.get_event_loop().time() < deadline:
            await asyncio.sleep(3)
            url = page.url
            if url != last_url:
                print(f"  URL: {url[:80]}")
                last_url = url
            if ("force.com" in url or "salesforce.com" in url) and "login" not in url.lower():
                logged_in = True
                print(f"  Logged in.")
                break

        if not logged_in:
            print("ERROR: Login timed out — re-run the script and complete SSO within 5 minutes.")
            await context.close()
            return

        if "lightning.force.com" not in page.url:
            try:
                await page.goto("https://csone.lightning.force.com/", wait_until="commit", timeout=30000)
            except Exception:
                pass  # redirect to /one/one.app is expected
            await asyncio.sleep(4)

        print("  Waiting for PM Console...")
        for _ in range(20):
            await asyncio.sleep(3)
            try:
                ok = await page.locator(
                    "input[type='search'], .slds-global-header input"
                ).first.is_visible(timeout=1000)
                if ok:
                    print("  Console ready.")
                    break
            except Exception:
                pass

        print(f"\n{'='*55}")
        print(f"  Submitting RR: {args.role} on {args.pid}")
        if args.suggested_resource:
            print(f"  Suggested resource: {args.suggested_resource}")
        print(f"{'='*55}\n")

        if not await find_and_open_project(page, args.pid):
            print(f"ERROR: Could not open project {args.pid}")
            await context.close()
            return

        success = await submit_resource_request(page, args)

        if success:
            print(f"\n  Done — resource request submitted on {args.pid}.")
        else:
            print(f"\n  Script encountered errors — check debug_rr_*.png files.")

        print("Browser closes in 15 seconds...")
        await asyncio.sleep(15)
        await context.close()


if __name__ == "__main__":
    asyncio.run(main())
