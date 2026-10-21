#!/usr/bin/env python3
"""
OP Cost Forecast — Bulk Submit
Automates the Oracle Projects Cost Forecasting workflow:
  1. Opens Oracle Projects in Edge (or attaches to an existing session)
  2. Navigates to Projects Cost Forecasting
  3. Clicks Select_All_Shown
  4. Waits for all checkboxes to go blue
  5. Clicks Submit
  6. Waits for the Warning dialog, clicks Yes

Usage:
    python op_cost_forecast.py              # opens a new Edge window
    python op_cost_forecast.py --attach     # attaches to an already-open Edge (requires --remote-debugging-port=9222)

To use --attach, start Edge with:
    "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe" --remote-debugging-port=9222
"""

import sys
import time
import subprocess
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException
)

# ── Config ────────────────────────────────────────────────────────────────────

OP_HOME      = "https://wwwin-csfprd.cisco.com/OA_HTML/OA.jsp?OAFunc=OAHOMEPAGE"
LOGIN_WAIT   = 90      # seconds to wait for the user to complete SSO login
PAGE_WAIT    = 30      # seconds to wait for Oracle page loads
CHECKBOX_WAIT = 8      # seconds for checkboxes to turn blue after Select_All_Shown
WARNING_WAIT  = 12     # seconds to wait for the Warning dialog after Submit

# ── Driver setup ─────────────────────────────────────────────────────────────

def get_driver(attach=False):
    opts = Options()
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--start-maximized")

    if attach:
        opts.debugger_address = "localhost:9222"
        print("Attaching to existing Edge session on port 9222...")
    else:
        opts.add_argument("--new-window")

    try:
        # Selenium 4.x: Service auto-downloads msedgedriver if not present
        svc = Service()
        driver = webdriver.Edge(service=svc, options=opts)
        return driver
    except WebDriverException as e:
        print(f"ERROR: Could not start Edge driver.\n{e}")
        print("\nMake sure Edge is installed and msedgedriver is available.")
        print("Download msedgedriver from: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/")
        sys.exit(1)


# ── Helpers ───────────────────────────────────────────────────────────────────

def find_by_text(driver, tag, text, timeout=PAGE_WAIT):
    """Find an element by tag + visible text (case-insensitive contains)."""
    xpath = f"//{tag}[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{text.lower()}')]"
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )


def click_link_text(driver, text, timeout=PAGE_WAIT):
    el = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.LINK_TEXT, text))
    )
    el.click()
    return el


def wait_for_url_contains(driver, fragment, timeout=PAGE_WAIT):
    WebDriverWait(driver, timeout).until(EC.url_contains(fragment))


# ── Main workflow ─────────────────────────────────────────────────────────────

def run(driver, attach=False):
    wait = WebDriverWait(driver, PAGE_WAIT)

    # ── Step 1: Navigate to Oracle Projects home ──────────────────────────────
    if not attach:
        print("Opening Oracle Projects...")
        driver.get(OP_HOME)

        # Wait for SSO/login to complete — page title changes once logged in
        print(f"Waiting up to 180s for SSO login to complete...")
        print("  Log in when Edge opens — script watches for the Oracle Projects URL.")
        try:
            WebDriverWait(driver, 180).until(
                lambda d: "csfprd.cisco.com" in d.current_url
            )
            print(f"  Oracle Projects detected. Page: {driver.title}")
        except TimeoutException:
            print(f"  Timeout. Current URL: {driver.current_url[:80]}")
            print("  Proceeding anyway — you may need to navigate manually.")

    # ── Step 2: Navigate to Cisco AS Project Manager ──────────────────────────
    print("Navigating to Cisco AS Project Manager...")
    try:
        pm_link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(),'Cisco AS Project Manager')]")
        ))
        pm_link.click()
        print("  Clicked 'Cisco AS Project Manager'")
    except TimeoutException:
        # May already be on the right page
        print("  'Cisco AS Project Manager' link not found — may already be on the right page")

    # ── Step 3: Navigate to Projects Cost Forecasting ─────────────────────────
    print("Navigating to Projects Cost Forecasting...")
    try:
        cf_link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(),'Projects Cost Forecasting')]")
        ))
        cf_link.click()
        print("  Clicked 'Projects Cost Forecasting'")
        time.sleep(3)
    except TimeoutException:
        print("  Cost Forecasting link not found — may already be on the Project Details page")

    # ── Step 4: Wait for Project Details page ────────────────────────────────
    print("Waiting for Project Details page...")
    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[contains(text(),'Project Details')]")
        ))
        print("  Project Details page loaded.")
    except TimeoutException:
        print("  Could not confirm Project Details page — proceeding anyway.")

    # ── Step 5: Click Select_All_Shown ────────────────────────────────────────
    print("Clicking 'Select_All_Shown'...")
    selectors = [
        (By.LINK_TEXT, "Select_All_Shown"),
        (By.PARTIAL_LINK_TEXT, "Select_All_Shown"),
        (By.XPATH, "//a[contains(@id,'SelectAll')]"),
        (By.XPATH, "//a[normalize-space(text())='Select_All_Shown']"),
        (By.XPATH, "//*[contains(@onclick,'selectAll')]"),
    ]
    clicked = False
    for by, val in selectors:
        try:
            el = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((by, val)))
            el.click()
            print(f"  Clicked via selector: {val}")
            clicked = True
            break
        except (TimeoutException, NoSuchElementException):
            continue

    if not clicked:
        print("  ERROR: Could not find 'Select_All_Shown'. Waiting 15s for page to fully load...")
        time.sleep(15)
        # One more attempt
        for by, val in selectors:
            try:
                el = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((by, val)))
                el.click()
                print(f"  Clicked on retry via: {val}")
                clicked = True
                break
            except (TimeoutException, NoSuchElementException):
                continue
        if not clicked:
            print("  FAILED: Select_All_Shown still not found. Current URL:", driver.current_url)
            return

    # ── Step 6: Wait for all checkboxes to go blue ───────────────────────────
    print(f"  Waiting {CHECKBOX_WAIT}s for all checkboxes to check...")
    time.sleep(CHECKBOX_WAIT)
    print("  Checkboxes should now be selected.")

    # ── Step 7: Click Submit ──────────────────────────────────────────────────
    print("Clicking 'Submit'...")
    submit_selectors = [
        (By.XPATH, "//button[normalize-space(text())='Submit']"),
        (By.XPATH, "//input[@value='Submit']"),
        (By.XPATH, "//a[normalize-space(text())='Submit']"),
        (By.XPATH, "//*[contains(@id,'Submit')][contains(@type,'submit') or contains(@type,'button') or self::a]"),
        (By.XPATH, "//td[@class='OraButtonText' and contains(text(),'Submit')]"),
    ]
    submitted = False
    for by, val in submit_selectors:
        try:
            el = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((by, val)))
            driver.execute_script("arguments[0].scrollIntoView(true);", el)
            el.click()
            print(f"  Submitted via: {val}")
            submitted = True
            break
        except (TimeoutException, NoSuchElementException):
            continue

    if not submitted:
        # Try JavaScript approach — find by visible text
        try:
            driver.execute_script("""
                var els = document.querySelectorAll('a, button, input');
                for (var i=0; i<els.length; i++) {
                    var t = (els[i].textContent || els[i].value || '').trim();
                    if (t === 'Submit') { els[i].click(); break; }
                }
            """)
            print("  Submitted via JavaScript fallback.")
            submitted = True
        except Exception as e:
            print(f"  ERROR: Could not click Submit. {e}")

    # ── Step 8: Wait for Warning dialog ──────────────────────────────────────
    print(f"  Waiting up to {WARNING_WAIT}s for Warning dialog...")

    # Try browser-level JS alert first (fastest)
    try:
        alert = WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert.accept()
        print("  Browser alert accepted.")
        print("\nDone. Cost forecast submitted successfully.")
        return
    except TimeoutException:
        pass

    # Oracle OAF inline dialog — try multiple selectors
    yes_selectors = [
        (By.XPATH, "//button[normalize-space(text())='Yes']"),
        (By.XPATH, "//input[@value='Yes']"),
        (By.XPATH, "//a[normalize-space(text())='Yes']"),
        (By.XPATH, "//*[contains(@id,'Yes')]"),
        (By.XPATH, "//*[contains(@class,'OraButton') and normalize-space(text())='Yes']"),
        (By.XPATH, "//td[normalize-space(text())='Yes']"),
        # Oracle dialog patterns
        (By.XPATH, "//*[@class='OraButtonText' and text()='Yes']"),
        (By.XPATH, "//span[text()='Yes']/ancestor::a"),
        (By.XPATH, "//span[text()='Yes']/ancestor::button"),
    ]

    yes_btn = None
    deadline = time.time() + WARNING_WAIT
    while time.time() < deadline:
        for by, val in yes_selectors:
            try:
                el = driver.find_element(by, val)
                if el.is_displayed() and el.is_enabled():
                    yes_btn = el
                    break
            except NoSuchElementException:
                continue
        if yes_btn:
            break
        time.sleep(0.5)

    if not yes_btn:
        # Last resort: JavaScript click on any visible Yes button
        try:
            driver.execute_script("""
                var els = document.querySelectorAll('a, button, input, td, span');
                for (var i=0; i<els.length; i++) {
                    var t = (els[i].textContent || els[i].value || '').trim();
                    if (t === 'Yes' && els[i].offsetParent !== null) {
                        els[i].click();
                        return true;
                    }
                }
                return false;
            """)
            print("  Clicked Yes via JavaScript fallback.")
            print("\nDone. Cost forecast submitted.")
            return
        except Exception:
            pass

        print(f"\n  WARNING: Warning dialog not found after {WARNING_WAIT}s.")
        print("  Check the Oracle Projects page — if 'Last Submitted' shows 0, it worked.")
        print("  If not, re-run and watch for the Warning dialog appearing.")
        return

    print("  Warning dialog appeared.")

    # ── Step 9: Click Yes ─────────────────────────────────────────────────────
    print("Clicking 'Yes'...")
    try:
        yes_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", yes_btn)
    print("  Clicked Yes.")

    # ── Done ──────────────────────────────────────────────────────────────────
    time.sleep(2)
    print("\nDone. Cost forecast submitted successfully.")
    print("The 'Last Submitted' counter for all projects has been reset to 0.")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    attach_mode = "--attach" in sys.argv
    driver = get_driver(attach=attach_mode)
    try:
        run(driver, attach=attach_mode)
    except KeyboardInterrupt:
        print("\nCancelled.")
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        raise
    finally:
        print("\nScript complete. Keeping browser open — close it manually when done.")
        time.sleep(30)
        try:
            driver.quit()
        except Exception:
            pass
