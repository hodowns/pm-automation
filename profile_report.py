#!/usr/bin/env python3
"""
Profile Report — builds a Project Contact List from Outlook mail and calendar.

Usage:
    python profile_report.py F51119            # write to spreadsheet
    python profile_report.py F51119 --preview  # print list, do not write
    python profile_report.py F51119 --days 90  # override lookback window
"""

import argparse
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import win32com.client
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ═══════════════════════════════════════════════════════════════════
# USER-EDITABLE SETTINGS  ← update these as you add new PIDs
# ═══════════════════════════════════════════════════════════════════

SHAREPOINT_BASE = Path(
    r"C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS"
)

TEMPLATE_FILE = Path(r"C:\Users\hodowns\Downloads\Project Contact List.xlsx")

CISCO_DOMAINS = {"cisco.com"}

# Add / update customer domains whenever you run a new PID.
# Partner / reseller domains go in the same set as the customer.
CUSTOMER_DOMAINS_BY_PID: dict[str, set] = {
    "F51119": {"houstontx.gov"},
    "F24373": {"uchicagomedicine.org", "meridianitinc.com"},
    "F22532": {"accenture.com"},
    "F29426": {"jaxsheriff.org"},
    "F49302": set(),   # NSO — add domains when known
    "F48852": {"maryland.gov"},
    "C16507": {"state.mn.us"},
    "F43673": {"hennepincounty.gov", "hennepin.us", "cdw.com"},
    "F49303": set(),   # NSO — add domains when known
    "C16507": {"state.mn.us"},
}

# If any of these addresses appear in an email, skip that entire email —
# don't extract any contacts from it. Use for internal Cisco threads
# (finance, ops, etc.) that are not relevant to the customer project.
POISON_ADDRESSES: set = {
    "cost-revenue-forecast@cisco.com",
}

# Specific addresses to always exclude from every report (not whole-thread,
# just the address itself). Add any system accounts or departed employees here.
EXCLUDE_ADDRESSES: set = {
    "cisco_as_oracle_projects@cisco.com",  # internal Cisco Oracle system
    "grerbe@cisco.com",                    # left the company
}

# ═══════════════════════════════════════════════════════════════════
# COLOURS
# ═══════════════════════════════════════════════════════════════════
_GREY   = PatternFill(start_color="FFC0C0C0", end_color="FFC0C0C0", fill_type="solid")
_YELLOW = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")
_ORANGE = PatternFill(start_color="FFFFD966", end_color="FFFFD966", fill_type="solid")

# ═══════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════

_EXCLUDE_RE = re.compile(
    r"noreply|no[-_]reply|donotreply|do[-_]not[-_]reply"
    r"|scheduling.?assistant|meeting[-_.]?room"
    r"|notifications@|alerts@|postmaster@|mailer.daemon"
    r"|^dl[-_.@]|distrib|listserv|bounce@|auto[-_]?reply"
    r"|^messenger@",          # Webex notification address
    re.IGNORECASE,
)

# Cisco contractor marker in Exchange display names: "Firstname Lastname -T (alias - COMPANY at Cisco)"
_CISCO_CONTRACTOR_RE = re.compile(r"\s-T\s+\(", re.IGNORECASE)

# Ordered list of (pattern, role label) for rooms, devices, and service queues.
# Matched against the combined email address + display name.
_RESOURCE_TYPES = [
    (re.compile(r"surface.?hub",          re.I), "Surface Hub"),
    (re.compile(r"webex.?board|webexboard",re.I), "Webex Board"),
    (re.compile(r"conf.?rm\d*|confroom",  re.I), "Conference Room"),
    (re.compile(r"trainingrm|training.?room", re.I), "Training Room"),
    (re.compile(r"networkservices|servicedesk|helpdesk|^snow[-_.]|snow.networkservices", re.I), "ServiceNow Queue"),
]

def detect_resource_type(email: str, name: str) -> str:
    """Return a role label if this address is a room / device / queue, else ''."""
    combined = f"{email} {name}"
    for pattern, label in _RESOURCE_TYPES:
        if pattern.search(combined):
            return label
    return ""

# Pattern to strip City of Houston's "- IT" / "- IT (Contractor)" / "- IT (FTE)" tags
_COH_DEPT_RE = re.compile(
    r"\s*-\s*(IT|IS|HR|FIN|Finance|Legal|Procurement)\s*(\(.*?\))?\s*",
    re.IGNORECASE,
)

_TITLE_RE = re.compile(
    r"(Senior|Principal|Lead|Staff|Associate|Jr\.?\s+|Chief\s+)?"
    r"(Director|Manager|Engineer|Architect|Analyst|Consultant|Specialist"
    r"|Executive|Officer|President|VP|CTO|CIO|CEO|PMO|PM\b|Coordinator"
    r"|Administrator|Developer|Representative|Advisor|Supervisor|Liaison)",
    re.IGNORECASE,
)


def is_excluded(email: str) -> bool:
    return not email or "@" not in email or bool(_EXCLUDE_RE.search(email))


def normalize_name(raw: str, email: str = "") -> str:
    """
    'JOHN DOE' -> 'John Doe'
    'Doe, John' -> 'John Doe'
    'Aida - IT (Contractor) DeHoyos' -> 'Aida DeHoyos'
    If name looks like an email address, build name from email username instead.
    """
    if not raw:
        raw = ""
    raw = raw.strip()

    # If the "name" is just an email address, build from the email username
    if "@" in raw or not raw:
        src = email or raw
        if "@" in src:
            username = src.split("@")[0]
            # "chi.vu" -> "Chi Vu", "diana.garcia2" -> "Diana Garcia"
            parts = re.split(r"[._\-]", username)
            # Strip trailing digits (e.g. "garcia2" -> "garcia")
            parts = [re.sub(r"\d+$", "", p) for p in parts]
            raw = " ".join(p.capitalize() for p in parts if p)
        if not raw:
            return ""

    # Strip CoH directory artifacts: ' ' (apostrophe-space-apostrophe separator)
    raw = raw.replace("' '", " ")

    # Strip department tags like "- IT", "- IT (Contractor)", "- IS (FTE)"
    raw = _COH_DEPT_RE.sub(" ", raw).strip()

    # "Last, First" or "Last, First Middle"
    if "," in raw:
        parts = [p.strip() for p in raw.split(",", 1)]
        if len(parts) == 2 and parts[1]:
            raw = f"{parts[1]} {parts[0]}"

    # Fix ALL-CAPS or all-lowercase
    if raw.isupper() or raw.islower():
        raw = raw.title()

    # Collapse any double-spaces left by the dept-tag strip
    raw = re.sub(r"\s{2,}", " ", raw).strip()
    return raw


def smtp_from_ae(ae) -> tuple:
    """Return (smtp_email, display_name) from an Outlook AddressEntry."""
    try:
        if ae.AddressEntryUserType in (0, 10):  # Exchange user / remote
            user = ae.GetExchangeUser()
            if user and user.PrimarySmtpAddress:
                smtp = user.PrimarySmtpAddress.lower()
                return smtp, normalize_name(ae.Name, smtp)
        addr = ae.Address
        if addr and "@" in addr:
            return addr.lower(), normalize_name(ae.Name, addr.lower())
    except Exception:
        pass
    return "", ""


def smtp_from_sender(item) -> tuple:
    """Return (smtp_email, display_name) for the sender of a MailItem."""
    try:
        ae = item.Sender
        if ae:
            result = smtp_from_ae(ae)
            if result[0]:
                return result
        addr = item.SenderEmailAddress
        if addr and "@" in addr:
            return addr.lower(), normalize_name(item.SenderName, addr.lower())
    except Exception:
        pass
    return "", ""


# ═══════════════════════════════════════════════════════════════════
# FOLDER FINDER
# ═══════════════════════════════════════════════════════════════════

def find_pid_folder(pid: str) -> Path | None:
    pid_up = pid.upper()
    for child in SHAREPOINT_BASE.iterdir():
        if not child.is_dir():
            continue
        for level2 in child.iterdir():
            if not level2.is_dir():
                continue
            # Direct match at level 2 (normal case)
            if pid_up in level2.name.upper():
                return level2
            # One level deeper for portfolio subfolders (e.g. MD DHS Portfolio)
            for level3 in level2.iterdir():
                if level3.is_dir() and pid_up in level3.name.upper():
                    return level3
    return None


# ═══════════════════════════════════════════════════════════════════
# LOOKBACK WINDOW
# ═══════════════════════════════════════════════════════════════════

def get_lookback_days(contact_file: Path, override) -> int:
    if override is not None:
        return override
    if not contact_file.exists():
        return 180
    try:
        wb = load_workbook(contact_file, read_only=True, data_only=True)
        ws = wb.active
        has_contacts = False
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if i == 1:
                continue
            if any(c for c in row):
                cell0 = str(row[0] or "")
                if "Resource Contact" not in cell0 and "UNCLASSIFIED" not in cell0:
                    has_contacts = True
                    break
        wb.close()
        return 30 if has_contacts else 180
    except Exception:
        return 180


# ═══════════════════════════════════════════════════════════════════
# OUTLOOK SEARCH
# ═══════════════════════════════════════════════════════════════════

def search_outlook(pid: str, days: int, own_email: str, ns) -> dict:
    """
    Returns {email: {name, sources, pid_mail_bodies}} for all contacts
    found in mail or calendar items that mention the PID.
    """
    contacts: dict = {}
    cutoff     = datetime.now() - timedelta(days=days)
    cutoff_str = cutoff.strftime("%m/%d/%Y %H:%M")
    pid_up     = pid.upper()

    def add(email: str, name: str, source: str, body: str = ""):
        email = email.lower().strip()
        if not email or is_excluded(email) or email == own_email:
            return
        if email in EXCLUDE_ADDRESSES:
            return
        # Skip Cisco contractors: display name contains "-T (" marker
        if _CISCO_CONTRACTOR_RE.search(name):
            return
        if email not in contacts:
            contacts[email] = {"name": normalize_name(name, email), "sources": set(), "bodies": []}
        else:
            # Only update name if the new one is longer AND wasn't built from email address
            new_name = normalize_name(name, email)
            if len(new_name) > len(contacts[email]["name"]) and "@" not in name:
                contacts[email]["name"] = new_name
        contacts[email]["sources"].add(source)
        if body:
            contacts[email]["bodies"].append(body)

    # ── Mail ──────────────────────────────────────────────────────────
    # Two-pass approach:
    #   Pass 1 (fast, deep): subject contains PID — scans Inbox + subfolders
    #   Pass 2 (slower, top-level only): body contains PID — catches email
    #           threads where the PID appears in the message text but not subject
    print("  Scanning mail folders...", flush=True)
    mail_count  = 0
    seen_ids    = set()   # track EntryIDs to avoid double-counting

    def process_item(item, body_text=""):
        nonlocal mail_count
        try:
            entry_id = item.EntryID
            if entry_id in seen_ids:
                return
            seen_ids.add(entry_id)
            if item.Class != 43:  # olMail
                return
            # Skip entire email if it involves a poison address (internal Cisco threads)
            try:
                for recip in item.Recipients:
                    re_, _ = smtp_from_ae(recip.AddressEntry)
                    if re_ in POISON_ADDRESSES:
                        return
            except Exception:
                pass
            se, _ = smtp_from_sender(item)
            if se and se in POISON_ADDRESSES:
                return
            mail_count += 1
            if not body_text:
                body_text = item.Body or ""
            se, sn = smtp_from_sender(item)
            if se:
                add(se, sn, "mail", body_text)
            for recip in item.Recipients:
                try:
                    re_, rn = smtp_from_ae(recip.AddressEntry)
                    if re_:
                        add(re_, rn, "mail")
                except Exception:
                    pass
        except Exception:
            pass

    def scan_folder_subject(folder, depth=0):
        """Pass 1: subject filter, Inbox + subfolders (depth 1)."""
        if depth > 1:
            return
        try:
            restriction = (
                f"[ReceivedTime] >= '{cutoff_str}' AND "
                f"[Subject] LIKE '%{pid_up}%'"
            )
            for item in folder.Items.Restrict(restriction):
                process_item(item)
        except Exception:
            pass
        try:
            for sub in folder.Folders:
                scan_folder_subject(sub, depth + 1)
        except Exception:
            pass

    def scan_folder_body(folder):
        """Pass 2: full body check, top-level folder only (no recursion)."""
        try:
            restriction = f"[ReceivedTime] >= '{cutoff_str}'"
            for item in folder.Items.Restrict(restriction):
                try:
                    if item.Class != 43:
                        continue
                    body_text = item.Body or ""
                    if pid_up not in body_text.upper():
                        continue
                    process_item(item, body_text)
                except Exception:
                    continue
        except Exception:
            pass

    try:
        inbox = ns.GetDefaultFolder(6)   # Inbox
        sent  = ns.GetDefaultFolder(5)   # Sent Items
        # Pass 1: subject search, Inbox + subfolders + Sent
        scan_folder_subject(inbox)
        scan_folder_subject(sent)
        # Pass 2: body search, top-level only
        print("  Scanning email bodies (top-level folders)...", flush=True)
        scan_folder_body(inbox)
        scan_folder_body(sent)
    except Exception as e:
        print(f"  Warning: mail scan error - {e}", flush=True)

    print(f"  Matched emails: {mail_count}", flush=True)

    # ── Calendar ──────────────────────────────────────────────────────
    print("  Scanning calendar...", flush=True)
    cal_count = 0
    try:
        cal   = ns.GetDefaultFolder(9)  # olFolderCalendar
        items = cal.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        now_str    = datetime.now().strftime("%m/%d/%Y %H:%M")
        restricted = items.Restrict(
            f"[Start] >= '{cutoff_str}' AND [Start] <= '{now_str}'"
        )
        for item in restricted:
            try:
                subj = (item.Subject or "").upper()
                if pid_up not in subj:
                    continue
                cal_count += 1
                for recip in item.Recipients:
                    try:
                        re_, rn = smtp_from_ae(recip.AddressEntry)
                        if re_:
                            add(re_, rn, "calendar")
                    except Exception:
                        pass
                # Organiser may not be in Recipients list
                try:
                    org = item.Organizer or ""
                    if org:
                        r2 = ns.CreateRecipient(org)
                        r2.Resolve()
                        if r2.Resolved:
                            oe, on = smtp_from_ae(r2.AddressEntry)
                            if oe:
                                add(oe, on, "calendar")
                except Exception:
                    pass
            except Exception:
                continue
    except Exception as e:
        print(f"  Warning: calendar scan error - {e}", flush=True)

    print(f"  Matched calendar items: {cal_count}", flush=True)
    return contacts


# ═══════════════════════════════════════════════════════════════════
# ROLE LOOKUPS
# ═══════════════════════════════════════════════════════════════════

def lookup_cisco_role(email: str, ns) -> str:
    """Look up job title from the Outlook Global Address List."""
    try:
        r = ns.CreateRecipient(email)
        r.Resolve()
        if r.Resolved:
            ae = r.AddressEntry
            if ae and ae.AddressEntryUserType in (0, 10):
                user = ae.GetExchangeUser()
                if user and user.JobTitle:
                    return user.JobTitle
    except Exception:
        pass
    return ""


def scrape_signature_role(bodies: list) -> str:
    """
    Scan email bodies for a line that looks like a job title.
    Returns the first match found, or empty string.
    """
    for body in bodies:
        lines = body.splitlines()
        for line in lines:
            line = line.strip()
            if not line or len(line) > 70:
                continue
            # Skip lines with URLs, phone numbers, email addresses
            if re.search(r"@|http|www|\d{3}[-.\s]\d{3}|\+1|\(?\d{3}\)?", line):
                continue
            if _TITLE_RE.search(line) and len(line.split()) <= 8:
                return line
    return ""


# ═══════════════════════════════════════════════════════════════════
# CLASSIFY
# ═══════════════════════════════════════════════════════════════════

def classify(contacts: dict, customer_domains: set) -> tuple:
    customers, cisco_ppl, unclassified = [], [], []
    for email, info in contacts.items():
        domain = email.split("@")[1] if "@" in email else ""
        entry  = {**info, "email": email, "domain": domain}
        if domain in CISCO_DOMAINS:
            cisco_ppl.append(entry)
        elif domain in customer_domains:
            customers.append(entry)
        else:
            unclassified.append(entry)
    key = lambda x: x["name"].lower()
    return sorted(customers, key=key), sorted(cisco_ppl, key=key), sorted(unclassified, key=key)


# ═══════════════════════════════════════════════════════════════════
# PREVIEW
# ═══════════════════════════════════════════════════════════════════

def print_preview(customers: list, cisco_ppl: list, unclassified: list, days: int, pid: str):
    W = 72
    print()
    print("=" * W)
    print(f"  PREVIEW  ·  Profile report {pid}  ·  lookback: {days} days")
    print("=" * W)

    def block(label, people):
        print(f"\n  {label}  ({len(people)} contacts)")
        print("  " + "─" * (W - 2))
        if not people:
            print("    (none found)")
            return
        print(f"  {'Email':<40} {'Name':<26} {'Role':<22} Source")
        print("  " + "─" * (W - 2))
        for p in people:
            src  = "+".join(sorted(p["sources"]))
            role = p.get("role", "") or ""
            uv   = " ⚠" if p.get("role_unverified") else ""
            print(f"  {p['email']:<40} {p['name']:<26} {role+uv:<22} [{src}]")

    block("CUSTOMER CONTACTS", customers)
    block("CISCO CONTACTS",    cisco_ppl)

    if unclassified:
        block("⚠  UNCLASSIFIED — please review", unclassified)
        unknown = sorted({p["domain"] for p in unclassified if p["domain"]})
        if unknown:
            print(f"\n  Unknown domains spotted: {', '.join(unknown)}")
            print("  Let me know if any belong to the customer or a partner/reseller")
            print("  and I'll add them to the domain list before you save.")
    print()


# ═══════════════════════════════════════════════════════════════════
# EXCEL WRITER
# ═══════════════════════════════════════════════════════════════════

def _read_block(ws, start_row: int, end_row: int) -> dict:
    """Return {email_lower: row_index} for data rows in [start_row+1 .. end_row)."""
    result = {}
    for r in range(start_row + 1, end_row):
        val = ws.cell(row=r, column=1).value
        if val and "@" in str(val):
            result[str(val).strip().lower()] = r
    return result


def _fill_block(ws, new_people: list, start_row: int, end_row: int) -> int:
    """
    Write new_people into [start_row+1 .. end_row).
    - Existing contacts: fill blank cells only, reactivate if greyed out.
    - Absent contacts: grey out + mark (inactive).
    - New contacts: append, inserting rows if the block is full.
    Returns the (possibly grown) end_row.
    """
    existing   = _read_block(ws, start_row, end_row)
    new_emails = {p["email"] for p in new_people}

    # Grey out anyone who didn't show up this run
    for email, r in existing.items():
        if email not in new_emails:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = _GREY
            nc = ws.cell(row=r, column=2)
            if nc.value and "(inactive)" not in str(nc.value):
                nc.value = f"{nc.value} (inactive)"

    write_row = start_row + 1
    for person in new_people:
        email = person["email"]
        role  = person.get("role", "") or ""
        unv   = person.get("role_unverified", False)

        if email in existing:
            r = existing[email]
            # Reactivate
            for c in range(1, 8):
                if ws.cell(row=r, column=c).fill.fgColor.rgb == "FFC0C0C0":
                    ws.cell(row=r, column=c).fill = PatternFill(fill_type=None)
            nc = ws.cell(row=r, column=2)
            if nc.value:
                nc.value = str(nc.value).replace(" (inactive)", "")
            # Fill blanks only
            if not ws.cell(row=r, column=1).value:
                ws.cell(row=r, column=1).value = email
            if not ws.cell(row=r, column=2).value:
                ws.cell(row=r, column=2).value = person["name"]
            if not ws.cell(row=r, column=3).value and role:
                ws.cell(row=r, column=3).value = role
                if unv:
                    ws.cell(row=r, column=3).fill = _YELLOW
        else:
            # Advance write_row past any occupied rows
            while write_row < end_row and ws.cell(row=write_row, column=1).value:
                write_row += 1
            # Insert if block is full
            if write_row >= end_row:
                ws.insert_rows(end_row)
                end_row += 1
            ws.cell(row=write_row, column=1).value = email
            ws.cell(row=write_row, column=2).value = person["name"]
            if role:
                ws.cell(row=write_row, column=3).value = role
                if unv:
                    ws.cell(row=write_row, column=3).fill = _YELLOW
            write_row += 1

    return end_row


def write_excel(contact_file: Path, customers: list, cisco_ppl: list,
                unclassified: list, pid: str):
    if contact_file.exists():
        wb = load_workbook(contact_file)
    else:
        wb = load_workbook(TEMPLATE_FILE)

    ws = wb.active

    # ── Locate header rows ────────────────────────────────────────────
    cust_hdr = cisco_hdr = None
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        v = str(row[0] or "")
        if "Customer Resource Contact" in v and cust_hdr  is None:
            cust_hdr  = i
        if "Cisco Resource Contact"    in v and cisco_hdr is None:
            cisco_hdr = i

    cust_hdr  = cust_hdr  or 1
    cisco_hdr = cisco_hdr or 24

    # ── Customer block ────────────────────────────────────────────────
    _fill_block(ws, customers, cust_hdr, cisco_hdr)

    # Re-find Cisco header (may have shifted if rows were inserted above)
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if row[0] and "Cisco Resource Contact" in str(row[0]):
            cisco_hdr = i
            break

    # ── Cisco block ───────────────────────────────────────────────────
    unclass_hdr = None
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if row[0] and "UNCLASSIFIED" in str(row[0]):
            unclass_hdr = i
            break
    cisco_end = unclass_hdr if unclass_hdr else ws.max_row + 1

    _fill_block(ws, cisco_ppl, cisco_hdr, cisco_end)

    # ── UNCLASSIFIED block ────────────────────────────────────────────
    if unclassified:
        # Re-find or create header
        unclass_hdr = None
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if row[0] and "UNCLASSIFIED" in str(row[0]):
                unclass_hdr = i
                break
        if unclass_hdr is None:
            ws.append(["UNCLASSIFIED - PLEASE REVIEW", "Name", "Role",
                        "Project/Inititaive 1", "Project/Inititaive 2",
                        "Project/Inititaive 3", "Project/Inititaive 4"])
            unclass_hdr = ws.max_row
            for c in range(1, 8):
                ws.cell(row=unclass_hdr, column=c).fill = _ORANGE
                ws.cell(row=unclass_hdr, column=c).font = Font(bold=True)

        _fill_block(ws, unclassified, unclass_hdr, ws.max_row + 1)

    contact_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(contact_file)
    return contact_file


# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════

def main():
    # Ensure UTF-8 output on Windows terminals
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="Profile Report - build a Project Contact List from Outlook."
    )
    parser.add_argument("pid",       help="Project ID, e.g. F51119")
    parser.add_argument("--preview", action="store_true",
                        help="Print contact list without writing the file")
    parser.add_argument("--days",    type=int, default=None,
                        help="Override lookback window in days")
    args = parser.parse_args()

    pid = args.pid.upper()
    print(f"\n-- Profile report {pid} " + "-" * (50 - len(pid)), flush=True)

    # 1. Locate project folder
    print("\n[1/5] Locating project folder...", flush=True)
    pid_folder = find_pid_folder(pid)
    if not pid_folder:
        sys.exit(f"  ERROR: No folder found for '{pid}' under:\n    {SHAREPOINT_BASE}")
    print(f"  OK: {pid_folder}", flush=True)

    contact_file = pid_folder / "Customer Profile" / "Project Contact List.xlsx"
    print(f"  Contact file: {contact_file}", flush=True)

    # 2. Lookback window
    days = get_lookback_days(contact_file, args.days)
    reason = "first run - using deep scan" if days == 180 else "re-run"
    print(f"\n[2/5] Lookback: {days} days  ({reason})", flush=True)

    # 3. Connect to Outlook
    print("\n[3/5] Connecting to Outlook...", flush=True)
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns      = outlook.GetNamespace("MAPI")
        own_ae  = ns.CurrentUser
        r       = ns.CreateRecipient(own_ae.Address)
        r.Resolve()
        own_email = smtp_from_ae(r.AddressEntry)[0] if r.Resolved else own_ae.Address.lower()
        print(f"  Logged in as: {own_email}", flush=True)
    except Exception as e:
        sys.exit(f"  ERROR: Could not connect to Outlook. Is it open?\n  {e}")

    if pid not in CUSTOMER_DOMAINS_BY_PID or not CUSTOMER_DOMAINS_BY_PID[pid]:
        print(f"\n  WARNING: no customer domains defined for {pid}.", flush=True)
        print("     All non-Cisco contacts will appear in UNCLASSIFIED.", flush=True)
        print("     Edit CUSTOMER_DOMAINS_BY_PID in this script to add them.", flush=True)

    raw = search_outlook(pid, days, own_email, ns)
    print(f"  Total unique contacts collected: {len(raw)}", flush=True)

    # 4. Look up roles
    print("\n[4/5] Looking up roles...", flush=True)
    customer_domains = CUSTOMER_DOMAINS_BY_PID.get(pid, set())
    role_count = 0

    for email, info in raw.items():
        domain = email.split("@")[1] if "@" in email else ""
        name   = info.get("name", "")

        # Check for room / device / queue first (applies to any domain)
        resource = detect_resource_type(email, name)
        if resource:
            info["role"]            = resource
            info["role_unverified"] = False
            role_count += 1
        elif domain in CISCO_DOMAINS:
            role = lookup_cisco_role(email, ns)
            info["role"]            = role
            info["role_unverified"] = False
            if role:
                role_count += 1
        elif domain in customer_domains:
            role = scrape_signature_role(info.get("bodies", []))
            info["role"]            = f"{role} (unverified)" if role else ""
            info["role_unverified"] = bool(role)
            if role:
                role_count += 1
        else:
            info["role"]            = ""
            info["role_unverified"] = False

    print(f"  Roles found: {role_count}", flush=True)

    # 5. Classify
    print("\n[5/5] Classifying contacts...", flush=True)
    customers, cisco_ppl, unclassified = classify(raw, customer_domains)
    print(f"  Customer: {len(customers)}  |  Cisco: {len(cisco_ppl)}  |  Unclassified: {len(unclassified)}", flush=True)

    if unclassified:
        domains = sorted({p["domain"] for p in unclassified if p["domain"]})
        print(f"\n  Unknown domains: {', '.join(domains)}", flush=True)
        print("  These will appear in UNCLASSIFIED. Tell me which belong to the customer", flush=True)
        print("  or a partner and I will re-classify them.", flush=True)

    # Output
    if args.preview:
        print_preview(customers, cisco_ppl, unclassified, days, pid)
        print("  Preview only - nothing was written.", flush=True)
        print(f"  To save, run:  python profile_report.py {pid}", flush=True)
    else:
        print("\n  Writing spreadsheet...", flush=True)
        saved = write_excel(contact_file, customers, cisco_ppl, unclassified, pid)
        print(f"\n  Saved: {saved}", flush=True)
        print("  Open the file to review. Run with --preview anytime to check before saving.", flush=True)


if __name__ == "__main__":
    main()
