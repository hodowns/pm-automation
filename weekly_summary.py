#!/usr/bin/env python3
"""
Weekly Summary — generates a visual HTML meeting summary report for a PID.

Usage:
    python weekly_summary.py F51119
"""

import argparse
import re
import sys
import html as html_module
from datetime import datetime, timedelta
from pathlib import Path
from dataclasses import dataclass, field

# ═══════════════════════════════════════════════════════════════════
# SETTINGS
# ═══════════════════════════════════════════════════════════════════

SHAREPOINT_BASE = Path(
    r"C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS"
)

OUTPUT_BASE = Path(r"C:\Users\hodowns\OneDrive - Cisco\Desktop\Meeting Summaries")

WEBEX_TOKENS_FILE = Path(r"C:\Users\hodowns\webex-mcp-oauth\tokens.json")
WEBEX_BASE = "https://webexapis.com/v1"


def _load_webex_token() -> str:
    import json as _json
    try:
        return _json.loads(WEBEX_TOKENS_FILE.read_text())["access_token"]
    except Exception:
        return ""

LOOKBACK_DAYS = 7

CUSTOMER_DOMAINS_BY_PID: dict = {
    "F51119": {"houstontx.gov"},
    "F24373": {"uchicagomedicine.org", "meridianitinc.com"},
    "F49302": set(),
    "F49303": set(),
}

CUSTOMER_NAMES_BY_PID: dict = {
    "F51119": "City of Houston",
    "F24373": "UChicago Medicine",
}

POISON_ADDRESSES: set = {"cost-revenue-forecast@cisco.com"}

# ═══════════════════════════════════════════════════════════════════
# WORKSTREAM DETECTION
# ═══════════════════════════════════════════════════════════════════

WORKSTREAM_PATTERNS = [
    (re.compile(r"\bACI\b",                              re.I), "ACI"),
    (re.compile(r"SD.?WAN|SDWAN",                        re.I), "SD-WAN"),
    (re.compile(r"Catalyst.?Center|Cat.?Center|DNAC",    re.I), "Catalyst Center"),
    (re.compile(r"Umbrella|SASE|Architect.as.a.Service", re.I), "Umbrella / SASE"),
    (re.compile(r"\bICE\b|Identity.*Engines",            re.I), "ICE / Identity"),
    (re.compile(r"Meraki|Wireless|Wi.?Fi",               re.I), "Wireless / Meraki"),
    (re.compile(r"Terraform|Ansible|as.code|Automation", re.I), "Automation"),
    (re.compile(r"Status|Weekly|LCS",                    re.I), "General / Status"),
]


def detect_workstream(title: str) -> str:
    for pattern, label in WORKSTREAM_PATTERNS:
        if pattern.search(title):
            return label
    return "General"


# ═══════════════════════════════════════════════════════════════════
# DATA STRUCTURES
# ═══════════════════════════════════════════════════════════════════

@dataclass
class ActionItem:
    number: int
    action: str
    owner: str
    due_raw: str
    due_date: object   # datetime or None
    due_auto: bool
    source: str
    workstream: str


@dataclass
class MeetingNote:
    date: object       # datetime
    title: str
    workstream: str
    attendees: list = field(default_factory=list)
    summary: str = ""
    key_decisions: list = field(default_factory=list)
    action_items: list = field(default_factory=list)
    next_steps: list = field(default_factory=list)
    recording_url: str = ""


@dataclass
class EmailThread:
    date: object       # datetime
    subject: str
    sender: str
    snippet: str
    workstream: str


# ═══════════════════════════════════════════════════════════════════
# FOLDER FINDER  (same logic as profile_report.py)
# ═══════════════════════════════════════════════════════════════════

def find_pid_folder(pid: str):
    pid_up = pid.upper()
    for state_dir in SHAREPOINT_BASE.iterdir():
        if not state_dir.is_dir():
            continue
        for project_dir in state_dir.iterdir():
            if project_dir.is_dir() and pid_up in project_dir.name.upper():
                return project_dir
    return None


# ═══════════════════════════════════════════════════════════════════
# ROSTER READER
# ═══════════════════════════════════════════════════════════════════

def read_roster(contact_file: Path) -> list:
    if not contact_file.exists():
        return []
    try:
        from openpyxl import load_workbook
        wb = load_workbook(contact_file, read_only=True, data_only=True)
        ws = wb.active
        roster = []
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if i == 1:
                continue
            email = str(row[0] or "").strip().lower()
            name  = str(row[1] or "").strip()
            role  = str(row[2] or "").strip()
            if "@" in email and "Resource Contact" not in email:
                roster.append({
                    "email": email,
                    "name": name,
                    "role": role,
                    "domain": email.split("@")[1],
                })
        wb.close()
        return roster
    except Exception as e:
        print(f"  Warning: could not read roster — {e}", flush=True)
        return []


# ═══════════════════════════════════════════════════════════════════
# DUE DATE PARSING
# ═══════════════════════════════════════════════════════════════════

def _try_parse_date(s: str, year: int):
    s = re.sub(r"\s+", " ", s.strip())
    # Formats that include a year — parse directly
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%b %d %Y", "%B %d %Y", "%b %d, %Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    # Formats without a year — inject year first to avoid ambiguity warnings
    for fmt in ("%b %d", "%B %d", "%b. %d"):
        try:
            return datetime.strptime(f"{s} {year}", fmt + " %Y")
        except ValueError:
            continue
    return None


def parse_due_date(due_raw: str, meeting_date: datetime):
    """Return (datetime_or_None, is_auto_bool)."""
    if not due_raw:
        return meeting_date + timedelta(days=7), True

    raw = due_raw.strip()
    up  = raw.upper()

    if re.search(r"\bIMMEDIATELY\b|\bASAP\b|\bURGENT\b", up):
        return meeting_date + timedelta(days=3), True

    if re.search(r"\bTBD\b|\bN/A\b|\bTBC\b|\bNEXT SESSION\b|\bONGOING\b", up):
        return meeting_date + timedelta(days=7), True

    # "~2–3 weeks" or "~2-3 weeks"
    m = re.search(r"(\d+)\s*[–\-]\s*(\d+)\s*weeks?", raw, re.I)
    if m:
        weeks = (int(m.group(1)) + int(m.group(2))) // 2
        return meeting_date + timedelta(weeks=weeks), True
    m = re.search(r"(\d+)\s*weeks?", raw, re.I)
    if m:
        return meeting_date + timedelta(weeks=int(m.group(1))), True

    # "~X days"
    m = re.search(r"(\d+)\s*days?", raw, re.I)
    if m:
        return meeting_date + timedelta(days=int(m.group(1))), True

    # "Before Aug 17" / "Before migration" etc.
    m = re.search(r"before\s+(.+)", raw, re.I)
    if m:
        d = _try_parse_date(m.group(1).strip(), meeting_date.year)
        if d:
            return d, False
        return meeting_date + timedelta(days=7), True

    # "EOD Aug 14"
    m = re.search(r"EOD\s+(.+)", raw, re.I)
    if m:
        d = _try_parse_date(m.group(1).strip(), meeting_date.year)
        if d:
            return d, False

    # Try direct date parse
    d = _try_parse_date(raw, meeting_date.year)
    if d:
        return d, False

    return meeting_date + timedelta(days=7), True


# ═══════════════════════════════════════════════════════════════════
# MEETING NOTES PARSER  (.md files)
# ═══════════════════════════════════════════════════════════════════

def _extract_section(text: str, heading: str) -> str:
    pattern = re.escape(heading)
    m = re.search(pattern + r"\s*\n(.*?)(?=\n##\s|\Z)", text, re.S)
    return m.group(1).strip() if m else ""


def _short_title(title: str) -> str:
    title = re.sub(r"\[?EXTERNAL\]?|\[?External\]?", "", title, flags=re.I)
    title = re.sub(r"F\d{5}", "", title)
    title = re.sub(r"City of Houston", "CoH", title, flags=re.I)
    title = re.sub(r"\s+", " ", title).strip(" -–[]:")
    return title[:65]


def parse_meeting_notes(md_file: Path, cutoff: datetime):
    try:
        text = md_file.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return None

    # Date from frontmatter or filename
    date_m = re.search(r"^date:\s*(\d{4}-\d{2}-\d{2})", text, re.M)
    if date_m:
        try:
            meeting_date = datetime.strptime(date_m.group(1), "%Y-%m-%d")
        except ValueError:
            meeting_date = None
    else:
        fname_m = re.match(r"(\d{4}-\d{2}-\d{2})", md_file.stem)
        meeting_date = datetime.strptime(fname_m.group(1), "%Y-%m-%d") if fname_m else None

    if not meeting_date or meeting_date < cutoff:
        return None

    title_m = re.search(r"^#\s+(.+)$", text, re.M)
    title   = title_m.group(1).strip() if title_m else md_file.stem
    workstream = detect_workstream(title)

    # Attendees
    att_m = re.search(r"\*\*Attendees:\*\*\s*(.+)", text)
    attendees = [a.strip() for a in att_m.group(1).split(",")] if att_m else []

    # Recording
    rec_m = re.search(r"https://cisco\.webex\.com[^\s|\"<]+", text)
    recording_url = rec_m.group(0) if rec_m else ""

    # Summary
    summary = _extract_section(text, "## Summary")

    # Key Decisions
    dec_text = _extract_section(text, "## Key Decisions")
    decisions = []
    for line in dec_text.splitlines():
        line = line.strip()
        if line.startswith("- ") or line.startswith("* "):
            item = re.sub(r"^\*\*[^*]+\*\*:\s*", "", line[2:].strip())
            if item:
                decisions.append(item)

    # Action Items table
    ai_section = _extract_section(text, "## Action Items")
    action_items = []
    for line in ai_section.splitlines():
        parts = [p.strip() for p in line.split("|") if p.strip()]
        if len(parts) >= 4 and parts[0].isdigit():
            try:
                due_raw  = parts[3] if len(parts) > 3 else ""
                due_date, due_auto = parse_due_date(due_raw, meeting_date)
                action_items.append(ActionItem(
                    number   = int(parts[0]),
                    action   = parts[1],
                    owner    = parts[2],
                    due_raw  = due_raw,
                    due_date = due_date,
                    due_auto = due_auto,
                    source   = f"{_short_title(title)} | {meeting_date.strftime('%Y-%m-%d')}",
                    workstream = workstream,
                ))
            except (ValueError, IndexError):
                continue

    # Next Steps
    ns_text = _extract_section(text, "## Next Steps")
    next_steps = []
    for line in ns_text.splitlines():
        line = line.strip()
        if line.startswith("- ") or line.startswith("* "):
            ns = line[2:].strip()
            if ns:
                next_steps.append(ns)

    return MeetingNote(
        date          = meeting_date,
        title         = title,
        workstream    = workstream,
        attendees     = attendees,
        summary       = summary.strip(),
        key_decisions = decisions,
        action_items  = action_items,
        next_steps    = next_steps,
        recording_url = recording_url,
    )


# ═══════════════════════════════════════════════════════════════════
# WEBEX API
# ═══════════════════════════════════════════════════════════════════

def _webex_get(path: str, params: dict = None):
    try:
        import requests
        headers = {
            "Authorization": f"Bearer {_load_webex_token()}",
            "Content-Type": "application/json",
        }
        r = requests.get(
            f"{WEBEX_BASE}/{path.lstrip('/')}",
            headers=headers,
            params=params or {},
            timeout=15,
        )
        if r.status_code == 200:
            return r.json()
        if r.status_code == 401:
            print("  Warning: Webex token returned 401 (expired?). Skipping Webex.", flush=True)
            return None
        print(f"  Warning: Webex {path} → {r.status_code}", flush=True)
        return None
    except Exception as e:
        print(f"  Warning: Webex error — {e}", flush=True)
        return None


def pull_webex_data(pid: str, customer_name: str, since: datetime) -> list:
    me = _webex_get("people/me")
    if not me:
        return []
    print(f"  Webex: {me.get('displayName','?')} ({(me.get('emails') or ['?'])[0]})", flush=True)

    rooms_data = _webex_get("rooms", {"type": "group", "max": 200, "sortBy": "lastactivity"})
    if not rooms_data:
        return []

    pid_up   = pid.upper()
    cust_kws = [pid_up.lower(), customer_name.lower().replace("city of ", ""), "houston", "lcs"]

    matching = [
        r for r in rooms_data.get("items", [])
        if any(kw in (r.get("title") or "").lower() for kw in cust_kws)
    ]
    print(f"  Webex: {len(matching)} matching spaces of {len(rooms_data.get('items',[]))}", flush=True)

    since_str = since.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    results   = []

    for room in matching[:10]:
        msgs_data = _webex_get("messages", {
            "roomId": room["id"],
            "after":  since_str,
            "max":    100,
        })
        if not msgs_data:
            continue
        messages = msgs_data.get("items", [])
        if not messages:
            continue

        snippets = []
        for msg in messages:
            created = msg.get("created", "")[:10]
            sender  = msg.get("personEmail", "").split("@")[0]
            text    = (msg.get("text") or "").strip()[:200]
            if text:
                snippets.append(f"[{created}] {sender}: {text}")

        results.append({
            "space_name":    room.get("title", "Unknown"),
            "message_count": len(messages),
            "snippets":      snippets[:15],
            "workstream":    detect_workstream(room.get("title", "")),
        })

    return results


# ═══════════════════════════════════════════════════════════════════
# OUTLOOK EMAILS
# ═══════════════════════════════════════════════════════════════════

def pull_outlook_emails(pid: str, roster: list, customer_domains: set, since: datetime) -> list:
    try:
        import win32com.client
    except ImportError:
        print("  Warning: win32com not available, skipping Outlook.", flush=True)
        return []
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns      = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"  Warning: Outlook connect failed — {e}", flush=True)
        return []

    pid_up      = pid.upper()
    cutoff_str  = since.strftime("%m/%d/%Y %H:%M")
    cust_domain_list = list(customer_domains)
    emails      = []
    seen_ids    = set()

    def try_extract(item):
        try:
            eid = item.EntryID
            if eid in seen_ids or item.Class != 43:
                return
            seen_ids.add(eid)

            subj = item.Subject or ""
            body = (item.Body or "")[:600]
            subj_up = subj.upper()
            body_up = body.upper()

            pid_match    = pid_up in subj_up or pid_up in body_up
            domain_match = any(d.upper() in body_up for d in cust_domain_list)
            if not pid_match and not domain_match:
                return

            recv = item.ReceivedTime
            try:
                date_obj = datetime(recv.year, recv.month, recv.day, recv.hour, recv.minute)
            except Exception:
                date_obj = since

            snippet = re.sub(r"\s+", " ", body[:200]).strip()
            emails.append(EmailThread(
                date       = date_obj,
                subject    = subj[:120],
                sender     = item.SenderName or item.SenderEmailAddress or "",
                snippet    = snippet,
                workstream = detect_workstream(subj),
            ))
        except Exception:
            pass

    try:
        for folder_id in (6, 5):   # Inbox, Sent
            folder = ns.GetDefaultFolder(folder_id)
            restriction = f"[ReceivedTime] >= '{cutoff_str}'"
            for item in folder.Items.Restrict(restriction):
                try_extract(item)
    except Exception as e:
        print(f"  Warning: Outlook scan error — {e}", flush=True)

    emails.sort(key=lambda e: e.date)
    print(f"  Outlook: {len(emails)} matching emails", flush=True)
    return emails


# ═══════════════════════════════════════════════════════════════════
# STATUS CHIPS
# ═══════════════════════════════════════════════════════════════════

STATUS_COLORS = {
    "On Track": ("#15803d", "#dcfce7"),
    "At Risk":  ("#b45309", "#fef3c7"),
    "Blocked":  ("#b91c1c", "#fee2e2"),
}


def get_status(workstream: str, action_items: list, today: datetime) -> str:
    ws_items = [ai for ai in action_items if ai.workstream == workstream]
    for ai in ws_items:
        if "blocked" in (ai.action + " " + ai.due_raw).lower():
            return "Blocked"
    for ai in ws_items:
        if ai.due_date and ai.due_date < today - timedelta(days=3):
            return "At Risk"
    return "On Track"


# ═══════════════════════════════════════════════════════════════════
# PLAIN TEXT BUILDERS
# ═══════════════════════════════════════════════════════════════════

def plain_accomplished(notes: list) -> str:
    lines = ["ACCOMPLISHED THIS WEEK", "=" * 60, ""]
    for note in sorted(notes, key=lambda n: n.date):
        lines.append(f"▸ {note.title} ({note.date.strftime('%b %d')})")
        if note.summary:
            lines.append(f"  {note.summary[:350]}")
        for d in note.key_decisions[:3]:
            lines.append(f"  • {d}")
        lines.append("")
    return "\n".join(lines)


def plain_action_items(items: list) -> str:
    lines = ["ACTION ITEMS", "=" * 60,
             "* = due date auto-set (7 days from meeting); explicitly stated dates have no flag", ""]
    lines.append(f"{'Action':<50}  {'Owner':<25}  {'Due':<14}  Source")
    lines.append("-" * 110)
    for ai in sorted(items, key=lambda x: (x.due_date or datetime.max)):
        flag    = "*" if ai.due_auto else " "
        due_str = (ai.due_date.strftime("%Y-%m-%d") + flag) if ai.due_date else "TBD"
        action  = ai.action[:50]
        owner   = ai.owner[:25]
        lines.append(f"{action:<50}  {owner:<25}  {due_str:<14}  {ai.source}")
    return "\n".join(lines)


def plain_agenda(workstream: str, items: list, next_steps: list) -> str:
    lines = [f"DRAFT AGENDA — {workstream}", "=" * 60, ""]
    if next_steps:
        lines.append("Carry-over / Next Steps:")
        for ns in next_steps:
            lines.append(f"  • {ns}")
        lines.append("")
    if items:
        lines.append("Open Action Items:")
        for ai in sorted(items, key=lambda x: (x.due_date or datetime.max)):
            flag    = "*" if ai.due_auto else ""
            due_str = (ai.due_date.strftime("%m/%d") + flag) if ai.due_date else "TBD"
            lines.append(f"  [{due_str}] {ai.action}  —  {ai.owner}")
        lines.append("")
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════
# HTML GENERATOR
# ═══════════════════════════════════════════════════════════════════

def _e(s) -> str:
    return html_module.escape(str(s))


_CSS = """
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
  background: #f1f5f9; color: #1e293b; font-size: 14px; line-height: 1.55;
}

/* ── Header ── */
.report-header {
  background: linear-gradient(135deg, #1e3a5f 0%, #0f2940 100%);
  color: white; padding: 22px 32px;
  display: flex; align-items: center; gap: 16px; flex-wrap: wrap;
}
.pid-badge {
  background: #f59e0b; color: #1e293b; font-weight: 800; font-size: 17px;
  padding: 5px 16px; border-radius: 20px; letter-spacing: 1px; flex-shrink: 0;
}
.header-text h1 { font-size: 20px; font-weight: 600; }
.header-text .sub { font-size: 12px; opacity: 0.72; margin-top: 3px; }

/* ── KPI Row ── */
.kpi-row {
  display: flex; gap: 12px; padding: 20px 32px 8px; flex-wrap: wrap;
}
.kpi-tile {
  background: white; border: 1px solid #e2e8f0; border-radius: 10px;
  padding: 14px 20px; flex: 1; min-width: 120px; text-align: center;
  box-shadow: 0 1px 3px rgba(0,0,0,.05);
}
.kpi-num { font-size: 30px; font-weight: 800; color: #1e3a5f; }
.kpi-label { font-size: 11px; color: #64748b; margin-top: 2px; text-transform: uppercase; letter-spacing: .5px; }

/* ── Workstream chips ── */
.ws-row { padding: 12px 32px 20px; display: flex; flex-wrap: wrap; gap: 8px; }
.chip {
  display: inline-flex; align-items: center; padding: 5px 13px;
  border-radius: 20px; font-size: 12px; font-weight: 600;
}
.chip-sm {
  display: inline-block; padding: 2px 8px; border-radius: 12px;
  font-size: 11px; font-weight: 600;
}

/* ── Section card ── */
.section {
  background: white; margin: 0 32px 20px; border-radius: 10px;
  border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,.05);
  overflow: hidden;
}
.section-head {
  background: #1e3a5f; color: white; padding: 11px 20px;
  display: flex; align-items: center; justify-content: space-between;
}
.section-head h2 { font-size: 14px; font-weight: 600; }
.section-body { padding: 16px 20px; }
.section-foot { padding: 4px 20px 14px; }

/* ── Copy button ── */
.copy-btn {
  background: rgba(255,255,255,.15); color: white;
  border: 1px solid rgba(255,255,255,.3); padding: 4px 12px;
  border-radius: 6px; cursor: pointer; font-size: 12px;
  transition: background .15s;
}
.copy-btn:hover { background: rgba(255,255,255,.28); }

/* ── Meeting cards ── */
.meeting-card {
  border: 1px solid #e2e8f0; border-radius: 8px;
  padding: 13px 15px; margin-bottom: 12px;
}
.meeting-head {
  display: flex; align-items: center; gap: 9px; flex-wrap: wrap; margin-bottom: 7px;
}
.m-date {
  background: #1e3a5f; color: white; font-size: 11px; font-weight: 700;
  padding: 2px 8px; border-radius: 4px; flex-shrink: 0;
}
.m-title { font-weight: 600; font-size: 14px; flex: 1; }
.ws-tag {
  background: #eff6ff; color: #1d4ed8; font-size: 11px;
  padding: 2px 8px; border-radius: 10px; font-weight: 500;
}
.ws-tag-sm {
  background: #eff6ff; color: #1d4ed8; font-size: 11px;
  padding: 1px 6px; border-radius: 8px; font-weight: 500;
}
.rec-link { font-size: 11px; color: #0ea5e9; text-decoration: none; }
.rec-link:hover { text-decoration: underline; }
.attendees-line { font-size: 11px; color: #64748b; margin-bottom: 6px; }
.summary-p { font-size: 13px; color: #374151; margin-bottom: 6px; }
.decisions-list { padding-left: 18px; }
.decisions-list li { font-size: 12px; color: #374151; margin-bottom: 3px; }

/* ── Action items table ── */
.ai-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.ai-table th {
  background: #f1f5f9; color: #475569; font-weight: 600; font-size: 12px;
  padding: 8px 11px; text-align: left; border-bottom: 2px solid #e2e8f0;
}
.ai-table td { padding: 8px 11px; border-bottom: 1px solid #f1f5f9; vertical-align: top; }
.ai-table tr:hover td { background: #f8fafc; }
.ai-table tr.overdue td { background: #fff7ed; }
.ai-action { max-width: 300px; }
.ai-owner { font-weight: 500; color: #374151; }
.ai-due   { white-space: nowrap; font-weight: 600; }
.ai-source { color: #64748b; font-size: 12px; }
.auto-flag { color: #d97706; cursor: help; }

/* ── Agenda cards ── */
.agenda-card {
  border: 1px solid #e2e8f0; border-radius: 8px;
  margin-bottom: 14px; overflow: hidden;
}
.agenda-head {
  background: #f8fafc; padding: 9px 15px;
  display: flex; align-items: center; gap: 9px;
  border-bottom: 1px solid #e2e8f0;
}
.agenda-ws { font-weight: 700; font-size: 14px; flex: 1; }
.agenda-block { padding: 10px 15px; }
.agenda-block + .agenda-block { border-top: 1px solid #f1f5f9; }
.block-label {
  font-size: 11px; font-weight: 700; color: #64748b;
  text-transform: uppercase; letter-spacing: .5px; margin-bottom: 6px;
}
.agenda-block ul { padding-left: 18px; }
.agenda-block li { font-size: 13px; margin-bottom: 4px; }
.due-badge {
  background: #fef3c7; color: #92400e; font-size: 10px;
  padding: 1px 5px; border-radius: 4px; font-weight: 700;
  margin-right: 5px;
}
.owner-text { color: #64748b; font-style: italic; font-size: 12px; }

/* ── Plain text areas ── */
.plain-text {
  display: none; width: 100%; font-family: "Consolas", monospace; font-size: 12px;
  background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px;
  padding: 10px; resize: vertical; height: 120px; margin-top: 8px;
  color: #1e293b;
}
.toggle-plain {
  background: transparent; border: none; color: #94a3b8; font-size: 11px;
  cursor: pointer; padding: 4px 0; text-decoration: underline;
}

/* ── Webex supplementary ── */
.supplementary {
  background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px;
  padding: 12px 15px; margin-top: 12px; font-size: 13px;
}
.supplementary ul { padding-left: 18px; margin-top: 6px; }
.supplementary li { margin-bottom: 3px; }

/* ── Print ── */
@media print {
  .copy-btn, .toggle-plain, .plain-text { display: none !important; }
  body { background: white; }
  .section { box-shadow: none; margin: 0 0 16px; page-break-inside: avoid; }
  .kpi-row, .ws-row { padding-left: 0; }
}
"""


_JS = """
function copyText(id) {
  var el = document.getElementById(id);
  if (!el) return;
  var text = el.value || el.innerText || '';
  var btn  = event.target;
  navigator.clipboard.writeText(text).then(function() {
    btn.textContent = 'Copied!';
    setTimeout(function() { btn.textContent = 'Copy'; }, 1500);
  }).catch(function() {
    el.style.display = 'block';
    el.select();
    document.execCommand('copy');
  });
}

function togglePlain(id, btn) {
  var el = document.getElementById(id);
  if (!el) return;
  if (el.style.display === 'block') {
    el.style.display = 'none';
    btn.textContent  = 'Show plain text';
  } else {
    el.style.display = 'block';
    btn.textContent  = 'Hide plain text';
  }
}
"""


def generate_html(pid: str, customer_name: str, notes: list,
                  all_action_items: list, webex_spaces: list,
                  emails: list, date_range: tuple, today: datetime) -> str:

    start_str = date_range[0].strftime("%b %d")
    end_str   = date_range[1].strftime("%b %d, %Y")

    workstreams = sorted(set(
        [n.workstream for n in notes] +
        [ai.workstream for ai in all_action_items]
    ))

    ws_next_steps: dict = {}
    for note in notes:
        ws = note.workstream
        ws_next_steps.setdefault(ws, []).extend(note.next_steps)

    overdue_count = sum(
        1 for ai in all_action_items
        if ai.due_date and ai.due_date < today
    )

    # ── Workstream status chips ──────────────────────────────────────────────
    chip_html = ""
    for ws in workstreams:
        status = get_status(ws, all_action_items, today)
        fg, bg = STATUS_COLORS.get(status, ("#1e3a5f", "#e0e7ff"))
        chip_html += (
            f'<span class="chip" style="background:{bg};color:{fg};">'
            f'{_e(ws)} &nbsp;—&nbsp; {status}</span>\n'
        )

    # ── Accomplished section ─────────────────────────────────────────────────
    acc_cards = ""
    for note in sorted(notes, key=lambda n: n.date):
        dec_html = ""
        if note.key_decisions:
            lis = "".join(f"<li>{_e(d)}</li>" for d in note.key_decisions[:4])
            dec_html = f'<ul class="decisions-list">{lis}</ul>'
        rec_html = ""
        if note.recording_url:
            rec_html = (
                f' <a href="{_e(note.recording_url)}" target="_blank" class="rec-link">'
                f'&#9654; Recording</a>'
            )
        att_html = ""
        if note.attendees:
            short = ", ".join(note.attendees[:6])
            if len(note.attendees) > 6:
                short += f" +{len(note.attendees)-6} more"
            att_html = f'<div class="attendees-line">&#128101; {_e(short)}</div>'

        acc_cards += f"""
<div class="meeting-card">
  <div class="meeting-head">
    <span class="m-date">{note.date.strftime("%b %d")}</span>
    <span class="m-title">{_e(_short_title(note.title))}</span>
    <span class="ws-tag">{_e(note.workstream)}</span>
    {rec_html}
  </div>
  {att_html}
  <p class="summary-p">{_e(note.summary[:420])}</p>
  {dec_html}
</div>"""

    if not acc_cards:
        acc_cards = "<p><em>No meeting notes found in the last 7 days.</em></p>"

    webex_supp = ""
    if webex_spaces:
        items_html = "".join(
            f"<li><strong>{_e(sp['space_name'])}</strong> — {sp['message_count']} messages</li>"
            for sp in webex_spaces
        )
        webex_supp = f'<div class="supplementary"><strong>Webex Spaces Active This Week:</strong><ul>{items_html}</ul></div>'

    acc_plain = plain_accomplished(notes)

    # ── Action Items section ─────────────────────────────────────────────────
    ai_rows = ""
    for ai in sorted(all_action_items, key=lambda x: (x.due_date or datetime.max)):
        flag        = ' <span class="auto-flag" title="Due date auto-assigned (not stated in meeting)">&#9873;</span>' if ai.due_auto else ""
        due_str     = ai.due_date.strftime("%m/%d/%Y") if ai.due_date else "TBD"
        overdue_cls = "overdue" if ai.due_date and ai.due_date < today else ""
        ai_rows += (
            f'<tr class="{overdue_cls}">'
            f'<td class="ai-action">{_e(ai.action)}</td>'
            f'<td class="ai-owner">{_e(ai.owner)}</td>'
            f'<td class="ai-due">{_e(due_str)}{flag}</td>'
            f'<td class="ai-source">{_e(ai.source)}</td>'
            f'<td><span class="ws-tag-sm">{_e(ai.workstream)}</span></td>'
            f'</tr>\n'
        )

    if not ai_rows:
        ai_rows = '<tr><td colspan="5"><em>No action items found.</em></td></tr>'

    ai_plain = plain_action_items(all_action_items)

    # ── Agendas section ──────────────────────────────────────────────────────
    agenda_cards = ""
    for ws in workstreams:
        ws_items = [ai for ai in all_action_items if ai.workstream == ws]
        ws_ns    = ws_next_steps.get(ws, [])
        if not ws_items and not ws_ns:
            continue

        status = get_status(ws, all_action_items, today)
        fg, bg = STATUS_COLORS.get(status, ("#1e3a5f", "#e0e7ff"))
        ws_id  = re.sub(r"\W+", "_", ws)

        ns_block = ""
        if ws_ns:
            lis = "".join(f"<li>{_e(ns)}</li>" for ns in ws_ns[:6])
            ns_block = (
                f'<div class="agenda-block">'
                f'<div class="block-label">Carry-over / Next Steps</div>'
                f'<ul>{lis}</ul></div>'
            )

        item_lis = ""
        for ai in sorted(ws_items, key=lambda x: (x.due_date or datetime.max)):
            flag    = "*" if ai.due_auto else ""
            due_str = (ai.due_date.strftime("%m/%d") + flag) if ai.due_date else "TBD"
            item_lis += (
                f'<li><span class="due-badge">{_e(due_str)}</span>'
                f'{_e(ai.action)} '
                f'<span class="owner-text">— {_e(ai.owner)}</span></li>\n'
            )
        items_block = ""
        if item_lis:
            items_block = (
                f'<div class="agenda-block">'
                f'<div class="block-label">Open Action Items</div>'
                f'<ul>{item_lis}</ul></div>'
            )

        agenda_plain = plain_agenda(ws, ws_items, ws_ns)

        agenda_cards += f"""
<div class="agenda-card">
  <div class="agenda-head">
    <span class="agenda-ws">{_e(ws)}</span>
    <span class="chip-sm" style="background:{bg};color:{fg};">{status}</span>
    <button class="copy-btn" onclick="copyText('pa-{ws_id}')">Copy</button>
  </div>
  {ns_block}
  {items_block}
  <textarea id="pa-{ws_id}" class="plain-text">{_e(agenda_plain)}</textarea>
</div>"""

    if not agenda_cards:
        agenda_cards = "<p><em>No agendas generated.</em></p>"

    # ── Emails section ───────────────────────────────────────────────────────
    email_rows = ""
    for em in emails[:30]:
        email_rows += (
            f'<tr>'
            f'<td>{em.date.strftime("%m/%d")}</td>'
            f'<td>{_e(em.subject[:80])}</td>'
            f'<td>{_e(em.sender[:30])}</td>'
            f'<td><span class="ws-tag-sm">{_e(em.workstream)}</span></td>'
            f'</tr>\n'
        )

    emails_section = ""
    if email_rows:
        emails_section = f"""
<div class="section">
  <div class="section-head"><h2>&#128231; Emails This Week ({len(emails)})</h2></div>
  <div class="section-body" style="overflow-x:auto;">
    <table class="ai-table">
      <thead><tr><th>Date</th><th>Subject</th><th>From</th><th>Workstream</th></tr></thead>
      <tbody>{email_rows}</tbody>
    </table>
  </div>
</div>"""

    # ── Assemble ─────────────────────────────────────────────────────────────
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Weekly Summary – {_e(pid)} – {_e(end_str)}</title>
<style>{_CSS}</style>
</head>
<body>

<div class="report-header">
  <span class="pid-badge">{_e(pid)}</span>
  <div class="header-text">
    <h1>Weekly Summary &mdash; {_e(customer_name)}</h1>
    <div class="sub">{_e(start_str)} &ndash; {_e(end_str)} &nbsp;&middot;&nbsp; {len(notes)} meetings &nbsp;&middot;&nbsp; Generated {today.strftime("%B %d, %Y")}</div>
  </div>
</div>

<div class="kpi-row">
  <div class="kpi-tile"><div class="kpi-num">{len(notes)}</div><div class="kpi-label">Meetings</div></div>
  <div class="kpi-tile"><div class="kpi-num">{len(all_action_items)}</div><div class="kpi-label">Action Items</div></div>
  <div class="kpi-tile"><div class="kpi-num">{overdue_count}</div><div class="kpi-label">Overdue</div></div>
  <div class="kpi-tile"><div class="kpi-num">{len(workstreams)}</div><div class="kpi-label">Workstreams</div></div>
  <div class="kpi-tile"><div class="kpi-num">{len(emails)}</div><div class="kpi-label">Emails</div></div>
</div>

<div class="ws-row">{chip_html}</div>

<!-- Accomplished -->
<div class="section">
  <div class="section-head">
    <h2>&#10003; Accomplished This Week</h2>
    <button class="copy-btn" onclick="copyText('plain-acc')">Copy</button>
  </div>
  <div class="section-body">
    {acc_cards}
    {webex_supp}
  </div>
  <div class="section-foot">
    <button class="toggle-plain" onclick="togglePlain('plain-acc', this)">Show plain text</button>
    <textarea id="plain-acc" class="plain-text">{_e(acc_plain)}</textarea>
  </div>
</div>

<!-- Action Items -->
<div class="section">
  <div class="section-head">
    <h2>&#9889; Action Items</h2>
    <div style="display:flex;gap:10px;align-items:center;">
      <span style="font-size:11px;opacity:.7;">&#9873; = due date auto-assigned</span>
      <button class="copy-btn" onclick="copyText('plain-ai')">Copy</button>
    </div>
  </div>
  <div class="section-body" style="overflow-x:auto;">
    <table class="ai-table">
      <thead><tr><th>Action</th><th>Owner</th><th>Due Date</th><th>Source</th><th>Workstream</th></tr></thead>
      <tbody>{ai_rows}</tbody>
    </table>
  </div>
  <div class="section-foot">
    <button class="toggle-plain" onclick="togglePlain('plain-ai', this)">Show plain text</button>
    <textarea id="plain-ai" class="plain-text">{_e(ai_plain)}</textarea>
  </div>
</div>

<!-- Draft Agendas -->
<div class="section">
  <div class="section-head"><h2>&#128203; Draft Agendas by Workstream</h2></div>
  <div class="section-body">{agenda_cards}</div>
</div>

{emails_section}

<script>{_JS}</script>
</body>
</html>"""


# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════

def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Weekly Summary report for a PID.")
    parser.add_argument("pid", help="Project ID, e.g. F51119")
    args = parser.parse_args()
    pid  = args.pid.upper()

    today  = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    cutoff = today - timedelta(days=LOOKBACK_DAYS)

    print(f"\n── Weekly Summary {pid} {'─'*(48-len(pid))}", flush=True)
    print(f"   Period: {cutoff.strftime('%Y-%m-%d')} → {today.strftime('%Y-%m-%d')}", flush=True)

    # 1. Locate project folder
    print("\n[1/6] Locating project folder...", flush=True)
    pid_folder = find_pid_folder(pid)
    if not pid_folder:
        sys.exit(f"  ERROR: No folder for '{pid}' under:\n    {SHAREPOINT_BASE}")
    print(f"  OK: {pid_folder}", flush=True)

    customer_name    = CUSTOMER_NAMES_BY_PID.get(pid, f"Project {pid}")
    customer_domains = CUSTOMER_DOMAINS_BY_PID.get(pid, set())

    # 2. Roster
    print("\n[2/6] Reading roster...", flush=True)
    contact_file = pid_folder / "Customer Profile" / "Project Contact List.xlsx"
    roster       = read_roster(contact_file)
    print(f"  Roster: {len(roster)} contacts", flush=True)

    # 3. Meeting notes
    print("\n[3/6] Reading meeting notes...", flush=True)
    mm_folder = pid_folder / "Meeting Minutes"
    notes = []
    if mm_folder.exists():
        for md_file in sorted(mm_folder.glob("*.md")):
            note = parse_meeting_notes(md_file, cutoff)
            if note:
                notes.append(note)
                print(f"  ✓ {md_file.name}  [{note.workstream}]", flush=True)
    else:
        print("  No Meeting Minutes folder found.", flush=True)
    print(f"  Notes found: {len(notes)}", flush=True)

    # 4. Webex API
    print("\n[4/6] Pulling Webex data...", flush=True)
    webex_spaces = pull_webex_data(pid, customer_name, cutoff)
    print(f"  Spaces with activity: {len(webex_spaces)}", flush=True)

    # 5. Outlook emails
    print("\n[5/6] Pulling Outlook emails...", flush=True)
    emails = pull_outlook_emails(pid, roster, customer_domains, cutoff)

    # 6. Generate report
    print("\n[6/6] Generating report...", flush=True)
    all_action_items = [ai for note in notes for ai in note.action_items]
    print(f"  Action items collected: {len(all_action_items)}", flush=True)

    OUTPUT_BASE.mkdir(parents=True, exist_ok=True)
    out_file = OUTPUT_BASE / f"PID_{pid}_{today.strftime('%Y-%m-%d')}.html"

    html_content = generate_html(
        pid              = pid,
        customer_name    = customer_name,
        notes            = notes,
        all_action_items = all_action_items,
        webex_spaces     = webex_spaces,
        emails           = emails,
        date_range       = (cutoff, today),
        today            = today,
    )

    out_file.write_text(html_content, encoding="utf-8")
    print(f"\n  ✓ Report saved: {out_file}", flush=True)

    import subprocess
    subprocess.Popen(["start", "", str(out_file)], shell=True)
    print("  Opened in browser.", flush=True)


if __name__ == "__main__":
    main()
