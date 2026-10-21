#!/usr/bin/env python3
"""
Secretary Agent — Weekly PID-Level Action Item Report
Scans filed meeting notes (.md) and Webex recap emails for Houston's action items.

Usage:  python secretary_agent.py
        python secretary_agent.py --days 14
Trigger phrase: "Secretary Agent" or "Secretary Agent report"
"""

import re, json, sys, hashlib
import win32com.client
from datetime import datetime, timedelta, date
from pathlib import Path
from bs4 import BeautifulSoup

# ── Config ────────────────────────────────────────────────────────────────────

SHAREPOINT_BASE = Path(r"C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS")
OUTPUT_DIR      = Path(r"C:\Users\hodowns\OneDrive - Cisco\Desktop\Secretary Agent")
OUTPUT_FILE     = OUTPUT_DIR / "action-items.html"
WEBEX_BOT       = "messenger@webex.com"
DEFAULT_DAYS    = 8
PID_RE          = re.compile(r'\b[A-Z]\d{5}\b')

OWNER_RE = re.compile(
    r'\b(houston\s+downs?|hodowns|'
    r'hd\s+(to|will|needs?\s+to|should|can)\b|'
    r'pm\s+(to|will|needs?\s+to)\b|'
    r'i\s+will\b|i\'ll\b|i\s+need\s+to\b|'
    r'i\'m\s+going\s+to\b|i\s+am\s+going\s+to\b|'
    r'i\s+can\s+|i\s+should\b|'
    r'let\s+me\b|i\s+want\s+to\b)',
    re.IGNORECASE
)

FOOTER_RE = re.compile(
    r'need help\?|help\.webex\.com|cisco and/or its affiliates'
    r'|all rights reserved|privacy statement|terms of service'
    r'|view meeting content|view meeting recap|claude notification incident',
    re.IGNORECASE,
)

PID_NAMES = {
    'F51119': 'City of Houston — LCS',
    'E55085': 'City of Houston — WAN Security',
    'F29426': 'Jacksonville Sheriff — VVS SAS',
    'C19602': 'Ada County — VVS JMS',
    'F24373': 'UChicago Medicine — LCS 2Yr',
    'F18212': 'UChicago Medicine — SAS/LCS',
    'F48852': 'Maryland DHS — CAT ISE SAS',
    'F40696': 'NYS Senate — SAS',
    'F46100': 'State of Wisconsin — VVS SAS',
    'C16507': 'MN DOC — VVS',
    'F43673': 'Hennepin County — CC AST',
    'F49302': 'GDIT — Firewall SAS',
    'F49303': 'GDIT — Firewall SAS II',
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def fix_garbled(text):
    """Split CamelCase-concatenated words: 'HoustonDownswill' → 'Houston Downs will'."""
    text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
    text = re.sub(r'([A-Z]{2,})([A-Z][a-z])', r'\1 \2', text)
    return text.strip()


def is_mine(text):
    """Return True if the text signals Houston as the owner."""
    return bool(OWNER_RE.search(text))


def make_id(pid, action):
    key = f"{pid}::{action[:80].lower().strip()}"
    return hashlib.md5(key.encode()).hexdigest()[:12]


def parse_due_date(raw, identified_date):
    """
    Parse a due date string. Falls back to identified_date + 7 days.
    Returns a date object.
    """
    raw = raw.strip()
    fallback = identified_date + timedelta(days=7)
    if not raw or raw.upper() in ('TBD', 'N/A', '-', ''):
        return fallback
    for fmt in ('%m/%d/%y', '%m/%d/%Y', '%Y-%m-%d', '%B %d, %Y', '%b %d, %Y'):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return fallback


def fmt_date(d):
    return d.strftime('%m/%d/%y')


def date_status(due_date):
    """Return 'overdue', 'soon', or 'ok'."""
    today = date.today()
    if due_date < today:
        return 'overdue'
    if (due_date - today).days <= 2:
        return 'soon'
    return 'ok'


# ── Load existing state from previous HTML ────────────────────────────────────

def load_existing_state(path):
    """
    Parse the existing HTML for the embedded JSON data block.
    Returns (last_run_date, set_of_existing_ids).
    """
    if not path.exists():
        return None, set()
    try:
        soup = BeautifulSoup(path.read_text(encoding='utf-8'), 'html.parser')
        tag = soup.find('script', {'id': 'sa-data'})
        if not tag:
            return None, set()
        data = json.loads(tag.string)
        last_run = datetime.strptime(data.get('generated', ''), '%Y-%m-%d').date() if data.get('generated') else None
        existing_ids = {item['id'] for item in data.get('items', [])}
        return last_run, existing_ids
    except Exception:
        return None, set()


# ── Source: .md meeting notes ─────────────────────────────────────────────────

def scan_md_files(start_date, existing_ids):
    items = []
    seen_ids = set(existing_ids)

    for state_dir in sorted(SHAREPOINT_BASE.iterdir()):
        if not state_dir.is_dir():
            continue
        for proj_dir in sorted(state_dir.iterdir()):
            if not proj_dir.is_dir():
                continue
            pid_m = PID_RE.search(proj_dir.name)
            pid = pid_m.group(0) if pid_m else None
            mm_dir = proj_dir / 'Meeting Minutes'
            if not mm_dir.exists():
                continue
            for md_file in sorted(mm_dir.glob('*.md'), reverse=True):
                date_m = re.match(r'(\d{4}-\d{2}-\d{2})', md_file.name)
                if not date_m:
                    continue
                file_date = datetime.strptime(date_m.group(1), '%Y-%m-%d').date()
                if file_date < start_date:
                    continue
                new_items = _parse_md(md_file, pid, file_date, seen_ids)
                for it in new_items:
                    seen_ids.add(it['id'])
                items.extend(new_items)

    print(f"  .md files: {len(items)} item(s) found", flush=True)
    return items


def _parse_md(path, pid, file_date, seen_ids):
    items = []
    try:
        text = path.read_text(encoding='utf-8', errors='ignore')
    except Exception:
        return items

    # Meeting title
    title_m = re.search(r'^#\s+(.+)$', text, re.MULTILINE)
    raw_title = title_m.group(1).strip() if title_m else path.stem
    # Shorten for source label
    short_title = re.sub(r'\s*\[(?:EXTERNAL|INTERNAL|EXT|INT)\]\s*', ' ', raw_title, flags=re.IGNORECASE).strip()
    short_title = re.sub(r'\s+', ' ', short_title)[:50]
    source_label = f"Meeting Notes — {short_title}, {file_date.strftime('%b %d')}"

    # PID from frontmatter or title
    if not pid:
        fm_m = re.search(r'^project_id:\s*(\S+)', text, re.MULTILINE)
        pid = fm_m.group(1) if fm_m else None
    if not pid:
        pid_m = PID_RE.search(raw_title)
        pid = pid_m.group(0) if pid_m else None

    # Action items table
    table_m = re.search(
        r'##\s+Action Items\s*\n+\|[^\n]+\n\|[-| :]+\n((?:\|[^\n]+\n?)*)',
        text, re.IGNORECASE
    )
    if not table_m:
        return items

    for row in table_m.group(1).strip().splitlines():
        cols = [c.strip() for c in row.split('|')]
        cols = [c for c in cols if c]
        if len(cols) < 2:
            continue
        # Table format: | # | Action | Owner | Due Date |
        if len(cols) == 4:
            _, action, owner, due_raw = cols
        elif len(cols) == 3:
            _, action, owner = cols
            due_raw = ''
        else:
            action = cols[1] if len(cols) > 1 else cols[0]
            owner  = cols[2] if len(cols) > 2 else ''
            due_raw = cols[3] if len(cols) > 3 else ''

        action = action.strip()
        if not action or action.startswith('-') or FOOTER_RE.search(action):
            continue

        uncertain = False
        if not is_mine(owner + ' ' + action):
            # Still flag if PID is present and owner is vague
            uncertain = True

        due_date = parse_due_date(due_raw, file_date)
        item_id  = make_id(pid or 'NONE', action)

        if item_id in seen_ids:
            continue

        items.append({
            'id':              item_id,
            'pid':             pid or 'ADMIN',
            'action':          action,
            'owner':           owner,
            'due_date':        fmt_date(due_date),
            'identified_date': file_date.isoformat(),
            'source':          source_label,
            'source_type':     'meeting-notes',
            'uncertain':       uncertain,
        })

    return items


# ── Source: Webex recap emails ────────────────────────────────────────────────

def scan_webex_emails(outlook, start_date, existing_ids):
    items = []
    seen_ids = set(existing_ids)
    ns    = outlook.GetNamespace('MAPI')
    inbox = ns.GetDefaultFolder(6)

    mail_items = inbox.Items
    mail_items.Sort('[ReceivedTime]', True)
    cutoff = start_date.strftime('%m/%d/%Y %I:%M %p')
    try:
        mail_items = mail_items.Restrict(f"[ReceivedTime] >= '{cutoff}'")
    except Exception:
        pass

    processed = 0
    for msg in mail_items:
        try:
            if (getattr(msg, 'SenderEmailAddress', '') or '').lower() != WEBEX_BOT:
                continue
            subj = (getattr(msg, 'Subject', '') or '')
            if 'meeting content is available' not in subj.lower():
                continue

            body   = (getattr(msg, 'Body', '') or '').replace('\r\n', '\n').replace('\r', '\n')
            lines  = [l.strip() for l in body.split('\n') if l.strip()]

            # Meeting date
            meeting_date = None
            for line in lines:
                if re.match(r'(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),', line):
                    try:
                        meeting_date = datetime.strptime(line.strip(), '%A, %B %d, %Y').date()
                    except ValueError:
                        pass
                    break
            if not meeting_date:
                continue
            if meeting_date < start_date:
                continue

            # PIDs in subject/body
            pids_found = list(dict.fromkeys(PID_RE.findall(subj + ' ' + body)))
            primary_pid = pids_found[0] if pids_found else None

            # Meeting title
            title_raw = ''
            for i, line in enumerate(lines):
                if 'meeting content is available' in line.lower() and i + 1 < len(lines):
                    title_raw = lines[i + 1]
                    break
            short_title = re.sub(r'\[(?:EXTERNAL|INTERNAL|EXT|INT)\]\s*', '', title_raw, flags=re.I).strip()[:45]
            source_label = f"Webex Recap — {short_title or 'Meeting'}, {meeting_date.strftime('%b %d')}"

            # Action items section
            in_actions = False
            for line in lines:
                if re.search(r'Action items.{0,5}AI.generated', line, re.IGNORECASE):
                    in_actions = True
                    continue
                if in_actions:
                    if re.search(r'^(Recording|Chapters|Overview|Meeting summary|Detailed)', line, re.IGNORECASE):
                        break
                    clean = re.sub(r'^[\*\-•]\s*', '', line).strip()
                    clean = re.sub(r'\s*<https?://\S+>', '', clean).strip()
                    if not clean or FOOTER_RE.search(clean):
                        continue
                    clean = fix_garbled(clean)
                    if len(clean) < 10:
                        continue

                    uncertain = not is_mine(clean)
                    due_date  = meeting_date + timedelta(days=7)
                    pid       = primary_pid or 'ADMIN'
                    item_id   = make_id(pid, clean)

                    if item_id in seen_ids:
                        continue

                    items.append({
                        'id':              item_id,
                        'pid':             pid,
                        'action':          clean,
                        'owner':           '',
                        'due_date':        fmt_date(due_date),
                        'identified_date': meeting_date.isoformat(),
                        'source':          source_label,
                        'source_type':     'webex-email',
                        'uncertain':       uncertain,
                    })
                    seen_ids.add(item_id)

            processed += 1
        except Exception:
            continue

    print(f"  Webex emails: {processed} recap(s) scanned, {len(items)} item(s) found", flush=True)
    return items


# ── Source: broad Outlook inbox + sent scan ───────────────────────────────────

# Keywords that signal an email likely contains an action for Houston
ACTION_SUBJECT_RE = re.compile(
    r'\b(action\s+required|action\s+needed|please\s+complete|please\s+review|'
    r'please\s+submit|please\s+respond|please\s+confirm|reminder|due\s+date|'
    r'deadline|overdue|past\s+due|follow.?up|complete\s+by|submit\s+by|'
    r'required\s+training|mandatory|compliance|attestation|certification|'
    r'cost\s+and\s+revenue|c&r|time.?card|expense|approval\s+needed|'
    r'your\s+action|needs\s+your|awaiting\s+your)\b',
    re.IGNORECASE,
)

SKIP_SENDERS_RE = re.compile(
    r'messenger@webex\.com|noreply|no-reply|donotreply|do-not-reply|'
    r'mailer-daemon|postmaster|notifications@|alerts@|digest@',
    re.IGNORECASE,
)

SKIP_SUBJECT_RE = re.compile(
    r'^(accepted|declined|tentative|canceled?|fw:|fwd:|meeting content is available|'
    r'automatic reply|out of office|newsletter|unsubscribe)',
    re.IGNORECASE,
)


def scan_outlook_email(outlook, start_date, existing_ids):
    """
    Scan inbox (To-addressed) and sent items for admin/non-PID action items.
    Inbox: emails addressed TO Houston with action-signal subjects.
    Sent:  emails where Houston made a first-person commitment.
    """
    items    = []
    seen_ids = set(existing_ids)
    ns       = outlook.GetNamespace('MAPI')
    cutoff   = start_date.strftime('%m/%d/%Y %I:%M %p')
    processed = 0

    def _process_folder(folder, check_ownership):
        nonlocal processed
        mail_items = folder.Items
        mail_items.Sort('[ReceivedTime]', True)
        try:
            mail_items = mail_items.Restrict(f"[ReceivedTime] >= '{cutoff}'")
        except Exception:
            pass

        for msg in mail_items:
            try:
                subj   = (getattr(msg, 'Subject', '') or '')
                sender = (getattr(msg, 'SenderEmailAddress', '') or '').lower()
                body   = (getattr(msg, 'Body', '') or '').replace('\r\n', '\n').replace('\r', '\n')
                to_str = (getattr(msg, 'To', '') or '').lower()

                # Skip noise
                if SKIP_SENDERS_RE.search(sender):
                    continue
                if SKIP_SUBJECT_RE.search(subj):
                    continue
                if FOOTER_RE.search(subj):
                    continue
                # Skip pure Webex meeting recaps (handled separately)
                if 'meeting content is available' in subj.lower():
                    continue

                pid_in_subject = PID_RE.search(subj)

                if check_ownership:
                    # Sent items: look for Houston's own commitments in body
                    if not is_mine(body[:2000]):
                        continue
                    pid = pid_in_subject.group(0) if pid_in_subject else 'ADMIN'
                    # Extract the commitment sentence
                    action = _extract_commitment(body, subj)
                    if not action:
                        continue
                    uncertain = False
                else:
                    # Inbox: must be addressed TO Houston (not just CC)
                    if 'hodowns' not in to_str and 'houston downs' not in to_str:
                        continue
                    # Must have action signal in subject OR body
                    if not ACTION_SUBJECT_RE.search(subj) and not ACTION_SUBJECT_RE.search(body[:500]):
                        continue
                    pid = pid_in_subject.group(0) if pid_in_subject else 'ADMIN'
                    action = subj.strip()
                    uncertain = True  # inbox items always go to confirmation unless PID makes it clear

                # Parse date
                recv_date = _parse_received_date(getattr(msg, 'ReceivedTime', None))
                due_date  = recv_date + timedelta(days=7)
                # Look for explicit date in subject/body
                due_date  = _sniff_due_date(subj + ' ' + body[:500], recv_date) or due_date

                item_id = make_id(pid, action)
                if item_id in seen_ids:
                    continue
                seen_ids.add(item_id)

                source_label = f"Email — {subj[:45]}, {recv_date.strftime('%b %d')}"
                items.append({
                    'id':              item_id,
                    'pid':             pid,
                    'action':          action,
                    'owner':           'Houston Downs' if check_ownership else '',
                    'due_date':        fmt_date(due_date),
                    'identified_date': recv_date.isoformat(),
                    'source':          source_label,
                    'source_type':     'email',
                    'uncertain':       uncertain,
                })
                processed += 1
            except Exception:
                continue

    # Scan inbox (To-addressed action emails)
    try:
        _process_folder(ns.GetDefaultFolder(6), check_ownership=False)
    except Exception as e:
        print(f"  Inbox scan error: {e}", flush=True)

    # Scan sent items (Houston's own commitments)
    try:
        _process_folder(ns.GetDefaultFolder(5), check_ownership=True)
    except Exception as e:
        print(f"  Sent scan error: {e}", flush=True)

    print(f"  Outlook email: {processed} item(s) found (inbox actions + sent commitments)", flush=True)
    return items


def _extract_commitment(body, subj):
    """Pull the first sentence containing an ownership signal from email body."""
    for line in body.split('\n'):
        line = line.strip()
        if len(line) < 15:
            continue
        if is_mine(line):
            clean = re.sub(r'\s+', ' ', line).strip()
            return clean[:200] if clean else None
    return subj.strip() or None


def _parse_received_date(recv):
    """Convert Outlook ReceivedTime (pywintypes.datetime) to date."""
    try:
        return datetime.strptime(str(recv)[:10], '%Y-%m-%d').date()
    except Exception:
        return date.today()


def _sniff_due_date(text, identified_date):
    """Try to extract an explicit due date from text. Returns date or None."""
    patterns = [
        (r'\bby\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})', '%B %d %Y'),
        (r'\bdue\s+(\d{1,2}/\d{1,2}/\d{2,4})', None),
        (r'\bby\s+(\d{1,2}/\d{1,2}/\d{2,4})', None),
    ]
    for pattern, fmt in patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            try:
                if fmt:
                    parts = ' '.join(m.groups())
                    return datetime.strptime(parts, fmt).date()
                else:
                    raw = m.group(1)
                    for f in ('%m/%d/%y', '%m/%d/%Y'):
                        try:
                            return datetime.strptime(raw, f).date()
                        except ValueError:
                            continue
            except Exception:
                continue
    return None


# ── HTML generation ───────────────────────────────────────────────────────────

def generate_html(all_items, start_date, end_date, prev_items):
    today = date.today()

    # Merge new + existing (prev items not in new scan)
    merged = {it['id']: it for it in prev_items}
    for it in all_items:
        merged[it['id']] = it
    items = list(merged.values())

    mine      = [it for it in items if not it.get('uncertain')]
    uncertain = [it for it in items if it.get('uncertain')]
    overdue_n = sum(1 for it in mine if date_status(_parse_date(it['due_date'])) == 'overdue')

    # Group by PID
    by_pid = {}
    for it in mine:
        by_pid.setdefault(it['pid'], []).append(it)
    # Sort PIDs: known first (by name), then UNASSIGNED
    def pid_sort(pid):
        if pid == 'ADMIN': return 'ZZZ'  # Admin always last
        return PID_NAMES.get(pid, pid)
    sorted_pids = sorted(by_pid.keys(), key=pid_sort)

    data_json = json.dumps({
        'generated': today.isoformat(),
        'start':     start_date.isoformat(),
        'end':       end_date.isoformat(),
        'items':     items,
    }, indent=2)

    # Build item rows
    def item_row(it):
        due  = _parse_date(it['due_date'])
        status = date_status(due)
        status_class = {'overdue': 'due-overdue', 'soon': 'due-soon', 'ok': 'due-ok'}[status]
        due_icon = {'overdue': '⚑ ', 'soon': '◉ ', 'ok': ''}[status]
        src_type_label = {'meeting-notes': 'Notes', 'webex-email': 'Webex', 'transcript': 'Transcript'}.get(it.get('source_type',''), 'Source')
        return f'''<li class="item" data-id="{it['id']}">
          <label class="item-check">
            <input type="checkbox" data-id="{it['id']}">
            <span class="checkmark"></span>
          </label>
          <span class="item-action">{_esc(it['action'])}</span>
          <span class="item-meta">
            <span class="due-chip {status_class}">{due_icon}{_esc(it['due_date'])}</span>
            <span class="src-label"><span class="src-type">{src_type_label}</span>{_esc(_short_source(it['source']))}</span>
          </span>
        </li>'''

    def pid_section(pid, its):
        name = PID_NAMES.get(pid, pid if pid != 'ADMIN' else 'Admin')
        rows = '\n'.join(item_row(it) for it in sorted(its, key=lambda x: _parse_date(x['due_date'])))
        n = len(its)
        ov = sum(1 for it in its if date_status(_parse_date(it['due_date'])) == 'overdue')
        badge = f'<span class="badge {"badge-overdue" if ov else ""}">{n}</span>'
        ov_note = f'<span class="overdue-note">{ov} overdue</span>' if ov else ''
        return f'''<section class="pid-group" data-pid="{pid}">
        <button class="pid-header" aria-expanded="true">
          <span class="pid-code">{pid}</span>
          <span class="pid-name">{_esc(name)}</span>
          <span class="pid-right">{ov_note}{badge}<span class="chevron">▾</span></span>
        </button>
        <ul class="item-list">{rows}</ul>
        <div class="completed-row">
          <button class="show-done-btn" data-pid="{pid}">Show completed</button>
          <ul class="done-list"></ul>
        </div>
      </section>'''

    pid_sections_html = '\n'.join(pid_section(p, by_pid[p]) for p in sorted_pids)

    uncertain_rows = '\n'.join(item_row(it) for it in sorted(uncertain, key=lambda x: _parse_date(x['due_date'])))
    uncertain_section = f'''<section class="pid-group uncertain-group" data-pid="CONFIRM">
      <button class="pid-header" aria-expanded="true">
        <span class="pid-code uncertain-icon">?</span>
        <span class="pid-name">Needs Your Confirmation</span>
        <span class="pid-right"><span class="badge">{len(uncertain)}</span><span class="chevron">▾</span></span>
      </button>
      <ul class="item-list">{uncertain_rows}</ul>
      <div class="completed-row">
        <button class="show-done-btn" data-pid="CONFIRM">Show completed</button>
        <ul class="done-list"></ul>
      </div>
    </section>''' if uncertain else ''

    total_open = len(mine)

    return f'''<title>Secretary Agent</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap">

<style>
/* ── Tokens ── */
:root {{
  --bg:         #F5F6FA;
  --surface:    #FFFFFF;
  --surface-2:  #EEF0F7;
  --border:     #DDE1EB;
  --text:       #1A1D26;
  --text-2:     #5A6070;
  --text-3:     #8891A4;
  --accent:     #2B6CB0;
  --accent-dim: #EBF2FA;
  --overdue:    #C53030;
  --overdue-bg: #FFF5F5;
  --soon:       #B7791F;
  --soon-bg:    #FFFBEB;
  --ok:         #2F855A;
  --ok-bg:      #F0FFF4;
  --check:      #2B6CB0;
  --radius:     6px;
  --shadow:     0 1px 3px rgba(0,0,0,.06), 0 1px 2px rgba(0,0,0,.04);
}}
@media (prefers-color-scheme: dark) {{
  :root:not([data-theme="light"]) {{
    --bg:         #12141C;
    --surface:    #1C1F2B;
    --surface-2:  #22263A;
    --border:     #2A2F42;
    --text:       #E4E7F0;
    --text-2:     #8892A6;
    --text-3:     #5A6272;
    --accent:     #63A0D8;
    --accent-dim: #1A2A3A;
    --overdue:    #FC8181;
    --overdue-bg: #1A0A0A;
    --soon:       #F6AD55;
    --soon-bg:    #1A1200;
    --ok:         #68D391;
    --ok-bg:      #0A1A10;
    --check:      #63A0D8;
    --shadow:     0 1px 3px rgba(0,0,0,.3);
  }}
}}
:root[data-theme="dark"] {{
  --bg:         #12141C;
  --surface:    #1C1F2B;
  --surface-2:  #22263A;
  --border:     #2A2F42;
  --text:       #E4E7F0;
  --text-2:     #8892A6;
  --text-3:     #5A6272;
  --accent:     #63A0D8;
  --accent-dim: #1A2A3A;
  --overdue:    #FC8181;
  --overdue-bg: #1A0A0A;
  --soon:       #F6AD55;
  --soon-bg:    #1A1200;
  --ok:         #68D391;
  --ok-bg:      #0A1A10;
  --check:      #63A0D8;
  --shadow:     0 1px 3px rgba(0,0,0,.3);
}}

/* ── Reset & base ── */
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
html {{ font-size: 15px; -webkit-font-smoothing: antialiased; }}
body {{
  font-family: 'Inter', system-ui, sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
  padding-bottom: 4rem;
}}

/* ── Top bar ── */
.topbar {{
  position: sticky; top: 0; z-index: 100;
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  padding: .75rem 1.5rem;
  display: flex; align-items: center; gap: 1.5rem;
  box-shadow: var(--shadow);
}}
.topbar-title {{
  font-size: .7rem; font-weight: 600;
  letter-spacing: .12em; text-transform: uppercase;
  color: var(--text-3);
  flex-shrink: 0;
}}
.topbar-stats {{
  display: flex; gap: 1.25rem; align-items: center;
  font-family: 'JetBrains Mono', monospace;
  font-size: .75rem; color: var(--text-2);
}}
.stat-open {{ color: var(--accent); font-weight: 500; }}
.stat-overdue {{ color: var(--overdue); font-weight: 500; }}
.stat-range {{ color: var(--text-3); }}
.topbar-right {{ margin-left: auto; }}
.theme-btn {{
  background: none; border: 1px solid var(--border); border-radius: 20px;
  padding: .25rem .65rem; font-size: .7rem; font-family: 'Inter', sans-serif;
  color: var(--text-2); cursor: pointer; letter-spacing: .04em;
}}
.theme-btn:hover {{ background: var(--surface-2); }}

/* ── Content wrapper ── */
.content {{
  max-width: 860px; margin: 2rem auto; padding: 0 1.25rem;
  display: flex; flex-direction: column; gap: 1rem;
}}

/* ── PID sections ── */
.pid-group {{
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  overflow: hidden;
}}
.pid-header {{
  width: 100%; background: none; border: none;
  display: flex; align-items: center; gap: .75rem;
  padding: .8rem 1rem; cursor: pointer;
  border-bottom: 1px solid var(--border);
  text-align: left;
}}
.pid-header:hover {{ background: var(--surface-2); }}
.pid-code {{
  font-family: 'JetBrains Mono', monospace;
  font-size: .72rem; font-weight: 500;
  color: var(--accent);
  background: var(--accent-dim);
  padding: .18rem .45rem; border-radius: 4px;
  letter-spacing: .03em; flex-shrink: 0;
}}
.uncertain-icon {{
  color: var(--soon); background: var(--soon-bg);
}}
.pid-name {{
  font-size: .82rem; font-weight: 500; color: var(--text);
  flex: 1;
}}
.pid-right {{
  display: flex; align-items: center; gap: .5rem; margin-left: auto;
}}
.overdue-note {{
  font-family: 'JetBrains Mono', monospace;
  font-size: .65rem; color: var(--overdue);
  letter-spacing: .02em;
}}
.badge {{
  font-family: 'JetBrains Mono', monospace;
  font-size: .65rem; font-weight: 500;
  color: var(--text-2);
  background: var(--surface-2);
  border: 1px solid var(--border);
  padding: .1rem .4rem; border-radius: 10px;
}}
.badge-overdue {{ color: var(--overdue); border-color: var(--overdue); }}
.chevron {{
  font-size: .65rem; color: var(--text-3);
  transition: transform .2s;
}}
.pid-header[aria-expanded="false"] .chevron {{ transform: rotate(-90deg); }}
.pid-header[aria-expanded="false"] + .item-list {{ display: none; }}
.pid-header[aria-expanded="false"] + .item-list + .completed-row {{ display: none; }}

/* ── Item list ── */
.item-list, .done-list {{
  list-style: none;
  border-top: none;
}}
.item {{
  display: flex; align-items: flex-start; gap: .65rem;
  padding: .7rem 1rem;
  border-bottom: 1px solid var(--border);
  transition: background .1s;
}}
.item:last-child {{ border-bottom: none; }}
.item:hover {{ background: var(--surface-2); }}

/* ── Checkbox ── */
.item-check {{
  flex-shrink: 0; margin-top: .1rem; cursor: pointer;
}}
.item-check input {{ display: none; }}
.checkmark {{
  display: inline-block;
  width: 16px; height: 16px;
  border: 1.5px solid var(--border);
  border-radius: 4px;
  background: var(--surface);
  position: relative;
  transition: border-color .15s, background .15s;
}}
.item-check input:checked + .checkmark {{
  background: var(--check);
  border-color: var(--check);
}}
.item-check input:checked + .checkmark::after {{
  content: '';
  position: absolute; top: 2px; left: 5px;
  width: 4px; height: 8px;
  border: 2px solid #fff;
  border-top: none; border-left: none;
  transform: rotate(45deg);
}}

/* ── Item content ── */
.item-action {{
  flex: 1; font-size: .83rem; line-height: 1.5; color: var(--text);
}}
.item.done .item-action {{
  text-decoration: line-through; color: var(--text-3);
}}
.item-meta {{
  display: flex; flex-direction: column; align-items: flex-end;
  gap: .3rem; flex-shrink: 0; margin-left: .5rem;
}}
.due-chip {{
  font-family: 'JetBrains Mono', monospace;
  font-size: .68rem; font-weight: 500;
  padding: .15rem .45rem; border-radius: 4px;
  white-space: nowrap; letter-spacing: .02em;
}}
.due-overdue {{ color: var(--overdue); background: var(--overdue-bg); }}
.due-soon    {{ color: var(--soon);    background: var(--soon-bg); }}
.due-ok      {{ color: var(--ok);      background: var(--ok-bg); }}
.src-label {{
  font-family: 'JetBrains Mono', monospace;
  font-size: .62rem; color: var(--text-3);
  text-align: right; max-width: 180px;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}}
.src-type {{
  font-weight: 500; color: var(--text-2);
  margin-right: .25rem;
}}

/* ── Completed row ── */
.completed-row {{
  padding: .4rem 1rem;
  border-top: 1px solid var(--border);
  background: var(--surface-2);
}}
.show-done-btn {{
  background: none; border: none;
  font-family: 'Inter', sans-serif;
  font-size: .73rem; color: var(--text-3);
  cursor: pointer; padding: 0;
}}
.show-done-btn:hover {{ color: var(--accent); }}
.done-list {{ margin-top: .35rem; }}
.done-list .item {{ opacity: .55; }}

/* ── Uncertain group ── */
.uncertain-group .pid-header {{ background: var(--soon-bg); }}
.uncertain-group .pid-header:hover {{ filter: brightness(.97); }}
</style>

<div class="topbar">
  <span class="topbar-title">Secretary Agent</span>
  <div class="topbar-stats">
    <span class="stat-range">{start_date.strftime('%b %d')} – {end_date.strftime('%b %d, %Y')}</span>
    <span class="stat-open">{total_open} open</span>
    {"<span class='stat-overdue'>" + str(overdue_n) + " overdue</span>" if overdue_n else ""}
  </div>
  <div class="topbar-right">
    <button class="theme-btn" onclick="toggleTheme()">theme</button>
  </div>
</div>

<div class="content">
  {pid_sections_html}
  {uncertain_section}
</div>

<script type="application/json" id="sa-data">
{data_json}
</script>

<script>
const STORAGE_KEY = 'secretary-agent-v1';
const THEME_KEY   = 'secretary-theme';

function getState() {{
  try {{ return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {{checked: []}}; }}
  catch {{ return {{checked: []}}; }}
}}
function saveState(s) {{
  try {{ localStorage.setItem(STORAGE_KEY, JSON.stringify(s)); }} catch {{}}
}}

function applyChecks() {{
  const s = getState();
  const done = new Set(s.checked || []);
  document.querySelectorAll('input[type=checkbox][data-id]').forEach(cb => {{
    const id = cb.dataset.id;
    cb.checked = done.has(id);
    syncRow(cb);
  }});
  document.querySelectorAll('.pid-group').forEach(refreshDone);
}}

function syncRow(cb) {{
  const li = cb.closest('.item');
  if (!li) return;
  li.classList.toggle('done', cb.checked);
}}

function refreshDone(section) {{
  const items = section.querySelectorAll('.item-list .item');
  const doneList = section.querySelector('.done-list');
  const btn = section.querySelector('.show-done-btn');
  let doneCount = 0;
  items.forEach(li => {{
    if (li.classList.contains('done')) {{
      li.style.display = 'none';
      doneCount++;
    }} else {{
      li.style.display = '';
    }}
  }});
  doneList.innerHTML = '';
  if (doneCount > 0) {{
    btn.textContent = `Show ${{doneCount}} completed`;
    btn.style.display = '';
  }} else {{
    btn.style.display = 'none';
  }}
}}

document.addEventListener('change', e => {{
  if (!e.target.matches('input[type=checkbox][data-id]')) return;
  const id = e.target.dataset.id;
  const s  = getState();
  const done = new Set(s.checked || []);
  if (e.target.checked) done.add(id); else done.delete(id);
  s.checked = [...done];
  saveState(s);
  syncRow(e.target);
  const section = e.target.closest('.pid-group');
  if (section) refreshDone(section);
}});

document.querySelectorAll('.show-done-btn').forEach(btn => {{
  btn.addEventListener('click', () => {{
    const section = btn.closest('.pid-group');
    const doneList = section.querySelector('.done-list');
    const items = section.querySelectorAll('.item-list .item.done');
    if (doneList.children.length) {{
      doneList.innerHTML = '';
      btn.textContent = btn.textContent.replace('Hide', 'Show');
    }} else {{
      items.forEach(li => {{
        const clone = li.cloneNode(true);
        clone.style.display = '';
        doneList.appendChild(clone);
      }});
      const n = items.length;
      btn.textContent = `Hide ${{n}} completed`;
    }}
  }});
}});

document.querySelectorAll('.pid-header').forEach(btn => {{
  btn.addEventListener('click', () => {{
    const expanded = btn.getAttribute('aria-expanded') === 'true';
    btn.setAttribute('aria-expanded', String(!expanded));
  }});
}});

function toggleTheme() {{
  const root = document.documentElement;
  const current = root.getAttribute('data-theme');
  if (!current) {{
    const dark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    root.setAttribute('data-theme', dark ? 'light' : 'dark');
  }} else {{
    root.setAttribute('data-theme', current === 'dark' ? 'light' : 'dark');
  }}
  try {{ localStorage.setItem(THEME_KEY, root.getAttribute('data-theme')); }} catch {{}}
}}

// Restore theme
try {{
  const t = localStorage.getItem(THEME_KEY);
  if (t) document.documentElement.setAttribute('data-theme', t);
}} catch {{}}

applyChecks();
</script>
'''


# ── Utility ───────────────────────────────────────────────────────────────────

def _esc(s):
    return (s or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def _parse_date(s):
    for fmt in ('%m/%d/%y', '%m/%d/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, fmt).date()
        except (ValueError, TypeError):
            continue
    return date.today() + timedelta(days=7)

def _short_source(s):
    # Strip long date prefix, keep meaningful label
    m = re.sub(r'^(Meeting Notes|Webex Recap|Transcript)\s*[—-]\s*', '', s)
    return m[:45]


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    days = DEFAULT_DAYS
    if '--days' in sys.argv:
        try: days = int(sys.argv[sys.argv.index('--days') + 1])
        except (IndexError, ValueError): pass

    today      = date.today()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Check existing file for last run date
    last_run, existing_ids = load_existing_state(OUTPUT_FILE)
    if last_run:
        start_date = last_run
        print(f"Existing report found. Scanning from last run: {last_run}", flush=True)
    else:
        start_date = today - timedelta(days=days)
        print(f"First run. Scanning back {days} days: from {start_date}", flush=True)

    print("Scanning sources...", flush=True)

    new_items = []

    # Source 1: .md meeting notes
    new_items.extend(scan_md_files(start_date, existing_ids))

    # Source 2: Webex recap emails + broad Outlook scan
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        seen_so_far = existing_ids | {it['id'] for it in new_items}
        new_items.extend(scan_webex_emails(outlook, start_date, seen_so_far))
        seen_so_far = existing_ids | {it['id'] for it in new_items}
        new_items.extend(scan_outlook_email(outlook, start_date, seen_so_far))
    except Exception as e:
        print(f"  Outlook unavailable: {e}", flush=True)

    # Load full previous items for merge
    prev_items = []
    if OUTPUT_FILE.exists():
        try:
            soup = BeautifulSoup(OUTPUT_FILE.read_text(encoding='utf-8'), 'html.parser')
            tag = soup.find('script', {'id': 'sa-data'})
            if tag:
                prev_items = json.loads(tag.string).get('items', [])
        except Exception:
            pass

    # Stats
    mine      = [it for it in new_items if not it.get('uncertain')]
    uncertain = [it for it in new_items if it.get('uncertain')]
    all_pids   = {it['pid'] for it in mine if it['pid'] != 'ADMIN'}
    admin_n    = sum(1 for it in mine if it['pid'] == 'ADMIN')

    print(f"\nResults: {len(mine)} owned item(s) — {len(all_pids)} PID(s), {admin_n} Admin", flush=True)
    print(f"         {len(uncertain)} item(s) need confirmation", flush=True)
    print(f"Generating HTML...", flush=True)

    html = generate_html(new_items, start_date, today, prev_items)
    OUTPUT_FILE.write_text(html, encoding='utf-8')

    print(f"\nSaved: {OUTPUT_FILE}", flush=True)
    print(f"PIDs: {', '.join(sorted(all_pids)) or 'none'}", flush=True)

    overdue_n = sum(1 for it in mine if date_status(_parse_date(it['due_date'])) == 'overdue')
    if overdue_n:
        print(f"Overdue: {overdue_n} item(s) — check red items in the report", flush=True)
    if uncertain:
        print(f"Needs confirmation: {len(uncertain)} item(s) in the bottom section", flush=True)


if __name__ == '__main__':
    main()
