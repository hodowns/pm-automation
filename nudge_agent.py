#!/usr/bin/env python3
"""
Nudge Agent — Action Item Follow-Up Drafter
Scans meeting notes and Webex recaps for other people's open action items,
then opens one Outlook draft per owner as a concise reminder.

Usage:  python nudge_agent.py
        python nudge_agent.py --days 30   (override lookback window)
Trigger: "Nudge"
Schedule: Tuesday and Friday at 9:00 AM (Windows Task Scheduler)
"""

import re, json, hashlib, sys
import win32com.client
from datetime import datetime, timedelta, date
from pathlib import Path
from bs4 import BeautifulSoup

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ── Config ─────────────────────────────────────────────────────────────────────

SHAREPOINT_BASE = Path(r"C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS")
OUTPUT_DIR      = Path(r"C:\Users\hodowns\OneDrive - Cisco\Desktop\Nudge Agent")
NUDGE_LOG       = OUTPUT_DIR / "nudge-log.json"
WEBEX_BOT       = "messenger@webex.com"
SKIP_EMAILS     = {                          # never put these in a nudge To field
    'messenger@webex.com',
    'houston.downs16@gmail.com',
    'rtnealy78@gmail.com',               # Reginald Nealy personal — use reginald.nealy@houstontx.gov
}

# Canonical email overrides — when a name resolves to a blocked/wrong address,
# these force the correct one regardless of what the contact cache finds.
EMAIL_OVERRIDES = {
    'reginald nealy': 'reginald.nealy@houstontx.gov',
    'reginald':       'reginald.nealy@houstontx.gov',
    'nealy':          'reginald.nealy@houstontx.gov',
    'chris gillis':   'christopher.gillis@houstontx.gov',
    'christopher gillis': 'christopher.gillis@houstontx.gov',
    'chris':          'christopher.gillis@houstontx.gov',
}
DEFAULT_DAYS    = 8    # lookback window — ~1 week + 1-day buffer for late-filed notes
MAX_DAYS        = 8    # hard cap: never scan further back than this regardless of last_run
RENUDGE_DAYS    = 7    # days before re-nudging the same item

PID_RE = re.compile(r'\b[A-Z]\d{5}\b')

HOUSTON_RE = re.compile(
    r'\b(houston\s+downs?|hodowns|h\.?\s*downs?|houston\b|'
    r'hd\s+(to|will|needs?\s+to|should|can)\b|'
    r'pm\s+(to|will|needs?\s+to)\b|'
    r'i\s+will\b|i\'ll\b|i\s+need\s+to\b|'
    r'i\'m\s+going\s+to\b|i\s+am\s+going\s+to\b|'
    r'i\s+can\s+|i\s+should\b|let\s+me\b|i\s+want\s+to\b)',
    re.IGNORECASE
)

FOOTER_RE = re.compile(
    r'need help\?|help\.webex\.com|cisco and/or its affiliates'
    r'|all rights reserved|privacy statement|terms of service'
    r'|view meeting content|view meeting recap|claude notification incident',
    re.IGNORECASE,
)

# Detect "Name will/to/should/needs to..." at the start of an action item
OWNER_SENTENCE_RE = re.compile(
    r'^([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\s+'
    r'(?:will\b|to\b|should\b|needs?\s+to\b|is\s+going\s+to\b|'
    r'can\b|has\s+to\b|must\b|please\b)',
    re.IGNORECASE
)

SKIP_OWNER_WORDS = {
    'the', 'team', 'all', 'please', 'cisco', 'vendor', 'customer',
    'client', 'everyone', 'attendees', 'participants', 'both', 'each',
    'this', 'that', 'any', 'note', 'action', 'group', 'management',
    'leadership', 'engineering', 'operations', 'network', 'security',
}

# Words that end up capitalized after fix_garbled when they were mid-sentence verbs
# e.g. "AntonioNeedsTo..." → "Antonio Needs To..." — "Needs" is not part of the name
VERB_WORDS = {
    'needs', 'need', 'will', 'should', 'can', 'must', 'has', 'have',
    'is', 'are', 'was', 'were', 'may', 'might', 'shall', 'would',
    'could', 'please', 'to', 'also',
}

ROOM_SKIP_WORDS = {'conference', 'room', 'surface', 'hub', 'board', 'hits', 'snow', 'queue', 'webexboard'}

# Meeting titles matching this pattern are skipped entirely (no nudges generated)
SKIP_MEETING_RE = re.compile(r'personal\s+room', re.IGNORECASE)

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


# ── Nudge log ──────────────────────────────────────────────────────────────────

def load_nudge_log():
    if not NUDGE_LOG.exists():
        return {'last_run': None, 'nudged': {}}
    try:
        data = json.loads(NUDGE_LOG.read_text(encoding='utf-8'))
        last_run = datetime.strptime(data['last_run'], '%Y-%m-%d').date() if data.get('last_run') else None
        return {'last_run': last_run, 'nudged': data.get('nudged', {})}
    except Exception:
        return {'last_run': None, 'nudged': {}}


def save_nudge_log(log):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    NUDGE_LOG.write_text(json.dumps({
        'last_run': date.today().isoformat(),
        'nudged': log['nudged'],
    }, indent=2), encoding='utf-8')


def was_recently_nudged(item_id, log):
    entry = log['nudged'].get(item_id)
    if not entry:
        return False
    try:
        nudged_date = datetime.strptime(entry['nudged_date'], '%Y-%m-%d').date()
        return (date.today() - nudged_date).days < RENUDGE_DAYS
    except Exception:
        return False


# ── Helpers ────────────────────────────────────────────────────────────────────

def make_id(pid, action, owner):
    key = f"{pid}::{owner.lower().strip()}::{action[:80].lower().strip()}"
    return hashlib.md5(key.encode()).hexdigest()[:12]


def is_houston(text):
    return bool(HOUSTON_RE.search(text or ''))


def fix_garbled(text):
    text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
    text = re.sub(r'([A-Z]{2,})([A-Z][a-z])', r'\1 \2', text)
    return text.strip()


def fmt_date(d):
    return d.strftime('%m/%d/%y')


def parse_due_date(raw, fallback_date):
    raw = (raw or '').strip()
    if not raw or raw.upper() in ('TBD', 'N/A', '-', ''):
        return fallback_date + timedelta(days=7)
    for fmt in ('%m/%d/%y', '%m/%d/%Y', '%Y-%m-%d', '%B %d, %Y', '%b %d, %Y'):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return fallback_date + timedelta(days=7)


def detect_owner_from_text(action_text):
    m = OWNER_SENTENCE_RE.match(action_text.strip())
    if not m:
        return None
    candidate = m.group(1).strip()
    words = candidate.split()

    # Strip trailing verb words garbled into the name by fix_garbled
    # e.g. "Antonio Needs to share..." → candidate="Antonio Needs" → trim to "Antonio"
    while len(words) > 1 and words[-1].lower() in VERB_WORDS:
        words = words[:-1]
    candidate = ' '.join(words)

    # Reject if any individual word is a skip word (catches "All Participants", etc.)
    if any(w.lower() in SKIP_OWNER_WORDS for w in words):
        return None

    if is_houston(candidate):
        return None

    # Must be at least 2 characters and look like a real name (starts with capital)
    if len(candidate) < 2 or not candidate[0].isupper():
        return None

    return candidate


def _esc(s):
    return (s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def is_external_meeting(meeting_key):
    return bool(re.search(r'\[external', meeting_key, re.IGNORECASE))


def is_cisco_email(email):
    return bool(email) and '@cisco.com' in email.lower()


# ── Owner email resolution ─────────────────────────────────────────────────────

def build_contact_cache(outlook, days=30, max_msgs=500):
    """Scan recent inbox and sent items to build a name → email lookup.
    Keys: full name (lower), first name, last name, 'first last' (if middle present).
    Ambiguous keys (same short name maps to multiple distinct emails) are dropped
    so a first-name-only match like 'David' never picks the wrong David.
    Skips Exchange-internal /o= addresses and Webex guest accounts.
    Capped at max_msgs per folder to keep startup fast.
    """
    # staging: key -> set of lowercase emails (detect ambiguity)
    # first_seen: key -> email with original case (used in final cache)
    staging   = {}
    first_seen = {}
    ns = outlook.GetNamespace('MAPI')
    cutoff = (date.today() - timedelta(days=days)).strftime('%m/%d/%Y %I:%M %p')

    def _add(name, email):
        if not name or not email or '@' not in email:
            return
        if email.lower().startswith('/o=') or '@guest.webex' in email.lower():
            return
        if email.lower().endswith('@gmail.com'):
            return
        if email.lower() in SKIP_EMAILS:
            return
        name = name.strip()
        if not name:
            return
        parts = name.split()
        keys = {name.lower()}
        if parts:
            keys.add(parts[0].lower())
        if len(parts) >= 2:
            keys.add(parts[-1].lower())
            keys.add(f"{parts[0].lower()} {parts[-1].lower()}")
        for k in keys:
            if not k:
                continue
            staging.setdefault(k, set()).add(email.lower())
            if k not in first_seen:
                first_seen[k] = email

    for folder_id in (6, 5):  # inbox, sent
        try:
            folder = ns.GetDefaultFolder(folder_id)
            items  = folder.Items
            items.Sort('[ReceivedTime]', True)
            try:
                items = items.Restrict(f"[ReceivedTime] >= '{cutoff}'")
            except Exception:
                pass
            count = 0
            for msg in items:
                if count >= max_msgs:
                    break
                try:
                    _add(getattr(msg, 'SenderName', '') or '',
                         getattr(msg, 'SenderEmailAddress', '') or '')
                    for r in msg.Recipients:
                        _add(getattr(r, 'Name', '') or '',
                             getattr(r, 'Address', '') or '')
                    count += 1
                except Exception:
                    continue
        except Exception:
            pass

    # Only keep keys that map to exactly one distinct email — ambiguous short names
    # (e.g. "david" → two different Davids) are dropped to prevent wrong-person matches.
    cache = {k: first_seen[k] for k, emails in staging.items() if len(emails) == 1}
    ambiguous = len(staging) - len(cache)
    print(f"  Contact cache: {len(cache)} name→email mapping(s) built "
          f"({ambiguous} ambiguous key(s) dropped)", flush=True)
    return cache


def build_project_contact_cache():
    """Load name→email from Project Contact List.xlsx in every PID folder.
    Returns {pid: {name_lower: email}} — same ambiguity filter as build_contact_cache.
    Strips parenthetical usernames like '(dabeynon)' from display names.
    Skips conference rooms, equipment, and service queues.
    """
    if not HAS_OPENPYXL:
        print("  Project contacts: openpyxl not installed — skipping", flush=True)
        return {}

    pid_contacts = {}

    for state_dir in sorted(SHAREPOINT_BASE.iterdir()):
        if not state_dir.is_dir():
            continue
        for proj_dir in sorted(state_dir.iterdir()):
            if not proj_dir.is_dir():
                continue
            pid_m = PID_RE.search(proj_dir.name)
            if not pid_m:
                continue
            pid = pid_m.group(0)

            contact_xlsx = None
            for subdir in proj_dir.iterdir():
                if subdir.is_dir() and 'customer profile' in subdir.name.lower():
                    candidate = subdir / 'Project Contact List.xlsx'
                    if candidate.exists():
                        contact_xlsx = candidate
                        break
            if not contact_xlsx:
                continue

            try:
                wb = openpyxl.load_workbook(contact_xlsx, read_only=True, data_only=True)
                ws = wb.active
                staging    = {}
                first_seen = {}
                for row in ws.iter_rows(values_only=True):
                    email = str(row[0] or '').strip()
                    name  = str(row[1] or '').strip()
                    if not email or '@' not in email or not name:
                        continue
                    if email.lower().endswith('@gmail.com'):
                        continue
                    if email.lower() in SKIP_EMAILS:
                        continue
                    name = re.sub(r'\s*\([^)]*\)\s*$', '', name).strip()
                    if not name:
                        continue
                    words_lower = {w.lower() for w in name.split()}
                    if words_lower & ROOM_SKIP_WORDS:
                        continue
                    parts = name.split()
                    keys  = {name.lower()}
                    if parts:
                        keys.add(parts[0].lower())
                    if len(parts) >= 2:
                        keys.add(parts[-1].lower())
                        keys.add(f"{parts[0].lower()} {parts[-1].lower()}")
                    for k in keys:
                        if not k:
                            continue
                        staging.setdefault(k, set()).add(email.lower())
                        if k not in first_seen:
                            first_seen[k] = email
                wb.close()
                pid_contacts[pid] = {k: first_seen[k]
                                     for k, em in staging.items() if len(em) == 1}
            except Exception:
                continue

    total = sum(len(v) for v in pid_contacts.values())
    print(f"  Project contacts: {total} name→email mapping(s) across {len(pid_contacts)} PID(s)", flush=True)
    return pid_contacts


def resolve_email(outlook, display_name, cache=None):
    """Resolve a display name to an SMTP email.
    0. Check EMAIL_OVERRIDES for canonical corrections.
    1. Check the contact cache (handles first-name-only and partial matches).
    2. Fall back to Outlook GAL resolve (works for exact Exchange display names).
    """
    if not display_name:
        return None

    # 0. Canonical overrides (highest priority)
    override = EMAIL_OVERRIDES.get(display_name.lower().strip())
    if override and not override.lower().endswith('@gmail.com'):
        return override

    # 1. Cache lookup — try full name, then first name only
    if cache:
        key = display_name.lower().strip()
        if key in cache:
            return cache[key]
        first = key.split()[0] if key.split() else ''
        if first and len(first) > 1 and first in cache:
            return cache[first]

    # 2. Outlook GAL resolve
    try:
        recip = outlook.CreateRecipient(display_name)
        if recip.Resolve():
            ae = recip.AddressEntry
            if ae.AddressEntryUserType == 0:
                eu = ae.GetExchangeUser()
                if eu:
                    return eu.PrimarySmtpAddress
            try:
                props = ae.PropertyAccessor
                return props.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
            except Exception:
                pass
    except Exception:
        pass

    return None


def resolve_owner_emails(outlook, owner_name, cache):
    """Handle multi-owner strings like 'Kennie / Salim' — resolve each part."""
    parts = [p.strip() for p in re.split(r'\s*/\s*', owner_name) if p.strip()]
    emails = []
    for part in parts:
        email = resolve_email(outlook, part, cache)
        if email and email not in emails:
            emails.append(email)
    return emails


def resolve_owner_with_context(outlook, owner_name, global_cache, pid_contacts,
                               meeting_attendees, pid):
    """Extended resolution with three fallback layers.

    1. Global Outlook cache + GAL (ambiguity-filtered)
    2. Project Contact List for this PID
    3. Meeting attendees disambiguation — if owner is first-name-only and exactly
       one attendee shares that first name, resolve the full attendee name via
       layers 1 & 2

    Returns (emails: list[str], resolved: bool)
    resolved=False means nothing found; caller should fall back to display name.
    """
    emails = [e for e in resolve_owner_emails(outlook, owner_name, global_cache)
              if e.lower() not in SKIP_EMAILS]
    if emails:
        return emails, True

    # Layer 2: project contact list
    proj = pid_contacts.get(pid, {})
    if proj:
        layer2 = []
        for part in [p.strip() for p in re.split(r'\s*/\s*', owner_name) if p.strip()]:
            key = part.lower()
            if key in proj and proj[key].lower() not in SKIP_EMAILS:
                e = proj[key]
                if e not in layer2:
                    layer2.append(e)
        if layer2:
            return layer2, True

    # Layer 3: meeting attendees disambiguation (first-name-only owners only)
    layer3 = []
    for part in [p.strip() for p in re.split(r'\s*/\s*', owner_name) if p.strip()]:
        if ' ' in part:
            continue   # already a full name — layers 1 & 2 already tried it
        first      = part.lower()
        candidates = [a for a in meeting_attendees
                      if a.split() and a.split()[0].lower() == first]
        if len(candidates) != 1:
            continue   # zero or ambiguous — skip
        full_name = candidates[0]
        resolved  = [e for e in resolve_owner_emails(outlook, full_name, global_cache)
                     if e.lower() not in SKIP_EMAILS]
        if not resolved and proj:
            p  = full_name.split()
            fl = f"{p[0].lower()} {p[-1].lower()}" if len(p) >= 2 else full_name.lower()
            for k in (full_name.lower(), fl):
                if k in proj and proj[k].lower() not in SKIP_EMAILS:
                    resolved = [proj[k]]
                    break
        for e in resolved:
            if e not in layer3:
                layer3.append(e)

    if layer3:
        return layer3, True

    return [], False


# ── Source: .md meeting notes ──────────────────────────────────────────────────

def scan_md_files(start_date, log):
    items = []
    for state_dir in sorted(SHAREPOINT_BASE.iterdir()):
        if not state_dir.is_dir():
            continue
        for proj_dir in sorted(state_dir.iterdir()):
            if not proj_dir.is_dir():
                continue
            pid_m = PID_RE.search(proj_dir.name)
            pid   = pid_m.group(0) if pid_m else None
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
                items.extend(_parse_md_for_others(md_file, pid, file_date, log))

    print(f"  .md files: {len(items)} other-owner item(s)", flush=True)
    return items


def _parse_md_for_others(path, pid, file_date, log):
    items = []
    try:
        text = path.read_text(encoding='utf-8', errors='ignore')
    except Exception:
        return items

    title_m   = re.search(r'^#\s+(.+)$', text, re.MULTILINE)
    raw_title = title_m.group(1).strip() if title_m else path.stem
    if SKIP_MEETING_RE.search(raw_title):
        return items
    is_external = bool(re.search(r'\[external', raw_title, re.IGNORECASE))
    short_title = re.sub(r'\s*\[(?:EXTERNAL|INTERNAL|EXT|INT)\]\s*', ' ', raw_title, flags=re.IGNORECASE).strip()
    short_title = re.sub(r'\s+', ' ', short_title)
    # Normalise key: collapse spaces around punctuation, strip trailing PID, truncate
    # at 45 chars (same as Webex email scanner) so both sources produce identical keys.
    norm_title = re.sub(r'\s*:\s*', ': ', short_title).strip()
    norm_title = re.sub(r'\s+[A-Z]\d{5}\s*$', '', norm_title).strip()
    norm_title = norm_title.rstrip(':').strip()[:45]
    source_label = f"{norm_title}, {file_date.strftime('%b %d')}"

    # Extract attendees for first-name disambiguation
    attendees_m = re.search(r'\*\*Attendees:\*\*\s*(.+)', text)
    meeting_attendees = (
        [a.strip() for a in attendees_m.group(1).split(',') if a.strip()]
        if attendees_m else []
    )

    if not pid:
        fm_m = re.search(r'^project_id:\s*(\S+)', text, re.MULTILINE)
        pid = fm_m.group(1) if fm_m else None
    if not pid:
        pid_m = PID_RE.search(raw_title)
        pid   = pid_m.group(0) if pid_m else None

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

        if len(cols) == 4:
            _, action, owner, due_raw = cols
        elif len(cols) == 3:
            _, action, owner = cols
            due_raw = ''
        else:
            action  = cols[1] if len(cols) > 1 else cols[0]
            owner   = cols[2] if len(cols) > 2 else ''
            due_raw = cols[3] if len(cols) > 3 else ''

        action = action.strip()
        owner  = owner.strip()

        if not action or action.startswith('-') or FOOTER_RE.search(action):
            continue
        if not owner or is_houston(owner) or is_houston(action):
            continue
        if any(w.lower() in SKIP_OWNER_WORDS for w in owner.split()):
            continue

        due_date = parse_due_date(due_raw, file_date)
        item_id  = make_id(pid or 'NONE', action, owner)

        if was_recently_nudged(item_id, log):
            continue

        items.append({
            'id':               item_id,
            'pid':              pid or 'ADMIN',
            'action':           action,
            'owner':            owner,
            'due':              fmt_date(due_date),
            'source':           source_label,
            'source_type':      'meeting-notes',
            'meeting_key':      source_label,
            'meeting_title':    norm_title,
            'meeting_date':     file_date.isoformat(),
            'meeting_attendees': meeting_attendees,
            'is_external':      is_external,
        })

    return items


# ── Source: Webex recap emails ─────────────────────────────────────────────────

def scan_webex_emails(outlook, start_date, log):
    items      = []
    ns         = outlook.GetNamespace('MAPI')
    inbox      = ns.GetDefaultFolder(6)
    mail_items = inbox.Items
    mail_items.Sort('[ReceivedTime]', True)
    cutoff     = start_date.strftime('%m/%d/%Y %I:%M %p')
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

            body  = (getattr(msg, 'Body', '') or '').replace('\r\n', '\n').replace('\r', '\n')
            lines = [l.strip() for l in body.split('\n') if l.strip()]

            meeting_date = None
            for line in lines:
                if re.match(r'(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),', line):
                    try:
                        meeting_date = datetime.strptime(line.strip(), '%A, %B %d, %Y').date()
                    except ValueError:
                        pass
                    break
            if not meeting_date or meeting_date < start_date:
                continue

            pids_found  = list(dict.fromkeys(PID_RE.findall(subj + ' ' + body)))
            primary_pid = pids_found[0] if pids_found else None

            title_raw = ''
            for i, line in enumerate(lines):
                if 'meeting content is available' in line.lower() and i + 1 < len(lines):
                    title_raw = lines[i + 1]
                    break
            if SKIP_MEETING_RE.search(title_raw):
                continue
            is_external  = bool(re.search(r'\[external', title_raw, re.IGNORECASE))
            short_title  = re.sub(r'\[(?:EXTERNAL|INTERNAL|EXT|INT)\]\s*', '', title_raw, flags=re.I).strip()
            norm_title   = re.sub(r'\s*:\s*', ': ', short_title).strip()
            norm_title   = re.sub(r'\s+[A-Z]\d{5}\s*$', '', norm_title).strip()
            norm_title   = norm_title.rstrip(':').strip()[:45]
            source_label = f"{norm_title or 'Meeting'}, {meeting_date.strftime('%b %d')}"

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
                    if is_houston(clean):
                        continue

                    owner = detect_owner_from_text(clean)
                    if not owner:
                        continue  # can't attribute — skip to avoid false nudges

                    pid     = primary_pid or 'ADMIN'
                    due     = fmt_date(meeting_date + timedelta(days=7))
                    item_id = make_id(pid, clean, owner)

                    if was_recently_nudged(item_id, log):
                        continue

                    items.append({
                        'id':           item_id,
                        'pid':          pid,
                        'action':       clean,
                        'owner':        owner,
                        'due':          due,
                        'source':       source_label,
                        'source_type':  'webex-email',
                        'meeting_key':  source_label,
                        'meeting_title': norm_title or 'Meeting',
                        'meeting_date': meeting_date.isoformat(),
                        'is_external':  is_external,
                    })

            processed += 1
        except Exception:
            continue

    print(f"  Webex emails: {processed} recap(s) scanned, {len(items)} other-owner item(s)", flush=True)
    return items


# ── Build nudge email HTML ─────────────────────────────────────────────────────

def build_nudge_html(meeting_title, items):
    HDR = '#1565A0'

    rows = ''
    for i, it in enumerate(items):
        pid_label = PID_NAMES.get(it['pid'], it['pid']) if it['pid'] != 'ADMIN' else ''
        pid_part  = (
            f" <span style='color:#888;font-size:10pt;'>({_esc(pid_label)})</span>"
            if pid_label else ''
        )
        rows += (
            f"<tr>"
            f"<td style='padding:5px 14px;font-size:11pt;color:#444444;'>{_esc(str(i + 1))}</td>"
            f"<td style='padding:5px 14px 5px 0;font-size:11pt;font-weight:600;color:#1f1f1f;'>{_esc(it['action'])}{pid_part}</td>"
            f"<td style='padding:5px 14px 5px 0;font-size:11pt;color:#444444;white-space:nowrap;'>{_esc(it['owner'])}</td>"
            f"<td style='padding:5px 14px 5px 0;font-size:11pt;color:#444444;white-space:nowrap;'>{_esc(it['due'])}</td>"
            f"</tr>\n"
        )

    today_label = date.today().strftime('%B %d, %Y')

    table = (
        f"<table cellpadding='5' cellspacing='0' style='border-collapse:collapse;width:auto;margin:14px 0;'>"
        f"<thead><tr style='background-color:{HDR};color:#ffffff;'>"
        f"<th style='padding:6px 14px;font-size:11pt;font-weight:700;text-align:left;'>#</th>"
        f"<th style='padding:6px 14px 6px 0;font-size:11pt;font-weight:700;text-align:left;'>Action</th>"
        f"<th style='padding:6px 14px 6px 0;font-size:11pt;font-weight:700;text-align:left;white-space:nowrap;'>Owner</th>"
        f"<th style='padding:6px 14px 6px 0;font-size:11pt;font-weight:700;text-align:left;white-space:nowrap;'>Due</th>"
        f"</tr></thead>"
        f"<tbody>{rows}</tbody></table>"
    )

    return (
        "<div style=\"font-family:'Aptos',Calibri,Arial,sans-serif;font-size:12pt;"
        "color:#1f1f1f;max-width:720px;line-height:1.5;\">"
        "<p>Hi all,</p>"
        f"<p>Quick follow-up on the open action items from <strong>{_esc(meeting_title)}</strong>"
        f" &mdash; as of {today_label}:</p>"
        f"{table}"
        "<p style='margin-top:14px;'>Please reply if any of these are already complete "
        "or if you need anything to move them forward.</p>"
        "<p>Thanks,<br>Houston</p>"
        "</div>"
    )


# ── Open one Outlook draft per meeting ────────────────────────────────────────

def open_nudge_draft(outlook, meeting_key, items, all_emails, audience=None):
    """audience: None (single draft), 'cisco' (internal split), 'customer' (external split)"""
    meeting_title = items[0].get('meeting_title', meeting_key)
    meeting_date  = items[0].get('meeting_date', date.today().isoformat())
    try:
        date_str = datetime.strptime(meeting_date, '%Y-%m-%d').strftime('%m-%d-%Y')
    except Exception:
        date_str = meeting_date

    pids    = sorted({it['pid'] for it in items if it.get('pid') and it['pid'] != 'ADMIN'})
    pid_str = f" [{', '.join(pids)}]" if pids else ''
    if audience == 'cisco':
        subject = f"Action Items (Internal): {meeting_title} — {date_str}{pid_str}"
    else:
        subject = f"Action Items: {meeting_title} — {date_str}{pid_str}"

    inner_html = build_nudge_html(meeting_title, items)

    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Display(False)

    existing = mail.HTMLBody or ''
    body_tag = re.search(r'(<body[^>]*>)', existing, re.IGNORECASE)
    if body_tag:
        ins = body_tag.end()
        mail.HTMLBody = existing[:ins] + inner_html + existing[ins:]
    else:
        mail.HTMLBody = (
            "<html><body style=\"font-family:'Aptos',Calibri,Arial,sans-serif;"
            "font-size:12pt;color:#1f1f1f;max-width:720px;line-height:1.5;\">"
            + inner_html + "</body></html>"
        )

    for addr in all_emails:
        recip = mail.Recipients.Add(addr)
        recip.Type = 1
        try:
            recip.Resolve()   # resolve each recipient independently; failures leave the name for manual fix
        except Exception:
            pass

    addr_label = ', '.join(all_emails) if all_emails else '(no recipients resolved — add manually)'
    print(f"  Draft: {meeting_key} → {len(all_emails)} recipient(s), {len(items)} item(s) [{addr_label}]", flush=True)


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    days = DEFAULT_DAYS
    if '--days' in sys.argv:
        try:
            days = int(sys.argv[sys.argv.index('--days') + 1])
        except (IndexError, ValueError):
            pass

    meeting_filter = None
    if '--meeting' in sys.argv:
        try:
            meeting_filter = sys.argv[sys.argv.index('--meeting') + 1].lower()
        except IndexError:
            pass

    force = '--force' in sys.argv   # bypass 7-day re-nudge guard (for testing)

    today = date.today()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    log = load_nudge_log()
    print(f"Nudge Agent — {today.strftime('%A, %B %d, %Y')}", flush=True)
    if log['last_run'] and '--days' not in sys.argv:
        start_date = log['last_run']
        print(f"Scanning since last run: {log['last_run']}", flush=True)
    else:
        start_date = today - timedelta(days=days)
        print(f"Scanning back {days} days (from {start_date})", flush=True)

    # Hard cap — never go back more than MAX_DAYS regardless of last_run date
    earliest = today - timedelta(days=MAX_DAYS)
    if start_date < earliest:
        print(f"  (capped to {MAX_DAYS}-day max window → {earliest})", flush=True)
        start_date = earliest

    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
    except Exception as e:
        print(f"Cannot connect to Outlook: {e}", flush=True)
        return

    print("\nBuilding contact cache...", flush=True)
    cache         = build_contact_cache(outlook)
    pid_contacts  = build_project_contact_cache()

    if force:
        log['nudged'] = {}   # treat everything as un-nudged for this run
        print("(--force: re-nudge guard bypassed)", flush=True)

    print("\nScanning sources...", flush=True)
    all_items = []
    all_items.extend(scan_md_files(start_date, log))
    all_items.extend(scan_webex_emails(outlook, start_date, log))

    if not all_items:
        print("\nNo nudgeable action items found. Nothing to do.", flush=True)
        save_nudge_log(log)
        return

    # Keep only the most recent occurrence of each meeting series.
    # Series key = title prefix before the first colon (strips topic suffixes like
    # "ACI Office Hours: Change planning" → "ACI Office Hours"), so recurring
    # meetings with per-session subtopics collapse correctly.
    # meeting_date is YYYY-MM-DD so string comparison gives correct ordering.
    def series_key(it):
        title = it.get('meeting_title') or it.get('source', '')
        return title.split(':')[0].strip()

    series_latest = {}
    for it in all_items:
        key   = series_key(it)
        mdate = it.get('meeting_date', '')
        if key not in series_latest or mdate > series_latest[key]:
            series_latest[key] = mdate

    before = len(all_items)
    all_items = [
        it for it in all_items
        if it.get('meeting_date', '') == series_latest.get(series_key(it), '')
    ]
    print(f"  Latest-occurrence filter: {before} → {len(all_items)} item(s) "
          f"({len(series_latest)} series, keeping most recent date each)", flush=True)

    # Group by meeting
    by_meeting = {}
    for it in all_items:
        key = it.get('meeting_key', it['source'])
        by_meeting.setdefault(key, []).append(it)

    # Apply --meeting filter if specified
    if meeting_filter:
        by_meeting = {k: v for k, v in by_meeting.items() if meeting_filter in k.lower()}
        if not by_meeting:
            print(f"\nNo meetings matched filter '{meeting_filter}'. Nothing to do.", flush=True)
            return
        print(f"\n(Filtered to meetings matching '{meeting_filter}')", flush=True)

    print(f"\nFound {len(all_items)} item(s) across {len(by_meeting)} meeting(s):", flush=True)
    for meeting_key, its in sorted(by_meeting.items()):
        owners = sorted({it['owner'].strip().title() for it in its})
        print(f"  {meeting_key}: {len(its)} item(s), owners: {', '.join(owners)}", flush=True)

    print("\nOpening drafts...", flush=True)
    today_iso  = today.isoformat()
    unresolved = []   # (owner_display, meeting_key) for end-of-run warning
    drafts_opened = 0

    for meeting_key, items in sorted(by_meeting.items()):
        external = any(it.get('is_external') for it in items)

        # Collect meeting attendees for first-name disambiguation
        meeting_attendees = []
        for it in items:
            for a in it.get('meeting_attendees', []):
                if a not in meeting_attendees:
                    meeting_attendees.append(a)

        # Resolve every unique owner once; record emails and cisco/customer classification
        seen_owners   = set()
        owner_emails  = {}   # norm_owner -> list[email]
        owner_is_cisco = {}  # norm_owner -> bool (True = all cisco, False = any customer)

        for it in items:
            norm = it['owner'].strip().title()
            if norm in seen_owners:
                continue
            seen_owners.add(norm)

            emails, resolved = resolve_owner_with_context(
                outlook, norm, cache, pid_contacts, meeting_attendees, it['pid']
            )
            if resolved:
                owner_emails[norm]   = emails
                owner_is_cisco[norm] = all(is_cisco_email(e) for e in emails)
            else:
                owner_emails[norm]   = [norm]   # display name fallback
                owner_is_cisco[norm] = False     # unknown → customer side
                unresolved.append((norm, meeting_key))

        if not external:
            # Internal meeting — single draft, all recipients together
            all_recipients = []
            for norm, emails in owner_emails.items():
                for e in emails:
                    if e not in all_recipients:
                        all_recipients.append(e)
            open_nudge_draft(outlook, meeting_key, items, all_recipients)
            drafts_opened += 1
        else:
            # External meeting — split into Cisco draft and customer draft
            cisco_items, cisco_recipients       = [], []
            customer_items, customer_recipients = [], []

            for it in items:
                norm    = it['owner'].strip().title()
                emails  = owner_emails.get(norm, [norm])
                cisco_e = [e for e in emails if is_cisco_email(e)]
                cust_e  = [e for e in emails if not is_cisco_email(e)]

                if cisco_e:
                    cisco_items.append(it)
                    for e in cisco_e:
                        if e not in cisco_recipients:
                            cisco_recipients.append(e)
                if cust_e:
                    customer_items.append(it)
                    for e in cust_e:
                        if e not in customer_recipients:
                            customer_recipients.append(e)
                # Item with no resolved emails at all → customer draft with display name
                if not cisco_e and not cust_e:
                    customer_items.append(it)
                    if norm not in customer_recipients:
                        customer_recipients.append(norm)

            if cisco_items:
                open_nudge_draft(outlook, meeting_key, cisco_items,
                                 cisco_recipients, audience='cisco')
                drafts_opened += 1
            if customer_items:
                open_nudge_draft(outlook, meeting_key, customer_items,
                                 customer_recipients, audience='customer')
                drafts_opened += 1

        for it in items:
            log['nudged'][it['id']] = {
                'meeting':     meeting_key,
                'owner':       it['owner'],
                'nudged_date': today_iso,
                'action':      it['action'],
                'source':      it['source'],
                'pid':         it['pid'],
            }

    save_nudge_log(log)
    print(f"\nDone. {drafts_opened} draft(s) opened in Outlook. Review and send when ready.", flush=True)
    print(f"Log saved: {NUDGE_LOG}", flush=True)

    if unresolved:
        print(f"\n⚠  Could not resolve email for {len(unresolved)} owner(s) — please add manually:", flush=True)
        for owner, mtg in unresolved:
            print(f"   • {owner}  ({mtg})", flush=True)


if __name__ == '__main__':
    main()
