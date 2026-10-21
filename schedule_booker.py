#!/usr/bin/env python3
"""
Schedule Booker — Scans inbox for replies to Schedule Request emails,
detects the confirmed time option, checks if already on calendar,
and opens a pre-filled calendar invite draft for review.

Usage:  python schedule_booker.py
Trigger: "Schedule Booker"
"""

import re, json, hashlib, sys, argparse
import win32com.client
from datetime import datetime, timedelta, date
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
from pathlib import Path
from bs4 import BeautifulSoup

DEFAULT_TZ = 'America/Chicago'

TZ_ALIASES = {
    'ct': 'America/Chicago', 'cst': 'America/Chicago', 'cdt': 'America/Chicago',
    'central': 'America/Chicago',
    'et': 'America/New_York', 'est': 'America/New_York', 'edt': 'America/New_York',
    'eastern': 'America/New_York',
    'pt': 'America/Los_Angeles', 'pst': 'America/Los_Angeles', 'pdt': 'America/Los_Angeles',
    'pacific': 'America/Los_Angeles',
    'mt': 'America/Denver', 'mst': 'America/Denver', 'mdt': 'America/Denver',
    'mountain': 'America/Denver',
}

def resolve_tz(name):
    """Return a ZoneInfo object; raise ValueError with a clear message if unrecognised."""
    key = name.strip().lower()
    iana = TZ_ALIASES.get(key, name.strip())
    try:
        return ZoneInfo(iana), iana
    except ZoneInfoNotFoundError:
        raise ValueError(
            f"Unknown timezone '{name}'. Use an IANA name (e.g. 'America/Chicago') "
            f"or a short alias: central, eastern, pacific, mountain."
        )

REQUESTS_LOG  = Path(r"C:\Users\hodowns\OneDrive - Cisco\Desktop\Schedule Requester\pending-requests.json")
SCAN_DAYS     = 60   # how far back to scan inbox for replies

WEEKDAY_NAMES = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
NTH_NAMES     = ['','1st','2nd','3rd','4th','5th']

# Outlook recurrence type constants
OL_RECURS_MONTHLY_NTH = 5

# Outlook day-of-week bitmask
DOW_MASK = {0: 2, 1: 4, 2: 8, 3: 16, 4: 32}   # Mon–Fri

# Outlook recipient type
OL_REQUIRED = 1
OL_OPTIONAL = 2


# ── Pending-requests log ───────────────────────────────────────────────────────

def load_log():
    if not REQUESTS_LOG.exists():
        return []
    try:
        return json.loads(REQUESTS_LOG.read_text(encoding='utf-8'))
    except Exception:
        return []

def save_log(records):
    REQUESTS_LOG.parent.mkdir(parents=True, exist_ok=True)
    REQUESTS_LOG.write_text(json.dumps(records, indent=2), encoding='utf-8')


# ── Inbox scan ─────────────────────────────────────────────────────────────────

def find_reply_emails(ns):
    """Return all inbox emails whose subject starts with 'Re: Schedule Request:'."""
    inbox  = ns.GetDefaultFolder(6)
    items  = inbox.Items
    items.Sort('[ReceivedTime]', True)
    cutoff = (date.today() - timedelta(days=SCAN_DAYS)).strftime('%m/%d/%Y %I:%M %p')
    try:
        items = items.Restrict(f"[ReceivedTime] >= '{cutoff}'")
    except Exception:
        pass

    replies = []
    for msg in items:
        try:
            subj = (getattr(msg, 'Subject', '') or '').strip()
            if re.match(r'^Re:\s+Schedule Request:', subj, re.IGNORECASE):
                replies.append(msg)
        except Exception:
            continue
    return replies


# ── Proposal data recovery ─────────────────────────────────────────────────────

def original_subject(reply_subject):
    return re.sub(r'^Re:\s+', '', reply_subject, flags=re.IGNORECASE).strip()


def recover_data_from_log(orig_subj, log_records):
    """Find the matching pending record in the local log."""
    for rec in log_records:
        if rec.get('subject', '').strip() == orig_subj and rec.get('status') == 'pending':
            return rec
    return None


def recover_data_from_sent(ns, orig_subj):
    """Find the original sent email and extract embedded sr-data JSON comment."""
    sent  = ns.GetDefaultFolder(5)
    items = sent.Items
    items.Sort('[ReceivedTime]', True)
    cutoff = (date.today() - timedelta(days=SCAN_DAYS)).strftime('%m/%d/%Y %I:%M %p')
    try:
        items = items.Restrict(f"[ReceivedTime] >= '{cutoff}'")
    except Exception:
        pass

    for msg in items:
        try:
            if (getattr(msg, 'Subject', '') or '').strip() == orig_subj:
                html = getattr(msg, 'HTMLBody', '') or ''
                m = re.search(r'<!--\s*sr-data:\s*(\{.*?\})\s*-->', html,
                              re.DOTALL | re.IGNORECASE)
                if m:
                    return json.loads(m.group(1))
        except Exception:
            continue
    return None


# ── Option detection from reply body ──────────────────────────────────────────

def detect_option(reply_msg, options):
    """
    Parse reply body to identify which option the customer selected.
    Returns the matching option dict, or None if ambiguous.
    """
    raw  = (getattr(reply_msg, 'Body', '') or '').replace('\r\n', '\n')
    html = getattr(reply_msg, 'HTMLBody', '') or ''

    # Use plain text; strip quoted original (lines starting with ">")
    lines = [l for l in raw.split('\n')
             if not l.strip().startswith('>') and l.strip()]
    text  = ' '.join(lines[:30]).lower()   # first ~30 lines of reply content

    # 1. Explicit "Option A/B/C" or "option a"
    m = re.search(r'\boption\s+([a-e])\b', text)
    if m:
        label = m.group(1).upper()
        for opt in options:
            if opt.get('label') == label:
                return opt, 'label'

    # 2. Weekday name match
    for wday_idx, wday_name in enumerate(WEEKDAY_NAMES[:5]):
        if wday_name.lower() in text:
            for opt in options:
                if opt.get('wday') == wday_idx or (
                    opt.get('type') == 'one-time' and
                    datetime.fromisoformat(opt['start']).weekday() == wday_idx
                ):
                    return opt, 'weekday'

    # 3. Time mention (e.g. "10:00", "9:30")
    time_matches = re.findall(r'\b(\d{1,2}):(\d{2})\s*(?:am|pm)?\b', text)
    for h_str, m_str in time_matches:
        h, mn = int(h_str), int(m_str)
        for opt in options:
            opt_h = opt.get('hour', -1)
            opt_m = opt.get('minute', -1)
            if opt.get('type') == 'one-time':
                start = datetime.fromisoformat(opt['start'])
                opt_h, opt_m = start.hour, start.minute
            if abs(opt_h - h) <= 1 and opt_m == mn:
                return opt, 'time'

    # 4. "Works for us" / "that works" / "sounds good" — if only one option, pick it
    positive_signals = re.search(
        r'\b(works|perfect|sounds good|confirmed|confirm|let\'s go|that one|'
        r'works for us|works for me|good for us|good for me|yes|agree)\b', text)
    if positive_signals and len(options) == 1:
        return options[0], 'implicit'

    return None, None


# ── Already-scheduled check ────────────────────────────────────────────────────

def already_on_calendar(ns, meeting_name, pid):
    """Return True if an Outlook calendar event already exists for this series."""
    cal   = ns.GetDefaultFolder(9)
    items = cal.Items
    items.IncludeRecurrences = False   # look at masters only
    search_terms = [meeting_name.lower(), pid.lower()]
    try:
        for item in items:
            subj = (getattr(item, 'Subject', '') or '').lower()
            if all(t in subj for t in search_terms):
                return True
    except Exception:
        pass
    return False


# ── Calendar invite builder ────────────────────────────────────────────────────

def _to_outlook_dt(aware_dt):
    """Convert a tz-aware datetime to the form Outlook COM's appt.Start expects.

    Outlook on this machine stores and displays appointment times as raw values
    with no UTC conversion (existing calendar items read back as HH:MM+00:00 where
    HH:MM is the local display time).  win32com adds the machine's UTC offset when
    given a naive datetime (naive 10:30 ET → stored 14:30 → displays 2:30 PM).
    The fix: convert to local display time, then wrap with tzinfo=UTC so win32com
    passes it through unchanged, and Outlook stores exactly the display time we want.
    """
    from datetime import timezone as _utc
    local_naive = aware_dt.astimezone().replace(tzinfo=None)
    return local_naive.replace(tzinfo=_utc.utc)


def create_invite_draft(outlook, ns, data, opt, tz_zone, body=None):
    """Open a draft calendar invite (recurring or one-time) for review."""
    meeting  = data.get('meeting', 'Meeting')
    pid      = data.get('pid', '')
    duration = data.get('duration', 45)
    to_list  = data.get('to', [])
    cc_list  = data.get('cc', [])
    subject  = f"{meeting} [EXTERNAL] {pid}"

    appt = outlook.CreateItem(1)   # olAppointmentItem
    appt.Subject  = subject
    appt.Duration = duration
    appt.BusyStatus = 2            # olBusy
    if body:
        appt.Body = body

    if opt.get('type') == 'monthly':
        n    = opt['n']
        wday = opt['wday']
        h    = opt['hour']
        mn   = opt['minute']
        # Build aware time in the proposed TZ, then convert to local naive for COM
        kickoff_aware = datetime.strptime(opt['kickoff'], '%Y-%m-%d').replace(
            hour=h, minute=mn, tzinfo=tz_zone)
        kickoff_local = _to_outlook_dt(kickoff_aware)

        appt.Start = kickoff_local

        pattern = appt.GetRecurrencePattern()
        pattern.RecurrenceType   = OL_RECURS_MONTHLY_NTH
        pattern.DayOfWeekMask    = DOW_MASK[wday]
        pattern.Instance         = n
        pattern.PatternStartDate = kickoff_local

        nth_label  = NTH_NAMES[n]
        wday_label = WEEKDAY_NAMES[wday]
        recur_desc = f"{nth_label} {wday_label} of each month"
        print(f"  Recurrence: {recur_desc} @ "
              f"{kickoff_aware.strftime('%I:%M %p').lstrip('0')} {kickoff_aware.strftime('%Z')} "
              f"→ {kickoff_local.strftime('%I:%M %p').lstrip('0')} local")

    else:   # one-time
        # Build aware time in the proposed TZ, then convert to local naive for COM
        start_aware = datetime.fromisoformat(opt['start']).replace(tzinfo=tz_zone)
        start_local = _to_outlook_dt(start_aware)
        appt.Start  = start_local
        print(f"  One-time: {start_aware.strftime('%A, %B ') + str(start_aware.day)} "
              f"@ {start_aware.strftime('%I:%M %p').lstrip('0')} {start_aware.strftime('%Z')} "
              f"→ {start_local.strftime('%I:%M %p').lstrip('0')} local")

    # Add attendees — everyone (customer contacts and Cisco SMEs) as Required
    all_attendees = list(dict.fromkeys(to_list + cc_list))  # deduplicate, preserve order
    for email in all_attendees:
        r = appt.Recipients.Add(email)
        r.Type = OL_REQUIRED
    if all_attendees:
        try:
            appt.Recipients.ResolveAll()
        except Exception:
            pass

    appt.Display(False)   # Open as draft — do NOT send automatically
    return appt


# ── Main ───────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument('--manual',  action='store_true')
    p.add_argument('--subject', default=None)
    p.add_argument('--option',  default=None)
    p.add_argument('--tz', default=DEFAULT_TZ,
                   help='Timezone for the calendar invite (default: America/Chicago). '
                        'Accepts IANA names or aliases: central, eastern, pacific, mountain.')
    p.add_argument('--body', default=None,
                   help='Optional meeting description / goal text for the invite body.')
    return p.parse_args()


def main():
    args    = parse_args()
    outlook = win32com.client.Dispatch('Outlook.Application')
    ns      = outlook.GetNamespace('MAPI')

    # Resolve CLI timezone now so any bad value fails fast and visibly
    try:
        cli_tz_zone, cli_tz_name = resolve_tz(args.tz)
    except ValueError as e:
        print(f"ERROR: {e}", flush=True)
        sys.exit(1)

    print(f"Schedule Booker — {date.today().strftime('%A, %B %d, %Y')}", flush=True)
    print(f"Timezone: {cli_tz_name}", flush=True)
    print(f"Scanning inbox for Schedule Request replies (last {SCAN_DAYS} days)...\n", flush=True)

    # ── Manual override ──────────────────────────────────────────────────────
    if args.manual:
        if not args.subject or not args.option:
            print("--manual requires --subject and --option.")
            return
        orig_subj = args.subject.removeprefix('Re: ').removeprefix('RE: ').strip()
        log_records = load_log()
        data = recover_data_from_log(orig_subj, log_records)
        if not data:
            data = recover_data_from_sent(ns, orig_subj)
        if not data:
            print(f"Could not recover proposal data for: {orig_subj}")
            print("Please check pending-requests.json or the original sent email.")
            return
        options = data.get('options', [])
        selected = next((o for o in options if o.get('label','').upper() == args.option.upper()), None)
        if not selected:
            print(f"Option {args.option} not found. Available: {[o.get('label') for o in options]}")
            return
        record_tz = data.get('tz', cli_tz_name)
        try:
            tz_zone, tz_name = resolve_tz(record_tz)
        except ValueError:
            tz_zone, tz_name = cli_tz_zone, cli_tz_name
        print(f"Manual mode — Option {selected.get('label')}: {selected.get('display','')}", flush=True)
        create_invite_draft(outlook, ns, data, selected, tz_zone, body=args.body)
        print("Invite draft opened. Review and send when ready.")
        return
    # ── End manual override ───────────────────────────────────────────────────

    log_records = load_log()
    replies     = find_reply_emails(ns)

    if not replies:
        print("No 'Re: Schedule Request:' emails found in inbox.")
        return

    print(f"Found {len(replies)} reply email(s).\n", flush=True)

    booked = 0
    needs_review = []

    for msg in replies:
        reply_subj = (getattr(msg, 'Subject', '') or '').strip()
        orig_subj  = original_subject(reply_subj)
        sender     = getattr(msg, 'SenderName', '') or getattr(msg, 'SenderEmailAddress', '')
        recv_date  = str(getattr(msg, 'ReceivedTime', ''))[:10]

        print(f"  Reply: {orig_subj}", flush=True)
        print(f"    From: {sender}  |  Received: {recv_date}", flush=True)

        # Recover original proposal data
        data = recover_data_from_log(orig_subj, log_records)
        if not data:
            data = recover_data_from_sent(ns, orig_subj)
        if not data:
            print("    ⚠  Could not recover proposal data — skipping.", flush=True)
            needs_review.append((reply_subj, sender, "no proposal data found"))
            continue

        options = data.get('options', [])
        if not options:
            print("    ⚠  No options in proposal data — skipping.", flush=True)
            needs_review.append((reply_subj, sender, "no options in data"))
            continue

        # Check if already on calendar
        if already_on_calendar(ns, data.get('meeting', ''), data.get('pid', '')):
            print(f"    ✓  Already on calendar — skipping.\n", flush=True)
            # Update log status
            for rec in log_records:
                if rec.get('subject') == orig_subj:
                    rec['status'] = 'booked'
            continue

        # Detect which option was selected
        selected, method = detect_option(msg, options)

        if not selected:
            print(f"    ⚠  Could not determine selection — flagged for review.\n", flush=True)
            body_preview = ' '.join((getattr(msg, 'Body', '') or '').split()[:30])
            needs_review.append((reply_subj, sender, f"ambiguous reply: \"{body_preview}...\""))
            continue

        lbl = selected.get('label', '?')
        print(f"    ✓  Selection detected: Option {lbl} (via {method})", flush=True)

        # Resolve timezone: prefer what was stored in the log; fall back to CLI arg
        record_tz = data.get('tz', cli_tz_name)
        try:
            tz_zone, tz_name = resolve_tz(record_tz)
        except ValueError:
            tz_zone, tz_name = cli_tz_zone, cli_tz_name
        print(f"    Timezone: {tz_name}", flush=True)

        # Create the draft invite
        create_invite_draft(outlook, ns, data, selected, tz_zone, body=args.body)
        print(f"    Invite draft opened. Review and send when ready.\n", flush=True)
        booked += 1

        # Update log
        for rec in log_records:
            if rec.get('subject') == orig_subj and rec.get('status') == 'pending':
                rec['status'] = 'booked'
                rec['booked_option'] = lbl
                rec['booked_date'] = date.today().isoformat()

    save_log(log_records)

    print(f"Done. {booked} draft invite(s) opened.", flush=True)
    if needs_review:
        print(f"\n{len(needs_review)} item(s) need manual review:", flush=True)
        for subj, sender, reason in needs_review:
            print(f"  • {subj}", flush=True)
            print(f"    From: {sender}", flush=True)
            print(f"    Reason: {reason}", flush=True)


if __name__ == '__main__':
    main()
