#!/usr/bin/env python3
"""
Schedule Requester — Find mutual availability and open a scheduling request email draft.
Called by the Schedule Requester agent after the prompting flow is complete.

Usage:
  python schedule_requester.py \
    --meeting "Wireless Meraki Ask the Expert" \
    --pid F51119 \
    --to alexis.westmoreland@houstontx.gov \
    [--to another@customer.gov] \
    [--cc jnissao@cisco.com] \
    --sme chrirome@cisco.com \
    [--duration 45] \
    [--type monthly|one-time] \
    [--months 5]
"""

import argparse, re, sys, json, hashlib
import win32com.client
from datetime import datetime, timedelta, date
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
from pathlib import Path

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
    key  = name.strip().lower()
    iana = TZ_ALIASES.get(key, name.strip())
    try:
        return ZoneInfo(iana), iana
    except ZoneInfoNotFoundError:
        raise ValueError(
            f"Unknown timezone '{name}'. Use an IANA name (e.g. 'America/Chicago') "
            f"or a short alias: central, eastern, pacific, mountain."
        )

REQUESTS_LOG = Path(r"C:\Users\hodowns\OneDrive - Cisco\Desktop\Schedule Requester\pending-requests.json")

INTERVAL = 15  # free/busy resolution in minutes

WEEKDAY_NAMES = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
NTH_NAMES     = ['','1st','2nd','3rd','4th','5th']
OPTION_LABELS = ['A','B','C','D','E']


# ── Args ───────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument('--meeting',  required=True)
    p.add_argument('--pid',      required=True)
    p.add_argument('--to',       action='append', default=[], dest='to_list')
    p.add_argument('--cc',       action='append', default=[], dest='cc_list')
    p.add_argument('--sme',      action='append', default=[], dest='sme_list')
    p.add_argument('--duration', type=int, default=45)
    p.add_argument('--type',     choices=['one-time','monthly'], default='monthly')
    p.add_argument('--months',   type=int, default=5)
    p.add_argument('--target-date', default=None,
                   help='Limit one-time search to a specific date (YYYY-MM-DD)')
    p.add_argument('--skip-self', action='store_true',
                   help="Don't include Houston's calendar in the availability check")
    p.add_argument('--tz', default=DEFAULT_TZ,
                   help='Timezone for proposed times (default: America/Chicago). '
                        'Accepts IANA names or aliases: central, eastern, pacific, mountain.')
    return p.parse_args()


# ── Free/busy helpers ──────────────────────────────────────────────────────────

def fetch_fb(ns, email, anchor):
    """Return (fb_string, anchor_datetime) for midnight of anchor date."""
    start = datetime(anchor.year, anchor.month, anchor.day)
    try:
        recip = ns.CreateRecipient(email)
        if recip.Resolve():
            return recip.FreeBusy(start, INTERVAL, True), start
    except Exception:
        pass
    return None, start


def window_free(fb, fb_start, cs, ce):
    """True if every interval in [cs, ce) is '0' in fb; None if fb unavailable."""
    if fb is None:
        return None
    i0 = int((cs - fb_start).total_seconds() // 60 // INTERVAL)
    i1 = int((ce - fb_start).total_seconds() // 60 // INTERVAL)
    for i in range(i0, i1):
        if i >= len(fb) or fb[i] != '0':
            return False
    return True


def mutual_free(ns, emails, cs, ce):
    """True only when every email's calendar is free for [cs, ce)."""
    anchor = datetime(cs.year, cs.month, cs.day)
    for email in emails:
        fb, fb_start = fetch_fb(ns, email, anchor)
        result = window_free(fb, fb_start, cs, ce)
        if not result:   # False or None both block
            return False
    return True


# ── One-time slot finder ───────────────────────────────────────────────────────

def find_onetime_slots(ns, emails, duration, n=5, target_date=None):
    today = date.today()
    proposals = []
    day_range = [target_date] if target_date else [today + timedelta(days=i) for i in range(1, 21)]
    for day in day_range:
        if day.weekday() >= 5:
            continue
        anchor = datetime(day.year, day.month, day.day)
        # Pre-fetch fb for all emails for this day
        fbs = {e: fetch_fb(ns, e, day) for e in emails}
        day_picks = []
        for hour in range(9, 17):
            for minute in (0, 15, 30, 45):
                cs = anchor.replace(hour=hour, minute=minute)
                ce = cs + timedelta(minutes=duration)
                if ce.hour > 17 or (ce.hour == 17 and ce.minute > 0):
                    continue
                if all(window_free(fbs[e][0], fbs[e][1], cs, ce) for e in emails):
                    day_picks.append((cs, ce))
                    break  # one per hour
            if day_picks:
                break  # one per day
        proposals.extend(day_picks)
        if not target_date and len(proposals) >= n:
            break
    return proposals[:n]


# ── Monthly recurring slot finder ──────────────────────────────────────────────

def which_nth(dt):
    return (dt.day - 1) // 7 + 1, dt.weekday()


def nth_weekday(year, month, n, wday):
    first = datetime(year, month, 1)
    delta = (wday - first.weekday()) % 7
    return first + timedelta(days=delta + (n - 1) * 7)


def upcoming_months(count):
    today = date.today()
    out = []
    for i in range(1, count + 1):
        raw = today.month + i - 1
        y = today.year + raw // 12
        m = raw % 12 + 1
        out.append((y, m))
    return out


def find_recurring_slots(ns, emails, duration, months=5):
    today = date.today()
    # Scan the next 14 calendar days for anchor free slots
    anchor_day = today + timedelta(days=1)
    while anchor_day.weekday() >= 5:
        anchor_day += timedelta(days=1)

    seen, scored = set(), []
    for offset in range(0, 14):
        day = anchor_day + timedelta(days=offset)
        if day.weekday() >= 5:
            continue
        dt_day = datetime(day.year, day.month, day.day)
        fbs = {e: fetch_fb(ns, e, day) for e in emails}

        for hour in range(9, 17):
            for minute in (0, 15, 30, 45):
                cs = dt_day.replace(hour=hour, minute=minute)
                ce = cs + timedelta(minutes=duration)
                if ce.hour > 17 or (ce.hour == 17 and ce.minute > 0):
                    continue
                if not all(window_free(fbs[e][0], fbs[e][1], cs, ce) for e in emails):
                    continue

                n, wday = which_nth(cs)
                key = (n, wday, hour, minute)
                if key in seen:
                    continue
                seen.add(key)

                # Check this pattern across future months
                occs, free_count = [], 0
                for y, m in upcoming_months(months):
                    occ = nth_weekday(y, m, n, wday)
                    occ_cs = occ.replace(hour=hour, minute=minute)
                    occ_ce = occ_cs + timedelta(minutes=duration)
                    free = mutual_free(ns, emails, occ_cs, occ_ce)
                    occs.append((occ, free))
                    if free:
                        free_count += 1

                scored.append({
                    'n': n, 'wday': wday, 'hour': hour, 'minute': minute,
                    'kickoff': cs, 'kickoff_end': ce,
                    'occurrences': occs, 'score': free_count,
                })

    def rank(s):
        h = s['hour'] + s['minute'] / 60
        return (-s['score'], min(abs(h - 10), abs(h - 14)))

    scored.sort(key=rank)
    return scored[:4]


# ── Pending-requests log ──────────────────────────────────────────────────────

def load_requests_log():
    if not REQUESTS_LOG.exists():
        return []
    try:
        return json.loads(REQUESTS_LOG.read_text(encoding='utf-8'))
    except Exception:
        return []

def save_requests_log(records):
    REQUESTS_LOG.parent.mkdir(parents=True, exist_ok=True)
    REQUESTS_LOG.write_text(json.dumps(records, indent=2), encoding='utf-8')

def build_request_record(args, subject, proposals, proposal_dicts, tz_name):
    req_id = hashlib.md5(f"{subject}{date.today().isoformat()}".encode()).hexdigest()[:12]
    return {
        'id':        req_id,
        'subject':   subject,
        'meeting':   args.meeting,
        'pid':       args.pid.upper(),
        'type':      args.type,
        'duration':  args.duration,
        'tz':        tz_name,
        'sent_date': date.today().isoformat(),
        'to':        args.to_list,
        'cc':        args.cc_list + [s for s in args.sme_list if s not in args.cc_list],
        'options':   proposal_dicts,
        'status':    'pending',
    }

def proposals_to_dicts(proposals, meeting_type, duration):
    """Serialize proposals to JSON-safe dicts for the log and email embed."""
    out = []
    for i, p in enumerate(proposals[:5]):
        lbl = OPTION_LABELS[i] if i < len(OPTION_LABELS) else str(i+1)
        if meeting_type == 'monthly':
            out.append({
                'label':   lbl,
                'type':    'monthly',
                'n':       p['n'],
                'wday':    p['wday'],
                'hour':    p['hour'],
                'minute':  p['minute'],
                'kickoff': p['kickoff'].strftime('%Y-%m-%d'),
                'duration': duration,
            })
        else:
            cs, ce = p
            out.append({
                'label':    lbl,
                'type':     'one-time',
                'start':    cs.strftime('%Y-%m-%dT%H:%M:%S'),
                'end':      ce.strftime('%Y-%m-%dT%H:%M:%S'),
                'duration': duration,
            })
    return out


# ── HTML helpers ───────────────────────────────────────────────────────────────

def _esc(s):
    return (s or '').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')


def _fmt_time(dt):
    return dt.strftime('%I:%M').lstrip('0') or '12:00'


def _fmt_day(dt):
    """Short date without zero-padded day — Windows-safe."""
    return dt.strftime('%b ') + str(dt.day)


def _fmt_long_day(dt):
    return dt.strftime('%A, %B ') + str(dt.day)


def _table_row(cells, bg):
    """Render one table row with the standard two-tone style."""
    tds = ''.join(
        f"<td style='padding:5px {pad};font-size:11pt;{style};white-space:nowrap;background-color:{bg};'>{val}</td>"
        for val, pad, style in cells
    )
    return f"<tr>{tds}</tr>"


def _table_header(*labels):
    ths = ''.join(
        f"<th style='padding:6px 14px 6px {'0' if i else '14px'};font-size:11pt;font-weight:700;"
        f"color:#ffffff;text-align:left;white-space:nowrap;'>{lbl}</th>"
        for i, lbl in enumerate(labels)
    )
    return (
        "<table style='border-collapse:collapse;width:auto;margin:16px 0;border:1px solid #1565A0;'>"
        f"<thead><tr style='background-color:#1565A0;'>{ths}</tr></thead><tbody>"
    )


ROW_EVEN = '#FFFFFF'
ROW_ODD  = '#EBF3FB'


def build_monthly_table(patterns):
    html = _table_header('Option', 'Pattern', 'Kickoff', 'Upcoming Dates')
    for i, p in enumerate(patterns[:3]):
        bg    = ROW_EVEN if i % 2 == 0 else ROW_ODD
        lbl   = f"Option {OPTION_LABELS[i]}"
        nth   = NTH_NAMES[p['n']]
        wday  = WEEKDAY_NAMES[p['wday']]
        t0    = _fmt_time(p['kickoff'])
        t1    = p['kickoff_end'].strftime('%I:%M %p').lstrip('0')
        pattern_str = f"{nth} {wday} &bull; {t0}&ndash;{t1} CT"
        kick  = _fmt_day(p['kickoff'])
        future = ', '.join(_fmt_day(o[0]) for o in p['occurrences'][:4])
        html += _table_row([
            (lbl,         '14px', 'font-weight:700;color:#1565A0'),
            (pattern_str, '18px 5px 0', 'font-weight:600;color:#1f1f1f'),
            (kick,        '18px 5px 0', 'color:#444444'),
            (future,      '14px 5px 0', 'color:#444444'),
        ], bg)
    return html + "</tbody></table>"


def build_onetime_table(proposals):
    html = _table_header('Option', 'Date', 'Time (CT)')
    for i, (cs, ce) in enumerate(proposals[:5]):
        bg  = ROW_EVEN if i % 2 == 0 else ROW_ODD
        lbl = f"Option {OPTION_LABELS[i] if i < len(OPTION_LABELS) else i + 1}"
        html += _table_row([
            (lbl,                                              '14px', 'font-weight:700;color:#1565A0'),
            (_fmt_long_day(cs),                               '18px 5px 0', 'font-weight:600;color:#1f1f1f'),
            (f"{_fmt_time(cs)} &ndash; {ce.strftime('%I:%M %p').lstrip('0')}", '14px 5px 0', 'color:#444444'),
        ], bg)
    return html + "</tbody></table>"


def build_email_html(meeting, duration, meeting_type, proposals, first_names, data_json=None):
    greeting  = first_names[0] if first_names else 'Team'
    cadence   = 'a monthly recurring' if meeting_type == 'monthly' else 'a'
    table_html = (build_monthly_table(proposals) if meeting_type == 'monthly'
                  else build_onetime_table(proposals))
    # Embed structured data as a hidden comment so Schedule Booker can recover
    # the options from the reply thread without needing a local log lookup.
    data_comment = f"\n<!-- sr-data: {data_json} -->\n" if data_json else ''
    return (
        "<div style=\"font-family:'Aptos',Calibri,Arial,sans-serif;font-size:12pt;"
        "color:#1f1f1f;max-width:680px;line-height:1.5;\">"
        f"<p>Hi {_esc(greeting)},</p>"
        f"<p>I'd like to get {cadence} <strong>{_esc(meeting)}</strong> session on the calendar. "
        "Below are a few times that work on our end — please let me know which option fits best "
        "for you and your team.</p>"
        f"{table_html}"
        "<p>Once you confirm a preference, I'll send the calendar invite right away.</p>"
        "<p>Best regards,<br>Houston Downs<br>Cisco CX Program Manager</p>"
        f"{data_comment}"
        "</div>"
    )


# ── Outlook draft ──────────────────────────────────────────────────────────────

def open_draft(outlook, subject, html_body, to_list, cc_list):
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Display(False)

    existing = mail.HTMLBody or ''
    tag = re.search(r'(<body[^>]*>)', existing, re.IGNORECASE)
    if tag:
        ins = tag.end()
        mail.HTMLBody = existing[:ins] + html_body + existing[ins:]
    else:
        mail.HTMLBody = (
            "<html><body style=\"font-family:'Aptos',Calibri,Arial,sans-serif;"
            "font-size:12pt;color:#1f1f1f;max-width:680px;line-height:1.5;\">"
            + html_body + "</body></html>"
        )
    for email in to_list:
        r = mail.Recipients.Add(email); r.Type = 1
    for email in cc_list:
        r = mail.Recipients.Add(email); r.Type = 2
    if to_list or cc_list:
        mail.Recipients.ResolveAll()


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()

    outlook = win32com.client.Dispatch('Outlook.Application')
    ns      = outlook.GetNamespace('MAPI')

    try:
        eu = ns.CurrentUser.AddressEntry.GetExchangeUser()
        houston_email = eu.PrimarySmtpAddress
    except Exception:
        houston_email = 'hodowns@cisco.com'

    if args.skip_self:
        cisco_emails = [e for e in args.sme_list if e]
    else:
        cisco_emails = [houston_email] + [e for e in args.sme_list if e]

    try:
        tz_zone, tz_name = resolve_tz(args.tz)
    except ValueError as e:
        print(f"ERROR: {e}", flush=True)
        sys.exit(1)

    print(f"Schedule Requester — {args.meeting} | {args.pid.upper()}", flush=True)
    print(f"Type: {args.type}  |  Duration: {args.duration} min  |  Timezone: {tz_name}", flush=True)
    print(f"Checking: {', '.join(cisco_emails)}\n", flush=True)

    target_date = None
    if args.target_date:
        target_date = datetime.strptime(args.target_date, '%Y-%m-%d').date()

    if args.type == 'monthly':
        print("Finding recurring monthly patterns...", flush=True)
        proposals = find_recurring_slots(ns, cisco_emails, args.duration, args.months)
        if not proposals:
            print("No consistent recurring patterns found. Try --months with a larger number.")
            sys.exit(1)
        for i, p in enumerate(proposals[:3]):
            score = p['score']
            nth   = NTH_NAMES[p['n']]
            wday  = WEEKDAY_NAMES[p['wday']]
            t     = _fmt_time(p['kickoff'])
            print(f"  Option {OPTION_LABELS[i]}: {nth} {wday} @ {t} AM — {score}/{args.months} months free")
    else:
        print("Finding one-time slots...", flush=True)
        proposals = find_onetime_slots(ns, cisco_emails, args.duration, target_date=target_date)
        if not proposals:
            print("No mutual free slots found in the next 14 days.")
            sys.exit(1)
        for i, (cs, ce) in enumerate(proposals):
            print(f"  Option {OPTION_LABELS[i]}: {_fmt_long_day(cs)} {_fmt_time(cs)}–{ce.strftime('%I:%M %p').lstrip('0')}")

    # Derive first name(s) for email greeting
    first_names = []
    for email in args.to_list:
        local = email.split('@')[0]
        parts = re.split(r'[._]', local)
        if parts:
            first_names.append(parts[0].capitalize())

    subject        = f"Schedule Request: {args.meeting} [EXTERNAL] {args.pid.upper()}"
    proposal_dicts = proposals_to_dicts(proposals, args.type, args.duration)
    data_json      = json.dumps({'meeting': args.meeting, 'pid': args.pid.upper(),
                                 'type': args.type, 'duration': args.duration,
                                 'to': args.to_list, 'cc': args.cc_list,
                                 'options': proposal_dicts})
    html_body = build_email_html(args.meeting, args.duration, args.type, proposals,
                                 first_names, data_json=data_json)

    full_cc = list(args.cc_list)
    for sme in args.sme_list:
        if sme not in full_cc:
            full_cc.append(sme)

    open_draft(outlook, subject, html_body, args.to_list, full_cc)

    # Write to pending-requests log
    record  = build_request_record(args, subject, proposals, proposal_dicts, tz_name)
    records = load_requests_log()
    records.append(record)
    save_requests_log(records)
    print(f"Logged to: {REQUESTS_LOG}", flush=True)

    print(f"\nDraft opened: {subject}", flush=True)
    print(f"To:  {', '.join(args.to_list)}", flush=True)
    if full_cc:
        print(f"CC:  {', '.join(full_cc)}", flush=True)


if __name__ == '__main__':
    main()
