#!/usr/bin/env python3
"""
training_check.py
Scans Outlook inbox (last 60 days) for training obligations.
Applies critical thinking: deduplicates by name similarity, strips
deliveries the user is hosting, and cross-references completions so
finished trainings are not listed as due.
Output: Sends email to the current Outlook user — subject "Trainings Due"
Trigger: "what trainings are due?"
"""

import re
import sys
import html as html_module
from datetime import datetime, timedelta
from difflib import SequenceMatcher

import win32com.client

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

# ── Config ───────────────────────────────────────────────────────────────────
LOOKBACK_DAYS   = 60
EMAIL_SUBJECT   = "Trainings Due"


def get_current_user_email():
    """Return the SMTP email address of the currently signed-in Outlook user."""
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        accounts = outlook.Session.Accounts
        if accounts.Count > 0:
            return accounts.Item(1).SmtpAddress
    except Exception:
        pass
    return None

KNOWN_SENDERS   = [
    'skillsstudio@cisco.com',
    'learning@cisco.com',
    'no-reply@learning',
    'noreply@learning',
    'training@cisco.com',
    'compliance@cisco.com',
    'cisco-enablement@cisco.com',
    'enablement@cisco.com',
]

# Emails matching ANY of these (subject + body) are candidates
ASSIGNMENT_KEYWORDS = [
    'training due', 'training reminder', 'training assigned',
    'complete your training', 'training required', 'required training',
    'you have been assigned', 'training deadline', 'compliance training',
    'mandatory training', 'annual training', 'learning assigned',
    'course assigned', 'module assigned', 'action required.*train',
    'please complete.*train', 'train.*please complete',
    'complete by', 'due by.*train', 'train.*due by',
    'cx enablement', 'learning pathway', 'ai learning',
    'now available.*learn', 'new.*learning.*available',
    'enabler.*builder', 'builder.*enabler',
]

# Emails matching ANY of these signal completion
COMPLETION_KEYWORDS = [
    r'you have successfully completed',
    r'you(?:\'ve)? completed',
    r'certificate of completion',
    r'successfully completed',
    r'training completed',
    r'completion confirmation',
    r'course completed',
    r'congratulations.*complet',
    r'complet.*congratulations',
]

# If these appear Houston is likely delivering, not owing
DELIVERY_SIGNALS = [
    'you are hosting', 'you will facilitate', 'you are the instructor',
    'your attendees', 'your participants', 'class roster',
    'facilitator guide', 'you are scheduled to teach',
    'webex training invitation',
]

# Links containing these words are worth keeping
GOOD_LINK_WORDS = [
    'launch', 'start', 'begin', 'complete', 'take', 'access',
    'enroll', 'register', 'course', 'training', 'learning',
    'module', 'cisco.com',
]

# Links containing these are junk
JUNK_LINK_WORDS = [
    'unsubscribe', 'privacy', 'terms of', 'mailto:', 'tel:',
    'linkedin', 'twitter', 'facebook', 'instagram', '#',
]


# ── Helpers ──────────────────────────────────────────────────────────────────

def sim(a, b):
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()


def text_matches(text, patterns):
    t = text.lower()
    for p in patterns:
        if re.search(p, t):
            return True
    return False


def extract_links(html_body):
    """Pull training-relevant <a href> from HTML body.
    Tries good-match links first; falls back to any non-junk HTTP link."""
    if not html_body:
        return []

    def _parse(html):
        good, fallback = [], []
        if HAS_BS4:
            try:
                soup = BeautifulSoup(html, 'html.parser')
                anchors = [(a['href'], a.get_text(strip=True)) for a in soup.find_all('a', href=True)]
            except Exception:
                anchors = []
        else:
            anchors = [(m.group(1), m.group(2).strip())
                       for m in re.finditer(r'href=["\']([^"\']+)["\'][^>]*>([^<]*)', html)]

        for href, label in anchors:
            hl, ll = href.lower(), label.lower()
            if any(j in hl or j in ll for j in JUNK_LINK_WORDS):
                continue
            if not href.startswith(('http://', 'https://')):
                continue
            if any(g in hl or g in ll for g in GOOD_LINK_WORDS):
                good.append((label or href, href))
            else:
                fallback.append((label or href, href))
        return good, fallback

    good, fallback = _parse(html_body)
    results = good if good else fallback
    return results[:3]


def extract_due_date(body):
    """Return a due-date string from email body, or None."""
    patterns = [
        r'due\s+(?:date\s*:?\s*)?(?:by\s+)?([A-Za-z]+ \d{1,2}(?:st|nd|rd|th)?,?\s*\d{4})',
        r'(?:complete|finish|submit)\s+by\s+([A-Za-z]+ \d{1,2}(?:st|nd|rd|th)?,?\s*\d{4})',
        r'deadline\s*:?\s*([A-Za-z]+ \d{1,2}(?:st|nd|rd|th)?,?\s*\d{4})',
        r'due\s+(?:by\s+)?(\d{1,2}/\d{1,2}/\d{2,4})',
        r'complete\s+by\s+(\d{1,2}/\d{1,2}/\d{2,4})',
        r'by\s+([A-Za-z]+ \d{1,2},?\s+\d{4})',
        r'due\s*:?\s*(\d{4}-\d{2}-\d{2})',
        r'deadline\s*:?\s*(\d{1,2}/\d{1,2}/\d{2,4})',
    ]
    for p in patterns:
        m = re.search(p, body, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None


def extract_name(subject, body):
    """Return the cleanest training name we can find."""
    name = subject or ''
    # Strip common prefixes
    name = re.sub(
        r'^(re:|fw:|fwd:|reminder\s*:?|action\s+required\s*:?|notice\s*:?|alert\s*:?)\s*',
        '', name, flags=re.IGNORECASE
    ).strip()
    # Strip trailing noise
    name = re.sub(
        r'\s*[-–|]\s*(due\s*(?:date)?|reminder|notice|required|action|assigned).*$',
        '', name, flags=re.IGNORECASE
    ).strip()

    # If still looks generic, hunt the body for a quoted name
    if len(name) < 8 or text_matches(name, [r'^(training|reminder|action required|please complete)$']):
        m = re.search(r'["“‘]([^"”’]{10,100})["”’]', body)
        if m:
            name = m.group(1).strip()

    return name or subject or 'Unknown Training'


# ── Email scan ───────────────────────────────────────────────────────────────

def scan_inbox():
    outlook = win32com.client.Dispatch('Outlook.Application')
    ns      = outlook.GetNamespace('MAPI')
    inbox   = ns.GetDefaultFolder(6)

    cutoff = datetime.now() - timedelta(days=LOOKBACK_DAYS)

    items = inbox.Items
    items.Sort('[ReceivedTime]', True)   # newest first

    assignments  = []
    completions  = []

    for item in items:
        try:
            recv = item.ReceivedTime.replace(tzinfo=None)
        except Exception:
            continue
        if recv < cutoff:
            break

        try:
            subject = item.Subject or ''
            body    = item.Body    or ''
            sender  = (item.SenderEmailAddress or '').lower()
            full    = subject + ' ' + body

            # ── Completion email? ──────────────────────────────────────────
            if text_matches(full, COMPLETION_KEYWORDS):
                completions.append({'subject': subject, 'body': body, 'received': recv})
                continue  # a completion is never also an assignment

            # ── Delivery email? (Houston is hosting) ──────────────────────
            if text_matches(full, DELIVERY_SIGNALS):
                continue
            if 'webex.com' in sender and 'training' in full.lower() and 'you are required' not in full.lower():
                continue

            # ── Assignment/reminder? ──────────────────────────────────────
            from_known = any(k in sender for k in KNOWN_SENDERS)
            keyword_hit = text_matches(full, ASSIGNMENT_KEYWORDS)

            if from_known or keyword_hit:
                html_body = ''
                try:
                    html_body = item.HTMLBody or ''
                except Exception:
                    pass

                assignments.append({
                    'name'     : extract_name(subject, body),
                    'subject'  : subject,
                    'body'     : body,
                    'received' : recv,
                    'sender'   : sender,
                    'due_date' : extract_due_date(body),
                    'links'    : extract_links(html_body),
                })

        except Exception:
            continue

    return assignments, completions


# ── Dedup + status ───────────────────────────────────────────────────────────

def check_completed(name, completions):
    """Return completion date if a completion email clearly mentions this training, else None."""
    for c in completions:
        # Strong subject-line similarity (same title, basically)
        if sim(name, c['subject']) > 0.75:
            return c['received']
        # Name must be specific (20+ chars) and appear in the completion subject
        if len(name) >= 20 and name.lower()[:50] in c['subject'].lower():
            return c['received']
    return None


def build_training_list(assignments, completions):
    trainings  = []
    seen_names = []

    for a in assignments:
        name = a['name']

        # Deduplicate (keep earliest — likely the original notice)
        duplicate = False
        for seen in seen_names:
            if sim(name, seen) > 0.75:
                duplicate = True
                break
        if duplicate:
            continue
        seen_names.append(name)

        completed_date = check_completed(name, completions)
        if completed_date:
            status = 'completed'
        elif a['due_date']:
            status = 'due'
        else:
            status = 'optional'

        trainings.append({
            'name'          : name,
            'due_date'      : a['due_date'],
            'links'         : a['links'],
            'status'        : status,
            'completed_date': completed_date,
        })

    # Surface completion-only items (assignment older than 60 days)
    for c in completions:
        name = extract_name(c['subject'], c['body'])
        duplicate = any(sim(name, s) > 0.75 for s in seen_names)
        if not duplicate:
            seen_names.append(name)
            trainings.append({
                'name'          : name,
                'due_date'      : None,
                'links'         : [],
                'status'        : 'completed',
                'completed_date': c['received'],
            })

    return trainings


# ── HTML builder ─────────────────────────────────────────────────────────────

def build_html(trainings):
    due_list       = [t for t in trainings if t['status'] == 'due']
    optional_list  = [t for t in trainings if t['status'] == 'optional']
    completed_list = [t for t in trainings if t['status'] == 'completed']

    def make_row(t, idx, kind):
        bg   = '#f0f4f8' if idx % 2 == 0 else '#ffffff'
        name = html_module.escape(t['name'])

        if t['links']:
            _, url = t['links'][0]
            name_cell = f'<a href="{html_module.escape(url)}" style="color:#1565A0;">{name}</a>'
        else:
            name_cell = name

        if kind == 'completed':
            cd = t.get('completed_date')
            date_str = cd.strftime('%B %d, %Y') if cd and hasattr(cd, 'strftime') else '—'
            date_cell = f'<span style="color:#2e7d32;font-weight:bold;">&#10003; {date_str}</span>'
        elif t['due_date']:
            date_cell = html_module.escape(t['due_date'])
        else:
            date_cell = 'Optional'

        return (
            f'<tr style="background:{bg};">'
            f'<td style="padding:5px 10px;border:1px solid #ddd;">{name_cell}</td>'
            f'<td style="padding:5px 10px;border:1px solid #ddd;">{date_cell}</td>'
            f'</tr>'
        )

    def make_section(title, rows, kind, header_color, col2_label='Due Date'):
        if not rows:
            return ''
        row_html = ''.join(make_row(r, i, kind) for i, r in enumerate(rows))
        return (
            f'<p style="font-family:Aptos,Calibri,sans-serif;font-size:12pt;margin-top:18px;">'
            f'<strong>{title}</strong></p>'
            f'<table style="border-collapse:collapse;font-family:Aptos,Calibri,sans-serif;font-size:12pt;">'
            f'<tr style="background:{header_color};color:#fff;">'
            f'<th style="padding:5px 10px;border:1px solid #ddd;text-align:left;min-width:280px;">Training</th>'
            f'<th style="padding:5px 10px;border:1px solid #ddd;text-align:left;min-width:160px;">{col2_label}</th>'
            f'</tr>{row_html}</table>'
        )

    sections = (
        make_section('Required Trainings',     due_list,       'due',       '#1565A0', 'Due Date') +
        make_section('Recently Completed',     completed_list, 'completed', '#2e7d32', 'Date Completed') +
        make_section('Optional / No Due Date', optional_list,  'optional',  '#555555', 'Due Date')
    )

    if not sections:
        sections = '<p>No training obligations found in the past 60 days.</p>'

    generated = datetime.now().strftime('%B %d, %Y')
    return (
        f'<div style="font-family:Aptos,Calibri,sans-serif;font-size:12pt;">'
        f'<p>Here is a summary of your training obligations based on emails from the past {LOOKBACK_DAYS} days.</p>'
        f'{sections}'
        f'<p style="color:#888;font-size:10pt;margin-top:20px;">'
        f'Generated {generated} &nbsp;|&nbsp; Lookback: {LOOKBACK_DAYS} days</p>'
        f'</div>'
    )


# ── Send draft ────────────────────────────────────────────────────────────────

def create_draft(html_body, recipient):
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To       = recipient
    mail.Subject  = EMAIL_SUBJECT
    mail.HTMLBody = html_body
    mail.Send()
    print(f'Sent: {EMAIL_SUBJECT}')
    print(f'To: {recipient}')


# ── 50-day gate ───────────────────────────────────────────────────────────────

import os, pathlib

LAST_RUN_FILE  = pathlib.Path.home() / '.claude' / 'training_check_last_run.txt'
SCHEDULE_DAYS  = 50    # only do real work when this many days have passed
FORCE_FLAG     = '--force'   # bypass gate: python training_check.py --force


def scheduled_run_due():
    """Return True if 50+ days have passed since last scheduled run."""
    if FORCE_FLAG in sys.argv:
        return True
    if not LAST_RUN_FILE.exists():
        return True
    try:
        last = datetime.fromisoformat(LAST_RUN_FILE.read_text().strip())
        return (datetime.now() - last).days >= SCHEDULE_DAYS
    except Exception:
        return True


def mark_last_run():
    LAST_RUN_FILE.parent.mkdir(parents=True, exist_ok=True)
    LAST_RUN_FILE.write_text(datetime.now().isoformat())


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    # When invoked by Task Scheduler (no --force), only run every 50 days
    if '--force' not in sys.argv and not scheduled_run_due():
        print('Training Check — skipping (fewer than 50 days since last run).')
        return

    recipient = get_current_user_email()
    if not recipient:
        print('Error: could not detect current Outlook user email. Is Outlook open?')
        return
    print(f'Training Check — scanning Outlook inbox (last {LOOKBACK_DAYS} days)...')
    print(f'Recipient: {recipient}')
    assignments, completions = scan_inbox()
    print(f'  Assignment emails: {len(assignments)}')
    print(f'  Completion emails: {len(completions)}')

    trainings = build_training_list(assignments, completions)

    due       = sum(1 for t in trainings if t['status'] == 'due')
    optional  = sum(1 for t in trainings if t['status'] == 'optional')
    completed = sum(1 for t in trainings if t['status'] == 'completed')
    print(f'  Result — Due: {due}  |  Optional: {optional}  |  Completed: {completed}')

    html_body = build_html(trainings)
    create_draft(html_body, recipient)
    mark_last_run()


if __name__ == '__main__':
    main()
