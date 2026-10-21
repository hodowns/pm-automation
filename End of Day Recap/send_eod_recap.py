#!/usr/bin/env python3
"""
End of Day Recap Email Sender
Scans today's Outlook calendar for [Internal] and [External] meetings hosted
by the user, extracts the PID from each subject, finds matching Webex recap
emails, opens a pre-filled Outlook draft per meeting, and files a copy to
the PID's Meeting Minutes folder in SharePoint.

Usage:   python send_eod_recap.py
         python send_eod_recap.py E55085,F51119       (manual PID override)
         python send_eod_recap.py --date 2026-08-24   (different date)
         python send_eod_recap.py --yesterday          (shortcut for yesterday)
"""

import os
import sys
import re
import time
import json
import requests
import win32com.client
from bs4 import BeautifulSoup
from datetime import datetime, timedelta, timezone
from pathlib import Path

WEBEX_BOT        = "messenger@webex.com"
SEARCH_DAYS      = 14
ME_PATTERNS      = ["hodowns", "houston downs"]
WEBEX_BASE       = "https://webexapis.com/v1"
WEBEX_TOKENS_FILE = Path(r"C:\Users\hodowns\webex-mcp-oauth\tokens.json")


def _load_webex_token():
    try:
        return json.loads(WEBEX_TOKENS_FILE.read_text())["access_token"]
    except Exception:
        return ""


def _webex_get(path, token, params=None):
    try:
        r = requests.get(f"{WEBEX_BASE}/{path.lstrip('/')}",
                         headers={"Authorization": f"Bearer {token}"},
                         params=params or {}, timeout=15)
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return None


def fetch_webex_participants(title, date_str):
    """Return participants from the Webex API for a matching ended meeting."""
    token = _load_webex_token()
    if not token:
        return []
    try:
        meeting_date = datetime.strptime(date_str.strip(), "%A, %B %d, %Y")
        from_dt = (meeting_date - timedelta(days=1)).strftime("%Y-%m-%dT00:00:00Z")
        to_dt   = (meeting_date + timedelta(days=1)).strftime("%Y-%m-%dT23:59:59Z")
    except (ValueError, AttributeError):
        now = datetime.now(timezone.utc)
        from_dt = (now - timedelta(days=14)).strftime("%Y-%m-%dT%H:%M:%SZ")
        to_dt   = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    me = _webex_get("people/me", token)
    if not me:
        return []
    host_email = me.get("emails", [""])[0]
    data = _webex_get("meetings", token, {"hostEmail": host_email, "from": from_dt,
                                          "to": to_dt, "meetingType": "meeting",
                                          "state": "ended", "max": 50})
    if not data:
        return []
    kws = [w.lower() for w in title.split() if len(w) > 2]
    best, best_score = None, 0
    for m in data.get("items", []):
        score = sum(1 for k in kws if k in m.get("title", "").lower())
        if score > best_score:
            best_score, best = score, m
    if not best or best_score == 0:
        return []
    pdata = _webex_get("meetingParticipants", token, {"meetingId": best["id"]})
    if not pdata:
        return []
    seen, participants = set(), []
    for p in pdata.get("items", []):
        email = (p.get("email") or "").lower().strip()
        if email and "@guest.webex.localhost" not in email and email not in seen:
            seen.add(email)
            participants.append({"name": p.get("displayName") or email,
                                  "email": email, "type": 1})
    return participants
PID_RE          = re.compile(r'[A-Z]\d{4,6}')
SHAREPOINT_BASE = Path(
    r"C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS"
)

_FOOTER_RE = re.compile(
    r"need help\?|help\.webex\.com|cisco and/or its affiliates"
    r"|all rights reserved|privacy statement|terms of service"
    r"|view meeting content|view meeting recap"
    r"|claude notification incident",
    re.IGNORECASE,
)

def _is_footer(text):
    return bool(_FOOTER_RE.search(text))

def _strip_inline_url(text):
    return re.sub(r"\s*<https?://[^\s>]+>", "", text).strip()

def _strip_bullet(text):
    return re.sub(r"^[\*\-•]\s*", "", text).strip()


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

TARGET_DATE = datetime.today().date()  # overridden by --date / --yesterday at startup


def _parse_meeting_date(date_str):
    """Parse 'Monday, August 12, 2026' → date object. Returns None on failure."""
    try:
        return datetime.strptime(date_str.strip(), "%A, %B %d, %Y").date()
    except ValueError:
        return None

def _is_target_date(date_str):
    parsed = _parse_meeting_date(date_str)
    if parsed is None:
        return False
    return parsed == TARGET_DATE


# ---------------------------------------------------------------------------
# Outlook search — returns ALL matching recap emails, not just the first
# ---------------------------------------------------------------------------

def find_all_webex_emails(inbox, search_term):
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    cutoff = (datetime.now() - timedelta(days=SEARCH_DAYS)).strftime("%m/%d/%Y %I:%M %p")
    try:
        restricted = items.Restrict(f"[ReceivedTime] >= '{cutoff}'")
        if restricted.Count > 0:
            items = restricted
    except Exception:
        pass

    keywords = [w for w in search_term.lower().split() if w]
    matches = []
    for msg in items:
        try:
            if getattr(msg, "SenderEmailAddress", "").lower() != WEBEX_BOT:
                continue
            subject = (getattr(msg, "Subject", "") or "")
            if "meeting content is available" not in subject.lower():
                continue
            haystack = (subject + " " + (getattr(msg, "Body", "") or "")).lower()
            if all(kw in haystack for kw in keywords):
                matches.append(msg)
        except Exception:
            continue
    return matches


# ---------------------------------------------------------------------------
# Email parsing (identical to send_recap.py)
# ---------------------------------------------------------------------------

def parse_email(msg):
    raw  = (msg.Body or "").replace("\r\n", "\n").replace("\r", "\n")
    html = msg.HTMLBody or ""
    soup = BeautifulSoup(html, "html.parser")
    lines = [l.strip() for l in raw.split("\n") if l.strip()]

    data = {
        "title":           "",
        "host":            "",
        "date":            "",
        "overview":        "",
        "accomplishments": [],
        "action_items":    [],
        "rec_link":        None,
        "rec_password":    None,
        "chapters":        [],
    }

    for i, line in enumerate(lines):
        if "meeting content is available" in line.lower() and i + 1 < len(lines):
            data["title"] = _strip_inline_url(lines[i + 1])
            break

    for line in lines:
        if line.lower().startswith("host:"):
            data["host"] = line.split(":", 1)[1].strip()
            break

    for line in lines:
        if re.match(r"(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),", line):
            data["date"] = line
            break

    state = None
    overview_lines = []

    for line in lines:
        if line == "Overview":
            state = "overview"; continue
        if re.search(r"Meeting summary.{0,5}AI.generated", line, re.IGNORECASE):
            state = "overview"; continue
        if "Detailed Summary" in line:
            state = "detailed"; continue
        if re.search(r"Action items.{0,5}AI.generated", line, re.IGNORECASE):
            state = "actions"; continue
        if line == "Recording":
            state = "recording"; continue
        if re.search(r"Chapters.{0,5}AI.generated", line, re.IGNORECASE):
            state = "chapters"; continue

        if state == "overview":
            if re.match(r"^[\*\-]\t", line):
                state = None
            else:
                overview_lines.append(line)
        elif state == "detailed":
            if len(data["accomplishments"]) < 5:
                clean = _strip_inline_url(_strip_bullet(line))
                if clean and not _is_footer(clean):
                    data["accomplishments"].append(clean)

        elif state == "actions":
            clean = _strip_inline_url(_strip_bullet(line))
            if clean and not _is_footer(clean):
                data["action_items"].append(clean)
        elif state == "chapters":
            stripped = _strip_inline_url(_strip_bullet(line))
            m = re.match(r"(\d{2}:\d{2}:\d{2})\s+(.*)", stripped)
            if m:
                data["chapters"].append({"time": m.group(1), "title": m.group(2).strip()})
        elif state == "recording":
            url_m = re.search(r"<(https?://[^\s>]+)>", line)
            if url_m and not data["rec_link"]:
                data["rec_link"] = url_m.group(1)
                after = line[url_m.end():]
                pw_m  = re.search(r"\t\s*(\S+)", after)
                if pw_m:
                    data["rec_password"] = pw_m.group(1).strip()

    full_overview = " ".join(overview_lines).strip()
    sentences = re.split(r'(?<=[.!?])\s+', full_overview)
    data["overview"] = " ".join(sentences[:3])

    if not data["rec_link"]:
        _extract_recording_html(soup, lines, data)

    return data


def _extract_recording_html(soup, lines, data):
    for table in soup.find_all("table"):
        if "Topic" in table.get_text() and "Password" in table.get_text():
            for row in table.find_all("tr"):
                cells = row.find_all("td")
                if len(cells) >= 2:
                    link_tag = cells[0].find("a", href=True)
                    if link_tag and "webex.com" in link_tag.get("href", ""):
                        data["rec_link"]     = link_tag["href"]
                        data["rec_password"] = cells[1].get_text(strip=True)
                        return
            break

    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "webex.com" in href and any(k in href for k in ["RCID", "recording", "ldr.php"]):
            data["rec_link"] = href
            break

    if not data["rec_password"]:
        for i, line in enumerate(lines):
            if re.search(r"Topic\s+Password", line, re.IGNORECASE):
                if i + 2 < len(lines):
                    candidate = lines[i + 2].strip()
                    if re.match(r"^[A-Za-z0-9]{4,16}$", candidate):
                        data["rec_password"] = candidate
                break


# ---------------------------------------------------------------------------
# Calendar attendees
# ---------------------------------------------------------------------------

def _resolve_exchange_email(recipient):
    try:
        ae = recipient.AddressEntry
        if ae.AddressEntryUserType == 0:
            eu = ae.GetExchangeUser()
            if eu:
                return eu.PrimarySmtpAddress
        props = ae.PropertyAccessor
        return props.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
    except Exception:
        return None

def get_attendees(calendar, title):
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]", True)

    start_b = (datetime.now() - timedelta(days=SEARCH_DAYS)).strftime("%m/%d/%Y %I:%M %p")
    end_b   = (datetime.now() + timedelta(days=1)).strftime("%m/%d/%Y %I:%M %p")
    try:
        restricted = items.Restrict(f"[Start] >= '{start_b}' AND [Start] <= '{end_b}'")
        if restricted.Count > 0:
            items = restricted
    except Exception:
        pass

    clean_title = re.sub(r"\[(?:EXTERNAL|INTERNAL|EXT|INT)\]\s*", "", title, flags=re.IGNORECASE)
    clean_title = re.sub(r"[^\w\s]", " ", clean_title)
    keywords = [w for w in clean_title.lower().split() if len(w) >= 3]
    pid_match = re.search(r'\b[A-Z]\d{4,6}\b', title)
    required_pid = pid_match.group(0).lower() if pid_match else None
    min_matches = max(3, len(keywords) // 2)

    best_item  = None
    best_score = 0

    for item in items:
        try:
            subject = (getattr(item, "Subject", "") or "").lower()
            if required_pid and required_pid not in subject:
                continue
            subject_clean = re.sub(r"[^\w\s]", " ", subject)
            matches = sum(1 for kw in keywords if kw in subject_clean)
            if matches >= min_matches and matches > best_score:
                best_score = matches
                best_item  = item
        except Exception:
            continue

    if best_item is None:
        return []

    print(f"  Calendar:  {getattr(best_item, 'Subject', '(unknown)')} (score {best_score}/{len(keywords)})")

    try:
        recips = best_item.Recipients
        attendees = []
        for i in range(1, recips.Count + 1):
            r     = recips.Item(i)
            email = getattr(r, "Address", "") or ""
            name  = getattr(r, "Name",    "") or ""
            rtype = getattr(r, "Type",    1)
            if email.lower().startswith("/o="):
                email = _resolve_exchange_email(r) or ""
            if "@" in email:
                attendees.append({"name": name, "email": email, "type": rtype})
        return attendees
    except Exception:
        return []


# ---------------------------------------------------------------------------
# HTML email builder (identical to send_recap.py)
# ---------------------------------------------------------------------------

def build_email_html(data):
    if data["accomplishments"]:
        accomp_items = "\n".join(
            f'<li style="margin-bottom:6px;">{item}</li>'
            for item in data["accomplishments"]
        )
        accomp_block = f"""<hr style="border:none;border-top:1px solid #d0d0d0;margin:18px 0;">
<p style="font-weight:bold;font-size:14px;color:#1f1f1f;margin:0 0 8px 0;">DISCUSSION TOPICS</p>
<ul style="margin:0;padding-left:20px;">
{accomp_items}
</ul>
"""
    else:
        accomp_block = ""

    ai_items = "\n".join(
        f'<li style="margin-bottom:6px;">{item}</li>'
        for item in data["action_items"]
    ) if data["action_items"] else '<li>No action items found.</li>'

    rec_block = ""
    if data["rec_link"] or data["chapters"]:
        parts = [
            '<hr style="border:none;border-top:1px solid #d0d0d0;margin:18px 0;">',
            '<p style="font-weight:bold;font-size:14px;color:#1f1f1f;margin:0 0 8px 0;">RECORDING</p>',
        ]
        if data["rec_link"]:
            parts.append(
                f'<p style="margin:0 0 4px 0;">'
                f'<a href="{data["rec_link"]}" style="color:#0563C1;">&#9654;&nbsp;Watch Recording</a>'
                f'</p>'
            )
        if data["rec_password"]:
            parts.append(
                f'<p style="margin:0 0 12px 0;"><strong>Password:</strong> {data["rec_password"]}</p>'
            )
        if data["chapters"]:
            parts.append('<p style="font-weight:bold;margin:10px 0 6px 0;">Chapters</p>')
            parts.append('<table style="border-collapse:collapse;font-size:13px;">')
            for ch in data["chapters"]:
                parts.append(
                    f'<tr>'
                    f'<td style="padding:2px 14px 2px 0;font-family:monospace;color:#555;">{ch["time"]}</td>'
                    f'<td style="padding:2px 0;">{ch["title"]}</td>'
                    f'</tr>'
                )
            parts.append('</table>')
        rec_block = "\n".join(parts)

    summary = data["overview"] or "(No summary available for this meeting.)"

    agenda_items = "\n".join(
        f'<li style="margin-bottom:6px;">{item}</li>'
        for item in data["action_items"]
    ) if data["action_items"] else '<li></li>'

    return f"""
<div style="font-family:'Aptos',Calibri,Arial,sans-serif;font-size:12pt;color:#1f1f1f;max-width:680px;line-height:1.5;">
<p>Hi Team,</p>
<p>Please find below the recap from <strong>{data['title']}</strong>.</p>

<hr style="border:none;border-top:1px solid #d0d0d0;margin:18px 0;">
<p style="font-weight:bold;font-size:14px;color:#1f1f1f;margin:0 0 8px 0;">SUMMARY</p>
<p style="margin:0;">{summary}</p>

<hr style="border:none;border-top:1px solid #d0d0d0;margin:18px 0;">
<p style="font-weight:bold;font-size:14px;color:#1f1f1f;margin:0 0 8px 0;">ACTION ITEMS</p>
<ul style="margin:0;padding-left:20px;">
{ai_items}
</ul>

{accomp_block}
{rec_block}

<hr style="border:none;border-top:1px solid #d0d0d0;margin:18px 0;">
<p style="font-weight:bold;font-size:14px;color:#1f1f1f;margin:0 0 8px 0;">AGENDA FOR NEXT MEETING</p>
<ul style="margin:0;padding-left:20px;">
{agenda_items}
</ul>
</div>
"""


# ---------------------------------------------------------------------------
# Open one Outlook draft
# ---------------------------------------------------------------------------

def open_draft(outlook, calendar, data, pid=None):
    attendees = get_attendees(calendar, data["title"])
    _cal_emails = {a["email"].lower() for a in attendees}
    for _p in fetch_webex_participants(data["title"], data["date"]):
        if _p["email"] not in _cal_emails:
            attendees.append(_p)
            _cal_emails.add(_p["email"])
    to_list   = [a for a in attendees if a["type"] == 1]
    cc_list   = [a for a in attendees if a["type"] == 2]
    print(f"  Attendees: {len(to_list)} To, {len(cc_list)} CC ({len(attendees)} total incl. Webex participants)")

    subject    = f"Meeting Recap: {data['title']} — {data['date']}"
    inner_html = build_email_html(data)

    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Display(False)

    existing = mail.HTMLBody or ""
    body_tag  = re.search(r"(<body[^>]*>)", existing, re.IGNORECASE)
    if body_tag:
        ins = body_tag.end()
        mail.HTMLBody = existing[:ins] + inner_html + existing[ins:]
    else:
        mail.HTMLBody = (
            '<html><body style="font-family:\'Aptos\',Calibri,Arial,sans-serif;'
            'font-size:12pt;color:#1f1f1f;max-width:680px;line-height:1.5;">'
            + inner_html + "</body></html>"
        )

    if attendees:
        for att in attendees:
            recip = mail.Recipients.Add(att["email"])
            recip.Type = att["type"]
        mail.Recipients.ResolveAll()
        to_names = ", ".join(a["name"] for a in to_list[:4])
        cc_names = ", ".join(a["name"] for a in cc_list[:4])
        if to_names: print(f"  To:        {to_names}{'...' if len(to_list) > 4 else ''}")
        if cc_names: print(f"  CC:        {cc_names}{'...' if len(cc_list) > 4 else ''}")
    else:
        print("  WARNING:   No attendees found — add recipients manually in the draft.")

    print(f"  Draft opened: {subject}")

    # File to SharePoint if PID is known
    if pid:
        try:
            date_str = datetime.strptime(data["date"].strip(), "%A, %B %d, %Y").strftime("%Y-%m-%d")
        except (ValueError, AttributeError):
            date_str = datetime.today().strftime("%Y-%m-%d")
        file_to_sharepoint(inner_html, pid, data["title"], date_str)
    print()


# ---------------------------------------------------------------------------
# Calendar scan — today's [Internal]/[External] meetings I host
# ---------------------------------------------------------------------------

def get_todays_pids_from_calendar(calendar):
    """Return ordered list of PIDs from today's hosted [Internal]/[External] meetings."""
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    today_start = datetime.combine(TARGET_DATE, datetime.min.time()).strftime("%m/%d/%Y 12:00 AM")
    today_end   = datetime.combine(TARGET_DATE, datetime.max.time()).strftime("%m/%d/%Y 11:59 PM")
    try:
        items = items.Restrict(f"[Start] >= '{today_start}' AND [Start] <= '{today_end}'")
    except Exception:
        pass

    seen = set()
    pids = []
    for item in items:
        try:
            subj = getattr(item, "Subject", "") or ""
            org  = (getattr(item, "Organizer", "") or "").lower()

            if not any(p in org for p in ME_PATTERNS):
                continue
            if "[external]" not in subj.lower() and "[internal]" not in subj.lower():
                continue

            match = PID_RE.search(subj)
            if not match:
                print(f"  SKIP (no PID): {subj}")
                continue

            pid = match.group(0)
            if pid not in seen:
                seen.add(pid)
                pids.append(pid)
        except Exception:
            continue
    return pids


# ---------------------------------------------------------------------------
# SharePoint folder finder (mirrors weekly_summary.py)
# ---------------------------------------------------------------------------

def find_pid_folder(pid):
    pid_up = pid.upper()
    for state_dir in SHAREPOINT_BASE.iterdir():
        if not state_dir.is_dir():
            continue
        for project_dir in state_dir.iterdir():
            if project_dir.is_dir() and pid_up in project_dir.name.upper():
                return project_dir
    return None


def file_to_sharepoint(html_body, pid, meeting_title, date_str):
    pid_folder = find_pid_folder(pid)
    if not pid_folder:
        print(f"  WARNING: No SharePoint folder found for {pid} — skipping file.")
        return

    mm_folder = pid_folder / "Meeting Minutes"
    if not mm_folder.exists():
        print(f"  WARNING: No Meeting Minutes folder in {pid_folder.name} — skipping file.")
        return

    safe_title = re.sub(r'[<>:"/\\|?*]', '', meeting_title).strip()[:80]
    filename   = f"{date_str} - {safe_title} - EOD Recap.html"
    filepath   = mm_folder / filename

    full_html = (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        f'<title>{safe_title}</title></head><body>'
        + html_body
        + '</body></html>'
    )
    filepath.write_text(full_html, encoding="utf-8")
    print(f"  Filed:     {filepath.name}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(raw_pids=""):
    global TARGET_DATE

    outlook   = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox     = namespace.GetDefaultFolder(6)
    calendar  = namespace.GetDefaultFolder(9)

    today_str = TARGET_DATE.strftime("%A, %B %d, %Y")
    print(f"\nEnd of Day Recap — {today_str}")

    # Auto-detect from calendar if no PIDs given
    if raw_pids.strip():
        pids = [p.strip() for p in raw_pids.split(",") if p.strip()]
        print(f"PIDs (manual):    {', '.join(pids)}\n")
    else:
        pids = get_todays_pids_from_calendar(calendar)
        if not pids:
            print("No [Internal] or [External] meetings found on today's calendar.")
            return
        print(f"PIDs (calendar):  {', '.join(pids)}\n")

    total_drafts = 0

    for pid in pids:
        print(f"--- PID: {pid} ---")
        all_emails = find_all_webex_emails(inbox, pid)

        if not all_emails:
            print(f"  No Webex recap emails found for {pid} in the last {SEARCH_DAYS} days.\n")
            continue

        today_emails = []
        for msg in all_emails:
            data = parse_email(msg)
            if _is_target_date(data["date"]):
                today_emails.append((msg, data))

        if not today_emails:
            print(f"  Found {len(all_emails)} recap email(s) for {pid}, but none from today ({today_str}).\n")
            continue

        print(f"  Found {len(today_emails)} recap(s) from today for {pid}.\n")

        for i, (msg, data) in enumerate(today_emails):
            print(f"  [{i+1}/{len(today_emails)}] {msg.Subject}")
            print(f"  Title:   {data['title'] or '(not detected)'}")
            print(f"  Summary:         {'Yes' if data['overview'] else 'Not found'}")
            print(f"  Accomplishments: {len(data['accomplishments'])} item(s)")
            print(f"  Actions:         {len(data['action_items'])} item(s)")
            print(f"  Recorded: {'Yes — link found' if data['rec_link'] else 'No'}")

            open_draft(outlook, calendar, data, pid=pid)
            total_drafts += 1

            if i < len(today_emails) - 1:
                time.sleep(1)

        # Pause between PIDs too (only if more PIDs follow)
        if pid != pids[-1]:
            time.sleep(1)

    if total_drafts == 0:
        print("No drafts were opened. Make sure Outlook is open and recap emails have arrived.")
    else:
        print(f"Done. {total_drafts} draft(s) opened in Outlook. Review, edit, and send when ready.")


if __name__ == "__main__":
    args = sys.argv[1:]
    if "--yesterday" in args:
        TARGET_DATE = (datetime.today() - timedelta(days=1)).date()
        args = [a for a in args if a != "--yesterday"]
    elif "--date" in args:
        idx = args.index("--date")
        try:
            TARGET_DATE = datetime.strptime(args[idx + 1], "%Y-%m-%d").date()
            args = args[:idx] + args[idx + 2:]
        except (IndexError, ValueError):
            print("Usage: --date YYYY-MM-DD"); sys.exit(1)
    main(" ".join(args))
