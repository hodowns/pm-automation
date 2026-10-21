#!/usr/bin/env python3
"""
Extract Outlook calendar items containing 'F51119' and '[External]' in the subject
for the current month and output as JSON.
"""

import json
import sys
from datetime import datetime
from zoneinfo import ZoneInfo

import win32com.client

YEAR  = 2026
MONTH = 10
LOCAL_TZ = ZoneInfo("America/Chicago")

def run():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns      = outlook.GetNamespace("MAPI")
    cal     = ns.GetDefaultFolder(9)   # olFolderCalendar = 9
    items   = cal.Items

    items.IncludeRecurrences = True
    items.Sort("[Start]")

    # Date range: whole month
    start_str = f"10/01/2026 00:00 AM"
    end_str   = f"10/31/2026 11:59 PM"
    restriction = (
        f"[Start] >= '{start_str}' AND [Start] <= '{end_str}'"
    )
    filtered = items.Restrict(restriction)

    results = []
    for item in filtered:
        try:
            subj = item.Subject or ""
            if "F51119" not in subj:
                continue
            if "[External]" not in subj and "EXTERNAL" not in subj.upper():
                continue

            start_dt = item.Start
            end_dt   = item.End

            # Convert COM date to python datetime string
            start_iso = start_dt.strftime("%Y-%m-%dT%H:%M:%S") if hasattr(start_dt, "strftime") else str(start_dt)
            end_iso   = end_dt.strftime("%Y-%m-%dT%H:%M:%S")   if hasattr(end_dt,   "strftime") else str(end_dt)

            # Duration in minutes
            try:
                duration = item.Duration
            except Exception:
                duration = 60

            # Organizer
            try:
                organizer = item.Organizer or ""
            except Exception:
                organizer = ""

            # Required attendees (string list)
            try:
                required = item.RequiredAttendees or ""
            except Exception:
                required = ""

            # Location
            try:
                location = item.Location or ""
            except Exception:
                location = ""

            # Body preview (first 200 chars)
            try:
                body_preview = (item.Body or "")[:200].strip().replace("\r\n", " ").replace("\n", " ")
            except Exception:
                body_preview = ""

            results.append({
                "subject":    subj,
                "start":      start_iso,
                "end":        end_iso,
                "duration":   duration,
                "organizer":  organizer,
                "required":   required,
                "location":   location,
                "body":       body_preview,
            })
        except Exception as e:
            sys.stderr.write(f"Error reading item: {e}\n")

    print(json.dumps(results, indent=2))

if __name__ == "__main__":
    run()
