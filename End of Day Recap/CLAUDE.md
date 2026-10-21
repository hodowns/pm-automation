# End of Day Recap Tool

When asked to send an end of day recap, run:

```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\End of Day Recap\send_eod_recap.py"
```

No arguments needed — the script auto-detects today's meetings from Outlook.

## How to trigger
- "Send end of day recap"
- "Send end of day recap for External and Internal calls"
- "EOD recap"

## Manual PID override (optional)
If you want to target specific PIDs instead of auto-detecting:
```
python -u -X utf8 "C:\Users\hodowns\Claude Projects VS CODE\End of Day Recap\send_eod_recap.py" "E55085, F51119"
```

## How it works
1. Scans today's Outlook calendar for meetings where you are the organizer and the subject contains `[Internal]` or `[External]`
2. Extracts the PID from each meeting subject (e.g. F51119, E55085)
3. Searches the last 14 days of inbox for Webex recap emails (`messenger@webex.com`) matching each PID, filtered to today's date
4. Opens a separate pre-filled Outlook draft per meeting (pops open in its own window)
5. Files the recap HTML to the PID's `Meeting Minutes` folder in SharePoint sync

## SharePoint filing
- Base: `C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS`
- Filed as: `Meeting Minutes\YYYY-MM-DD - [Meeting Title] - EOD Recap.html`
- Uses the same `find_pid_folder()` logic as `weekly_summary.py` and `profile_report.py`

## Requirements
- Outlook must be open
- Python installed with: `pywin32`, `beautifulsoup4`

## Notes
- Filters by meeting date inside the Webex email body, not email received date — late-arriving recap emails are still caught
- If no recap email is found for a meeting, that meeting is reported and skipped
- PIDs with no matching SharePoint folder are flagged but the draft still opens
