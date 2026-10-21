"""
MCC Billing Agent
Weekly Webex Teams notification — every Monday at 10:30am EST
Scheduled via Windows Task Scheduler
"""

import sys
import calendar
import requests
import openpyxl
from datetime import date, datetime, timedelta

# ── CONFIG — update before first run ────────────────────────────────────────
WEBEX_TOKEN = "MmVmM2NmOGQtODJhNC00ZTJhLWIwNzQtMDc0NmZjNDU4MDNjN2VjM2RiZTktMjI2_PF84_1eb65fdf-9643-417f-9974-ad72cae0e10f"   # paste your Bot Access Token between the quotes
ROOM_ID     = "Y2lzY29zcGFyazovL3VzL1JPT00vZDhlN2ViZjAtODllNS0xMWYxLWExZWMtOGRiOGIxYWU4ZmUw"

MCC_FILE = (
    r"C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS"
    r"\XXXX Houston MCC Tracking\MCC Tracking.xlsx"
)

WINDOW_DAYS = 35

# ── FY27 FISCAL CALENDAR KEY DATES ──────────────────────────────────────────
# Source: FY27 AMER EMEAR GPMO Fiscal Calendar (Updated 11.05.26)
# Dates marked (verify) should be confirmed against the live calendar PDF.
FISCAL_DATES = {
    # JULY 2026 — Q4 FY26 close
    date(2026, 7, 31): "Month End (Q4 FY26 close)",

    # AUGUST 2026 — Q1 FY27 open
    date(2026, 8, 10): "MCC Submission Deadline",
    date(2026, 8, 11): "Country Enablement MCC Deadline",
    date(2026, 8, 12): "OP AS-T Rev Fcst Commit Deadline + Country Enablement MCC Deadline",
    date(2026, 8, 18): "MCC to Customer 5-Day Rule Deadline",
    date(2026, 8, 31): "Month End",

    # SEPTEMBER 2026 — Q1 FY27 mid
    date(2026, 9, 7):  "MCC Submission Deadline",                             # verify
    date(2026, 9, 8):  "Country Enablement MCC Deadline",                     # verify
    date(2026, 9, 9):  "OP AS-T Rev Fcst Commit Deadline",                    # verify
    date(2026, 9, 15): "MCC to Customer 5-Day Rule Deadline",                 # verify
    date(2026, 9, 19): "Cost/Rev Fcst Compliance Deadline",                   # verify
    date(2026, 9, 30): "Month End",

    # OCTOBER 2026 — Q1 FY27 close (end of quarter)
    date(2026, 10, 13): "MCC Submission Deadline",                            # verify
    date(2026, 10, 14): "Country Enablement MCC Deadline",                    # verify
    date(2026, 10, 15): "OP AS-T Rev Fcst Commit Deadline + Country Enablement MCC Deadline",  # verify
    date(2026, 10, 20): "MCC to Customer 5-Day Rule Deadline",                # verify
    date(2026, 10, 22): "Cost/Rev Fcst Compliance Deadline",                  # verify
    date(2026, 10, 31): "Month End (Q1 FY27 close)",

    # NOVEMBER 2026 — Q2 FY27 open
    date(2026, 11, 10): "MCC Submission Deadline",                            # verify
    date(2026, 11, 11): "Country Enablement MCC Deadline",                    # verify
    date(2026, 11, 17): "OP AS-T Rev Fcst Commit Deadline",                   # verify
    date(2026, 11, 30): "Month End",

    # DECEMBER 2026 — Q2 FY27 mid
    date(2026, 12, 8):  "MCC Submission Deadline",                            # verify
    date(2026, 12, 9):  "Country Enablement MCC Deadline",                    # verify
    date(2026, 12, 11): "OP AS-T Rev Fcst Commit Deadline",                   # verify
    date(2026, 12, 31): "Month End",

    # JANUARY 2027 — Q2 FY27 close
    date(2027, 1, 12): "MCC Submission Deadline",                             # verify
    date(2027, 1, 13): "Country Enablement MCC Deadline",                     # verify
    date(2027, 1, 14): "OP AS-T Rev Fcst Commit Deadline + Country Enablement MCC Deadline",  # verify
    date(2027, 1, 31): "Month End (Q2 FY27 close)",
}


# ── MILESTONE FILTER ─────────────────────────────────────────────────────────
def get_flagged_milestones(today: date) -> list:
    wb = openpyxl.load_workbook(MCC_FILE, data_only=True)
    ws = wb["MCC Tracking"]

    window_end         = today + timedelta(days=WINDOW_DAYS)
    fiscal_month_start = today.replace(day=1)
    fiscal_month_end   = today.replace(day=calendar.monthrange(today.year, today.month)[1])

    flagged = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue

        status       = row[0]
        pid          = row[1]
        name         = row[2]
        mcc          = row[3]
        mcc_name     = row[4]
        rec_date_raw = row[5]
        amount       = row[6]

        if isinstance(status, str) and status.startswith("http"):
            status = None
        if status == "Completed" or not mcc:
            continue

        if isinstance(rec_date_raw, datetime):
            rd = rec_date_raw.date()
        elif isinstance(rec_date_raw, date):
            rd = rec_date_raw
        else:
            continue

        days_out       = (rd - today).days
        in_window      = 0 <= days_out <= WINDOW_DAYS
        in_fiscal_month = fiscal_month_start <= rd <= fiscal_month_end

        if (in_window or in_fiscal_month) and days_out >= -15:
            label = f"{name} ({pid})" if name and pid else (name or pid or "Unknown")
            flagged.append({
                "project":   label,
                "mcc":       mcc,
                "milestone": (mcc_name or "")[:55],
                "date":      rd,
                "amount":    amount,
                "days_out":  days_out,
            })

    return sorted(flagged, key=lambda x: x["date"])


# ── FISCAL ALERT FILTER ───────────────────────────────────────────────────────
def get_fiscal_alerts(today: date) -> list:
    window_end = today + timedelta(days=WINDOW_DAYS)
    return [
        {"date": d, "event": ev}
        for d, ev in sorted(FISCAL_DATES.items())
        if today <= d <= window_end
    ]


# ── MESSAGE FORMATTER ─────────────────────────────────────────────────────────
def format_message(flagged: list, fiscal_alerts: list, today: date) -> str:
    lines = [f"## MCC Billing Alert — {today.strftime('%B %d, %Y')}\n"]

    lines.append("### Task 1 – Milestone Notifications")
    if flagged:
        lines += [
            "| Project | MCC | Milestone | Rec. Date | Amount | Days Out |",
            "|---------|-----|-----------|-----------|--------|----------|",
        ]
        for m in flagged:
            amt = f"${m['amount']:,.0f}" if m["amount"] else "—"
            lines.append(
                f"| {m['project']} | {m['mcc']} | {m['milestone']} "
                f"| {m['date'].strftime('%b %d, %Y')} | {amt} | {m['days_out']} |"
            )
    else:
        lines.append("*No milestones within 35-day window.*")

    lines.append("")
    lines.append("### Task 2 – Fiscal Date Alert")
    if fiscal_alerts:
        lines += ["| Date | Event |", "|------|-------|"]
        for a in fiscal_alerts:
            lines.append(f"| **{a['date'].strftime('%b %d')}** | {a['event']} |")
    else:
        lines.append("*No key fiscal dates in next 35 days.*")

    return "\n".join(lines)


# ── WEBEX POST ────────────────────────────────────────────────────────────────
def post_to_webex(message: str) -> int:
    resp = requests.post(
        "https://webexapis.com/v1/messages",
        headers={
            "Authorization": f"Bearer {WEBEX_TOKEN}",
            "Content-Type": "application/json",
        },
        json={"roomId": ROOM_ID, "markdown": message},
        timeout=10,
    )
    resp.raise_for_status()
    return resp.status_code


# ── MAIN ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    today = date.today()

    print(f"[{today}] Scanning MCC Tracking...")
    milestones    = get_flagged_milestones(today)
    fiscal_alerts = get_fiscal_alerts(today)
    message       = format_message(milestones, fiscal_alerts, today)

    print(message)
    print("\nPosting to Webex Teams...")
    status = post_to_webex(message)
    print(f"Done. HTTP {status}")
