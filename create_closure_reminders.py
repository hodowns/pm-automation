import win32com.client
from datetime import datetime

projects = [
    ("C19602", "2027-06-14"),
    ("E55085", "2026-10-01"),
    ("F51119", "2027-09-18"),
    ("F33607", "2027-10-20"),
    ("F33347", "2026-10-20"),
    ("F31542", "2026-11-21"),
    ("C16507", "2029-01-06"),
    ("F47009", "2027-08-04"),
    ("F43673", "2027-08-07"),
    ("F22532", "2026-11-04"),
    ("C10576", "2027-07-05"),
    ("F29426", "2031-08-29"),
    ("F27174", "2026-07-31"),
    ("F48852", "2027-10-03"),
    ("F48861", "2027-10-03"),
    ("F40696", "2027-05-14"),
    ("F46100", "2027-08-12"),
    ("F31176", "2026-12-24"),
    ("F18212", "2027-09-22"),
]

today = datetime.today().date()
outlook = win32com.client.Dispatch("Outlook.Application")

created = []
skipped = []

for pid, end_date_str in projects:
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()
    if end_date < today:
        skipped.append((pid, end_date_str, "past"))
        continue

    appt = outlook.CreateItem(1)  # olAppointmentItem
    appt.Subject = f"Project closure: {pid}"
    appt.Start = f"{end_date_str} 07:00"
    appt.Duration = 30
    appt.ReminderSet = True
    appt.ReminderMinutesBeforeStart = 1440  # 1-day reminder
    appt.BusyStatus = 0  # olFree
    appt.Save()

    created.append((pid, end_date_str))
    print(f"  Created: Project closure: {pid} on {end_date_str}")

print(f"\nDone. {len(created)} reminders created.")
if skipped:
    print(f"Skipped {len(skipped)} past-dated projects:")
    for pid, d, reason in skipped:
        print(f"  {pid} ({d}) — {reason}")
