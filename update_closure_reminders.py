import win32com.client
from datetime import datetime
from dateutil.relativedelta import relativedelta

projects = [
    ("C19602", "2027-06-14"),
    ("F51119", "2027-09-18"),
    ("F33607", "2027-10-20"),
    ("F31542", "2026-11-21"),
    ("C16507", "2029-01-06"),
    ("F47009", "2027-08-04"),
    ("F43673", "2027-08-07"),
    ("F22532", "2026-11-04"),
    ("C10576", "2027-07-05"),
    ("F29426", "2031-08-29"),
    ("F48852", "2027-10-03"),
    ("F48861", "2027-10-03"),
    ("F40696", "2027-05-14"),
    ("F46100", "2027-08-12"),
    ("F31176", "2026-12-24"),
    ("F18212", "2027-09-22"),
]

# Build lookup: subject -> new date (end_date - 2 months)
new_dates = {}
for pid, end_date_str in projects:
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
    new_date = end_date - relativedelta(months=2)
    new_dates[f"Project closure: {pid}"] = new_date

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar
items = calendar.Items

updated = []
not_found = []

for subject, new_date in new_dates.items():
    restriction = f"[Subject] = '{subject}'"
    filtered = items.Restrict(restriction)
    found = False
    for item in filtered:
        try:
            item.Start = new_date.strftime("%Y-%m-%d 07:00")
            item.Save()
            updated.append((subject, new_date.strftime("%Y-%m-%d")))
            print(f"  Updated: {subject} → {new_date.strftime('%Y-%m-%d')}")
            found = True
            break
        except Exception:
            continue
    if not found:
        not_found.append(subject)

print(f"\nDone. {len(updated)} updated.")
if not_found:
    print(f"Not found ({len(not_found)}):")
    for s in not_found:
        print(f"  {s}")
