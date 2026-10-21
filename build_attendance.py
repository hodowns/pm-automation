import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

with open(r'C:/Users/hodowns/Desktop/office_hours_attendance_raw.json') as f:
    raw = json.load(f)

sessions_ordered = [
    ('2026-09-11', '9/11/26\nMy Agent\nOffice Hours'),
    ('2026-09-18', '9/18/26\nPhase 1 Closing\nOffice Hours'),
    ('2026-10-02', '10/2/26\nPhase 2:\nUsage'),
    ('2026-10-09', '10/9/26\nPhase 2:\nToken-nomics'),
]

cleaned = {}
for date, _ in sessions_ordered:
    info = raw[date]
    seen = set()
    unique = []
    for p in info['participants']:
        name = p.get('displayName', '').strip()
        email = p.get('email', '').strip()
        if not email or 'Desk Mini' in name or '@' not in email:
            continue
        key = email.lower()
        if key not in seen:
            seen.add(key)
            unique.append({'name': name, 'email': email})
    cleaned[date] = {'title': info['title'], 'participants': unique}

all_attendees = {}
for date, _ in sessions_ordered:
    for p in cleaned[date]['participants']:
        email = p['email'].lower()
        if email not in all_attendees:
            all_attendees[email] = p['name']

def sort_key(item):
    parts = item[1].split()
    return parts[-1].lower() if parts else item[1].lower()

sorted_attendees = sorted(all_attendees.items(), key=sort_key)

CISCO_BLUE = '1565A0'
LIGHT_BLUE = 'D6E4F0'
WHITE = 'FFFFFF'
GREEN = '217346'
LIGHT_GREEN = 'E2EFDA'
GRAY = 'F2F2F2'

thin = Side(border_style='thin', color='BFBFBF')
border = Border(top=thin, left=thin, right=thin, bottom=thin)
header_fill = PatternFill('solid', fgColor=CISCO_BLUE)
header_font = Font(name='Calibri', bold=True, size=10, color=WHITE)

wb = openpyxl.Workbook()

# ── Sheet 1: Attendance Matrix ──
ws = wb.active
ws.title = 'Attendance Matrix'
ws.sheet_view.showGridLines = False

ws.merge_cells('A1:F1')
c = ws['A1']
c.value = 'Claude Cockpit / MyAgent — Office Hours Attendance'
c.font = Font(name='Calibri', bold=True, size=14, color=WHITE)
c.fill = PatternFill('solid', fgColor=CISCO_BLUE)
c.alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[1].height = 28

ws.merge_cells('A2:F2')
s = ws['A2']
s.value = 'Series started September 11, 2026  |  Fridays 12:30 PM ET  |  ' + str(len(sorted_attendees)) + ' unique attendees across 4 sessions'
s.font = Font(name='Calibri', italic=True, size=10, color='595959')
s.fill = PatternFill('solid', fgColor='EEF4FA')
s.alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[2].height = 18

headers = ['Attendee Name', '9/11/26\nMy Agent\nOffice Hours', '9/18/26\nPhase 1 Closing\nOffice Hours', '10/2/26\nPhase 2: Usage', '10/9/26\nPhase 2:\nToken-nomics', 'Sessions\nAttended']
for col_idx, h in enumerate(headers, 1):
    cell = ws.cell(row=3, column=col_idx, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = border
ws.row_dimensions[3].height = 48

for row_idx, (email, name) in enumerate(sorted_attendees, 4):
    is_odd = (row_idx % 2 == 1)
    name_bg = PatternFill('solid', fgColor=GRAY if is_odd else WHITE)

    c = ws.cell(row=row_idx, column=1, value=name)
    c.font = Font(name='Calibri', size=10)
    c.fill = name_bg
    c.alignment = Alignment(vertical='center', indent=1)
    c.border = border

    total = 0
    for col_offset, (date, _) in enumerate(sessions_ordered, 2):
        attended = any(p['email'].lower() == email for p in cleaned[date]['participants'])
        cell = ws.cell(row=row_idx, column=col_offset, value=chr(10003) if attended else '')
        if attended:
            total += 1
            cell.fill = PatternFill('solid', fgColor='C6EFCE' if is_odd else LIGHT_GREEN)
            cell.font = Font(name='Calibri', bold=True, size=12, color=GREEN)
        else:
            cell.fill = name_bg
            cell.font = Font(name='Calibri', size=10, color='BFBFBF')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    count_cell = ws.cell(row=row_idx, column=6, value=total)
    count_cell.fill = PatternFill('solid', fgColor=LIGHT_BLUE if total >= 2 else WHITE)
    count_cell.font = Font(name='Calibri', bold=(total >= 2), size=10)
    count_cell.alignment = Alignment(horizontal='center', vertical='center')
    count_cell.border = border
    ws.row_dimensions[row_idx].height = 18

ws.column_dimensions['A'].width = 30
for col in ['B', 'C', 'D', 'E']:
    ws.column_dimensions[col].width = 14
ws.column_dimensions['F'].width = 10
ws.freeze_panes = 'B4'

# ── Sheets 2-5: Individual Sessions ──
session_meta = [
    ('2026-09-11', 'Claude/My Agent Office Hours (45 Mins) — September 11, 2026', '9-11 My Agent OH'),
    ('2026-09-18', 'Claude/Cockpit Phase 1 Closing Office Hours (45 Mins) — September 18, 2026', '9-18 Ph1 Closing OH'),
    ('2026-10-02', 'Claude/Cockpit Phase 2: Usage (45 Mins) — October 2, 2026', '10-2 Ph2 Usage'),
    ('2026-10-09', 'Claude/Cockpit Phase 2: Token-nomics — October 9, 2026', '10-9 Ph2 Tokennomics'),
]

for date, full_title, sheet_name in session_meta:
    ws2 = wb.create_sheet(title=sheet_name)
    ws2.sheet_view.showGridLines = False

    participants_sorted = sorted(cleaned[date]['participants'], key=lambda p: p['name'].split()[-1].lower())

    ws2.merge_cells('A1:C1')
    t = ws2['A1']
    t.value = full_title
    t.font = Font(name='Calibri', bold=True, size=13, color=WHITE)
    t.fill = PatternFill('solid', fgColor=CISCO_BLUE)
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[1].height = 26

    ws2.merge_cells('A2:C2')
    s2 = ws2['A2']
    s2.value = str(len(participants_sorted)) + ' attendees  |  Fridays 12:30 PM ET  |  Host: Houston Downs'
    s2.font = Font(name='Calibri', italic=True, size=10, color='595959')
    s2.fill = PatternFill('solid', fgColor='EEF4FA')
    s2.alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[2].height = 16

    for col, label in enumerate(['#', 'Name', 'Email'], 1):
        c = ws2.cell(row=3, column=col, value=label)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = border
    ws2.row_dimensions[3].height = 22

    for i, p in enumerate(participants_sorted, 1):
        row = i + 3
        is_odd = (i % 2 == 1)
        bg = PatternFill('solid', fgColor=GRAY if is_odd else WHITE)

        c1 = ws2.cell(row=row, column=1, value=i)
        c1.fill = PatternFill('solid', fgColor=CISCO_BLUE)
        c1.font = Font(name='Calibri', bold=True, size=10, color=WHITE)
        c1.alignment = Alignment(horizontal='center', vertical='center')
        c1.border = border

        c2 = ws2.cell(row=row, column=2, value=p['name'])
        c2.fill = bg
        c2.font = Font(name='Calibri', size=10)
        c2.alignment = Alignment(vertical='center', indent=1)
        c2.border = border

        c3 = ws2.cell(row=row, column=3, value=p['email'])
        c3.fill = bg
        c3.font = Font(name='Calibri', size=10, color='595959')
        c3.alignment = Alignment(vertical='center', indent=1)
        c3.border = border

        ws2.row_dimensions[row].height = 17

    ws2.column_dimensions['A'].width = 5
    ws2.column_dimensions['B'].width = 28
    ws2.column_dimensions['C'].width = 32

out_path = r'C:/Users/hodowns/Desktop/Claude_Cockpit_Office_Hours_Attendance.xlsx'
wb.save(out_path)
print('Saved:', out_path)
print('Unique attendees:', len(sorted_attendees))
