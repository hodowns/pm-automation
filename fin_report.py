import os
import sys
import glob
import shutil
import csv
import re
import webbrowser
from datetime import datetime

import win32com.client
import pythoncom

# ── Config ─────────────────────────────────────────────────────────────────────
BLANK_TEMPLATE_DIRS = [
    os.path.expanduser('~/OneDrive - Cisco/Desktop'),
    os.path.expanduser('~/Desktop'),
]
SEARCH_DIRS = [
    os.path.expanduser('~/Downloads'),
    os.path.expanduser('~/OneDrive - Cisco/Desktop'),
    os.path.expanduser('~/Desktop'),
]
DOCCENTAL_URL = 'https://docs.cisco.com/share/page/site/nextgen-edcs/documentlibrary'
NOTIFY_ALIAS = 'sled_financial_report'


# ── File helpers ───────────────────────────────────────────────────────────────

def find_blank_template():
    """Look for a file named *BLANK* in Desktop dirs first."""
    for d in BLANK_TEMPLATE_DIRS:
        if not os.path.isdir(d):
            continue
        for pattern in ['*BLANK*.xlsx', '*blank*.xlsx', '*Blank*.xlsx', '*template*.xlsx', '*Template*.xlsx']:
            matches = [f for f in glob.glob(os.path.join(d, pattern)) if os.path.isfile(f)]
            if matches:
                matches.sort(key=os.path.getmtime, reverse=True)
                return matches[0]
    return None


def find_latest_template():
    """Fall-back: most recent Financial Report Excel in search dirs."""
    candidates = []
    for d in SEARCH_DIRS:
        if os.path.isdir(d):
            candidates += glob.glob(os.path.join(d, 'Financial Report*SLED*.xlsx'))
            candidates += glob.glob(os.path.join(d, 'Financial Report*USPS*.xlsx'))
    candidates = [c for c in candidates if os.path.isfile(c)]
    if not candidates:
        return None
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates[0]


def find_recent_csvs():
    d = os.path.expanduser('~/Downloads')
    csvs = glob.glob(os.path.join(d, '*.csv'))
    csvs.sort(key=os.path.getmtime, reverse=True)
    return csvs[:5]


def prompt(label, default=None):
    if default:
        answer = input(f'{label}\n  [{os.path.basename(default)}]  Enter to accept: ').strip().strip('"\'')
        return answer or default
    return input(f'{label}: ').strip().strip('"\'')


def read_csv(path):
    csv.field_size_limit(10_000_000)
    for enc in ('utf-8-sig', 'utf-16', 'latin-1'):
        try:
            with open(path, newline='', encoding=enc) as f:
                sample = f.read(8192)
                f.seek(0)
                delimiter = '\t' if sample.count('\t') > sample.count(',') else ','
                return list(csv.reader(f, delimiter=delimiter))
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError(f'Could not decode {path}')


def rows_to_array(rows):
    """Pad rows to uniform width; return as list of tuples for Excel bulk-write."""
    if not rows:
        return []
    width = max(len(r) for r in rows)
    return [tuple(r + [''] * (width - len(r))) for r in rows]


# ── Excel operations ───────────────────────────────────────────────────────────

def clear_from(ws, start_row, start_col):
    """Clear everything from (start_row, start_col) to the end of the sheet's used range."""
    ur = ws.UsedRange
    if ur is None:
        return
    last_row = ur.Row + ur.Rows.Count - 1
    last_col = ur.Column + ur.Columns.Count - 1
    if last_row >= start_row and last_col >= start_col:
        ws.Range(
            ws.Cells(start_row, start_col),
            ws.Cells(last_row, last_col)
        ).ClearContents()


def paste_gppr(wb, csv_path):
    """Skip first row (Tableau header) and last row (Grand Total); paste at O2 (col 15).
       Source: all 103 columns A-CY from CSV → template O2:DM."""
    ws = wb.Worksheets('GPPR')
    rows = read_csv(csv_path)
    if len(rows) < 3:
        print('  GPPR CSV too short — skipping paste')
        return
    data_rows = rows[1:-1]            # drop Tableau header row + Grand Total row
    data = rows_to_array(data_rows)
    ncols = len(data[0])

    # Full clear from O2 outward — eliminates any overhang from previous week
    clear_from(ws, 2, 15)

    # Bulk-write starting at O2
    ws.Range(ws.Cells(2, 15), ws.Cells(1 + len(data), 14 + ncols)).Value = data
    print(f'  GPPR: cleared previous data, pasted {len(data)} rows at O2')


def paste_expenditure(wb, csv_path):
    """Clear entire Expenditure tab, paste full CSV at A1, delete last row (Grand Total)."""
    ws = wb.Worksheets('Expenditure')
    rows = read_csv(csv_path)
    if not rows:
        print('  Expenditure CSV empty — skipping paste')
        return
    data = rows_to_array(rows)
    ncols = len(data[0])

    # Full clear — no overhang possible
    ws.Cells.ClearContents()
    ws.Range(ws.Cells(1, 1), ws.Cells(len(data), ncols)).Value = data

    # Delete last row (Grand Total)
    last_row = ws.UsedRange.Rows.Count
    ws.Rows(last_row).Delete()
    print(f'  Expenditure: cleared tab, pasted {len(rows)} rows, deleted Grand Total row')


def refresh_all_pivots(wb):
    for ws in wb.Worksheets:
        try:
            for pt in ws.PivotTables():
                pt.RefreshTable()
        except Exception:
            pass
    print('  All pivot tables refreshed')


def fix_labor_summary(wb):
    """Deselect (blank) in EXPENDITURE_TYPE column field."""
    ws = wb.Worksheets('Labor Summary')
    pt = ws.PivotTables('PivotTable3')
    pf = pt.PivotFields('EXPENDITURE_TYPE')
    for pi in pf.PivotItems():
        if pi.Name.strip() in ['(blank)', '']:
            try:
                pi.Visible = False
            except Exception:
                pass
    print('  Labor Summary: deselected (blank) in EXPENDITURE_TYPE')


def fix_labor_monthly(wb):
    """Fiscal Year. → sort Z to A, deselect (blank)/#N/A.
       Fiscal Month # / Fiscal Month Name → sort Z to A, keep Select All."""
    ws = wb.Worksheets('Labor Monthly')
    pt = ws.PivotTables('PivotTable1')

    pf_year = pt.PivotFields('Fiscal Year.')
    pf_year.AutoSort(2, 'Fiscal Year.')            # 2 = xlDescending
    for pi in pf_year.PivotItems():
        if pi.Name.strip() in ['(blank)', '#N/A', '']:
            try:
                pi.Visible = False
            except Exception:
                pass

    for fname in ['Fiscal Month #', 'Fiscal Month Name']:
        try:
            pf = pt.PivotFields(fname)
            pf.AutoSort(2, fname)
            # Ensure all items visible (Select All)
            for pi in pf.PivotItems():
                pi.Visible = True
        except Exception:
            pass

    print('  Labor Monthly: Fiscal Year sorted Z to A, blanks deselected; Fiscal Month sorted')


def update_instructions_date(wb, report_date):
    ws = wb.Worksheets('Instructions')
    ws.Range('A8').Value = report_date.strftime('%m/%d/%Y')
    print(f'  Instructions A8: {report_date.strftime("%m/%d/%Y")}')


def hide_tabs(wb):
    for name in ['Instructions', 'Expenditure', 'GPPR']:
        try:
            wb.Worksheets(name).Visible = False
        except Exception as e:
            print(f'  Warning — could not hide {name}: {e}')
    print('  Hidden: Instructions, Expenditure, GPPR')


# ── Outlook draft ──────────────────────────────────────────────────────────────

def create_outlook_draft(doc_link, report_date):
    ol = win32com.client.Dispatch('Outlook.Application')
    mail = ol.CreateItem(0)
    mail.To = NOTIFY_ALIAS
    mail.Subject = f'Financial Report - USPS SLED - {report_date.strftime("%m-%d-%Y")}'

    # Display first so Outlook inserts the default signature block
    mail.Display()

    msg_html = (
        '<p style="font-family:Aptos,Calibri,sans-serif;font-size:12pt;margin:0 0 12pt 0;">'
        'Hello SLED Team,</p>'
        '<p style="font-family:Aptos,Calibri,sans-serif;font-size:12pt;margin:0 0 12pt 0;">'
        "This week&#39;s report has been uploaded to Doc Central. "
        'The report can be accessed by clicking on the download icon on the right. '
        f'Here is the direct link: <a href="{doc_link}">{doc_link}</a></p>'
        '<p style="font-family:Aptos,Calibri,sans-serif;font-size:12pt;margin:0 0 12pt 0;">'
        'Regards,<br>Houston</p>'
    )

    # Inject message at start of body, preserving existing signature
    existing = mail.HTMLBody
    match = re.search(r'(<body[^>]*>)', existing, re.IGNORECASE)
    if match:
        mail.HTMLBody = existing[:match.end()] + msg_html + existing[match.end():]
    else:
        mail.HTMLBody = msg_html + existing

    print(f'  Outlook draft opened  →  To: {NOTIFY_ALIAS}')


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--gppr',        help='Path to GPPR CSV')
    parser.add_argument('--exp',         help='Path to Expenditure CSV')
    parser.add_argument('--link',        help='DocCentral direct link for email draft')
    parser.add_argument('--draft-only',  action='store_true',
                        help='Skip Excel — just open the Outlook draft (requires --link)')
    parser.add_argument('--auto',        action='store_true',
                        help='No prompts — auto-select most recent CSVs (Task Scheduler mode)')
    args = parser.parse_args()

    today    = datetime.today()
    date_str = today.strftime('%m-%d-%Y')

    # ── Draft-only shortcut (Phase 2 of Claude-orchestrated run) ─────────────────
    if args.draft_only:
        if not args.link:
            print('ERROR: --draft-only requires --link')
            sys.exit(1)
        print(f'=== USPS SLED Financial Report — {date_str} ===')
        print('\n--- Outlook Draft ---')
        create_outlook_draft(args.link, today)
        print(f'\n=== Draft opened — review and send to {NOTIFY_ALIAS} ===')
        return

    print(f'=== USPS SLED Financial Report — {date_str} ===\n')
    if args.auto:
        print('[AUTO MODE — no prompts]\n')

    # 1. Locate template — blank template on Desktop takes priority
    blank = find_blank_template()
    if blank:
        print(f'Blank template found: {os.path.basename(blank)}')
        template = blank
    else:
        print('No blank template found on Desktop — falling back to most recent report.')
        template = find_latest_template()
        if template:
            print(f'Found: {os.path.basename(template)}')
            if not args.auto and not args.gppr:
                if input('Use as this week\'s template? [Y/n]: ').strip().lower() == 'n':
                    template = prompt('Template Excel path')
        else:
            if args.auto or args.gppr:
                print('ERROR: No template found.')
                sys.exit(1)
            template = prompt('Template Excel path')

    # 2. Create new dated report file
    out_dir  = os.path.dirname(template)
    new_name = f'Financial Report USPS SLED {date_str}.xlsx'
    new_path = os.path.join(out_dir, new_name)
    shutil.copy2(template, new_path)
    print(f'Created: {new_name}\n')

    # 3. CSV selection
    recent = find_recent_csvs()
    if args.gppr:
        gppr_csv = args.gppr
        print(f'GPPR:        {os.path.basename(gppr_csv)}')
    elif args.auto:
        if not recent:
            print('ERROR: No CSVs found in Downloads.')
            sys.exit(1)
        gppr_csv = recent[0]
        print(f'GPPR (auto): {os.path.basename(gppr_csv)}')
    else:
        print('--- Tableau CSV Files ---')
        print('Download both CSVs from Tableau before continuing.\n')
        gppr_csv = prompt('GPPR CSV  (AS Project Financial Summary — first+last rows auto-skipped)',
                          recent[0] if recent else None)

    if args.exp:
        exp_csv = args.exp
        print(f'Expenditure: {os.path.basename(exp_csv)}\n')
    elif args.auto:
        if len(recent) < 2:
            print('ERROR: Need at least 2 CSVs in Downloads for auto mode.')
            sys.exit(1)
        exp_csv = recent[1]
        print(f'Expenditure (auto): {os.path.basename(exp_csv)}\n')
    else:
        exp_csv = prompt('Expenditure CSV  (PETR download — Grand Total row auto-deleted)',
                         recent[1] if len(recent) > 1 else None)

    # 4. Excel processing
    print('\n--- Excel Processing ---')
    pythoncom.CoInitialize()
    xl = win32com.client.Dispatch('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Open(os.path.abspath(new_path))
        paste_gppr(wb, gppr_csv)
        paste_expenditure(wb, exp_csv)
        refresh_all_pivots(wb)
        fix_labor_summary(wb)
        fix_labor_monthly(wb)
        update_instructions_date(wb, today)
        hide_tabs(wb)
        wb.Save()
        print(f'\nSaved: {new_name}')
    except Exception as e:
        print(f'\nExcel error: {e}')
        if wb:
            try:
                wb.Close(False)
            except Exception:
                pass
        xl.Quit()
        sys.exit(1)
    finally:
        if wb:
            try:
                wb.Close(False)
            except Exception:
                pass
        try:
            xl.Quit()
        except Exception:
            pass

    # 4.5 HTML dashboard — Houston only; skipped for other users
    if os.getenv('USERNAME', '').lower() == 'hodowns':
        try:
            import importlib.util
            _spec = importlib.util.spec_from_file_location(
                'gen_html',
                os.path.join(os.path.dirname(__file__), 'generate_fin_report_html.py')
            )
            _mod = importlib.util.module_from_spec(_spec)
            _spec.loader.exec_module(_mod)
            print('\n--- HTML Dashboard ---')
            _mod.generate(new_path)
        except Exception as _he:
            print(f'  HTML dashboard skipped: {_he}')

    # 5. DocCentral upload  (TODO: Playwright full-automation once Tableau access granted)
    print('\n--- DocCentral Upload ---')
    print(f'  1. Move last week\'s report → Archives folder')
    print(f'  2. Upload:  {new_name}')
    print(f'  3. Flag as "Cisco Highly Confidential"')
    print(f'  4. Right-click filename → Copy Link Location')
    webbrowser.open(DOCCENTAL_URL)

    if args.link:
        doc_link = args.link
        print(f'\n--- Outlook Draft ---')
        create_outlook_draft(doc_link, today)
        print(f'\n=== Done!  Review draft and send to {NOTIFY_ALIAS} ===')
    elif args.auto:
        doc_link = '[INSERT DOCCENTAL LINK]'
        create_outlook_draft(doc_link, today)
        print('  Auto mode: DocCentral opened — update placeholder link in draft before sending.')
        print(f'\n=== Done!  Review draft and send to {NOTIFY_ALIAS} ===')
    elif args.gppr:
        print(f'\nExcel file ready: {new_name}')
        print('DocCentral is open in your browser.')
        print('Upload the file, copy the link, then paste it here in Claude to generate the email draft.')
    else:
        doc_link = input('\nPaste the DocCentral link here: ').strip()
        create_outlook_draft(doc_link, today)
        print(f'\n=== Done!  Review draft and send to {NOTIFY_ALIAS} ===')


if __name__ == '__main__':
    main()
