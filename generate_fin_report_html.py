"""
generate_fin_report_html.py
Reads Financial Report USPS SLED Excel and produces a standalone interactive HTML dashboard.
Called automatically at end of fin_report.py, or run directly.
"""

import os, sys, json, glob, re
from datetime import datetime, date
import win32com.client, pythoncom


# ── Value helpers ──────────────────────────────────────────────────────────────

def _is_xl_error(v):
    return isinstance(v, int) and v < -1_000_000_000

def _fmt(v):
    if v is None or _is_xl_error(v):
        return ''
    if isinstance(v, (datetime, date)):
        try:
            return v.strftime('%m/%d/%Y')
        except Exception:
            return str(v)
    if isinstance(v, float):
        return int(v) if v == int(v) else round(v, 4)
    if isinstance(v, int):
        return v
    return str(v).strip()


# ── Header readers ─────────────────────────────────────────────────────────────

def _read_row(ws, row, ncols):
    vals = []
    for c in range(1, ncols + 1):
        v = ws.Cells(row, c).Value
        vals.append('' if (v is None or _is_xl_error(v)) else str(v).strip())
    return vals


def read_headers_simple(ws, row, scan_cols=300):
    last = 0
    for c in range(1, scan_cols + 1):
        v = ws.Cells(row, c).Value
        if v is not None and not _is_xl_error(v) and str(v).strip():
            last = c
    if last == 0:
        return []
    return _read_row(ws, row, last)


def read_headers_pivot(ws, group_rows, label_row, scan_cols=300):
    """Build composite column names from multi-level pivot headers."""
    # Find last meaningful label column
    labels_raw = _read_row(ws, label_row, scan_cols)
    last = max((i for i, h in enumerate(labels_raw) if h.strip()), default=-1) + 1
    if last == 0:
        return []

    # Forward-fill each group row (merged cells report None in continuation cells)
    groups = []
    for grow in group_rows:
        row, carry = [], ''
        for c in range(1, last + 1):
            v = ws.Cells(grow, c).Value
            if v is not None and not _is_xl_error(v):
                s = str(v).strip()
                if s:
                    carry = s
            row.append(carry)
        groups.append(row)

    labels = labels_raw[:last]
    skip_group_words = {'Values', 'EXPENDITURE_TYPE', 'Fiscal Year.', 'Fiscal Month #',
                        'Fiscal Month Name', 'Fiscal Month', 'Fiscal Year'}

    headers, seen = [], {}
    for c in range(last):
        lbl = labels[c]
        prefix = ''
        for g in groups:
            gv = g[c]
            if gv and gv not in skip_group_words and gv != lbl and not re.fullmatch(r'\d+\.?0?', gv):
                prefix = gv
                break

        h = f'{prefix} — {lbl}' if (prefix and lbl) else (lbl or prefix or f'Col{c+1}')
        if h in seen:
            seen[h] += 1
            h = f'{h} ({seen[h]})'
        else:
            seen[h] = 1
        headers.append(h)

    return headers


# ── Data reader ────────────────────────────────────────────────────────────────

def read_data(ws, headers, data_start, empty_streak_limit=5):
    if not headers:
        return []
    ncols = len(headers)

    ur = ws.UsedRange
    last_row = ur.Row + ur.Rows.Count - 1
    if data_start > last_row:
        return []

    try:
        rng = ws.Range(ws.Cells(data_start, 1), ws.Cells(last_row, ncols))
        raw = rng.Value
    except Exception as e:
        print(f'  Warning reading data: {e}')
        return []

    if raw is None:
        return []
    if not isinstance(raw[0], (tuple, list)):
        raw = (raw,)

    rows, streak = [], 0
    for row in raw:
        vals = [_fmt(v) for v in row[:ncols]]
        if all(v == '' for v in vals):
            streak += 1
            if streak >= empty_streak_limit:
                break
            continue
        streak = 0
        rows.append(vals)
    return rows


# ── HTML template ──────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="en">
<title>USPS SLED Fin Report</title>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">

<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">

<style>
/* ── Tokens ─────────────────────────────────────────────────── */
:root {
  color-scheme: light;
  --bg:        #F0F4F8;
  --surface:   #FFFFFF;
  --surface2:  #F8FAFC;
  --border:    #D1DAE5;
  --border2:   #E4EAF1;
  --text:      #1A2740;
  --text2:     #4A607A;
  --text3:     #7A90A8;
  --accent:    #0284C7;
  --accent-bg: #EFF8FF;
  --tab-act:   #0284C7;
  --tab-text:  #4A607A;
  --header-bg: #1E3A5F;
  --header-fg: #FFFFFF;
  --row-alt:   #F5F8FC;
  --pos:       #059669;
  --neg:       #DC2626;
  --shadow:    0 1px 3px rgba(0,0,0,.10), 0 1px 8px rgba(0,0,0,.06);
}

@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    color-scheme: dark;
    --bg:        #080E1A;
    --surface:   #0F1826;
    --surface2:  #131E2E;
    --border:    #1E2D42;
    --border2:   #243348;
    --text:      #E2E8F2;
    --text2:     #90A8C0;
    --text3:     #5A7590;
    --accent:    #38BDF8;
    --accent-bg: #0C2036;
    --tab-act:   #38BDF8;
    --tab-text:  #7090AC;
    --header-bg: #0A1A30;
    --header-fg: #B0CCE4;
    --row-alt:   #0C1624;
    --pos:       #34D399;
    --neg:       #F87171;
    --shadow:    0 1px 3px rgba(0,0,0,.4);
  }
}

:root[data-theme="dark"] {
  color-scheme: dark;
  --bg:        #080E1A;
  --surface:   #0F1826;
  --surface2:  #131E2E;
  --border:    #1E2D42;
  --border2:   #243348;
  --text:      #E2E8F2;
  --text2:     #90A8C0;
  --text3:     #5A7590;
  --accent:    #38BDF8;
  --accent-bg: #0C2036;
  --tab-act:   #38BDF8;
  --tab-text:  #7090AC;
  --header-bg: #0A1A30;
  --header-fg: #B0CCE4;
  --row-alt:   #0C1624;
  --pos:       #34D399;
  --neg:       #F87171;
  --shadow:    0 1px 3px rgba(0,0,0,.4);
}

/* ── Reset / base ───────────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; }
html, body { height: 100%; margin: 0; }
body {
  background: var(--bg);
  color: var(--text);
  font-family: 'Inter', system-ui, sans-serif;
  font-size: 13px;
  line-height: 1.5;
  display: flex;
  flex-direction: column;
}

/* ── Header bar ─────────────────────────────────────────────── */
.app-header {
  background: var(--header-bg);
  color: var(--header-fg);
  padding: 14px 20px 12px;
  display: flex;
  align-items: baseline;
  gap: 16px;
  flex-shrink: 0;
  box-shadow: 0 2px 6px rgba(0,0,0,.3);
}
.app-header h1 {
  font-family: 'Inter', sans-serif;
  font-size: 15px;
  font-weight: 600;
  letter-spacing: .02em;
  margin: 0;
  color: #FFFFFF;
}
.app-header .week-badge {
  font-family: 'IBM Plex Mono', monospace;
  font-size: 11px;
  color: #78A8CC;
  letter-spacing: .04em;
}
.theme-btn {
  margin-left: auto;
  background: none;
  border: 1px solid #2A4060;
  border-radius: 4px;
  color: #78A8CC;
  font-size: 11px;
  padding: 3px 8px;
  cursor: pointer;
  font-family: 'Inter', sans-serif;
}
.theme-btn:hover { background: rgba(255,255,255,.05); }

/* ── Tab nav ─────────────────────────────────────────────────── */
.tab-nav {
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  display: flex;
  gap: 0;
  padding: 0 16px;
  flex-shrink: 0;
  overflow-x: auto;
}
.tab-btn {
  padding: 10px 18px;
  font-size: 12px;
  font-weight: 500;
  color: var(--tab-text);
  background: none;
  border: none;
  border-bottom: 2px solid transparent;
  cursor: pointer;
  white-space: nowrap;
  font-family: 'Inter', sans-serif;
  letter-spacing: .02em;
  transition: color .15s, border-color .15s;
}
.tab-btn:hover { color: var(--text); }
.tab-btn.active {
  color: var(--tab-act);
  border-bottom-color: var(--tab-act);
}
.tab-count {
  display: inline-block;
  background: var(--border);
  color: var(--text2);
  font-size: 10px;
  border-radius: 10px;
  padding: 0 6px;
  margin-left: 6px;
  font-family: 'IBM Plex Mono', monospace;
}
.tab-btn.active .tab-count {
  background: var(--accent-bg);
  color: var(--accent);
}

/* ── Panel ───────────────────────────────────────────────────── */
.panel { display: none; flex: 1; flex-direction: column; overflow: hidden; }
.panel.active { display: flex; }

.panel-toolbar {
  padding: 10px 16px 8px;
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
  background: var(--surface2);
  border-bottom: 1px solid var(--border2);
  flex-shrink: 0;
}
.panel-toolbar label {
  font-size: 11px;
  color: var(--text2);
  font-weight: 500;
  letter-spacing: .03em;
  text-transform: uppercase;
}
.search-box {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 5px;
  color: var(--text);
  font-size: 12px;
  font-family: 'Inter', sans-serif;
  padding: 5px 10px;
  outline: none;
  width: 240px;
  max-width: 100%;
}
.search-box:focus { border-color: var(--accent); }
.search-box::placeholder { color: var(--text3); }
.row-count {
  font-family: 'IBM Plex Mono', monospace;
  font-size: 11px;
  color: var(--text3);
  margin-left: auto;
}

/* DataTables button overrides */
.dt-buttons { display: flex; gap: 6px; }
.dt-button {
  background: var(--surface) !important;
  border: 1px solid var(--border) !important;
  border-radius: 4px !important;
  color: var(--text2) !important;
  font-size: 11px !important;
  font-family: 'Inter', sans-serif !important;
  padding: 4px 10px !important;
  cursor: pointer !important;
  box-shadow: none !important;
}
.dt-button:hover {
  background: var(--accent-bg) !important;
  color: var(--accent) !important;
  border-color: var(--accent) !important;
}
.dt-button-collection {
  background: var(--surface) !important;
  border: 1px solid var(--border) !important;
  border-radius: 6px !important;
  box-shadow: var(--shadow) !important;
  padding: 4px 0 !important;
  max-height: 360px;
  overflow-y: auto;
}
.dt-button-collection .dt-button {
  display: block !important;
  width: 100% !important;
  border: none !important;
  border-radius: 0 !important;
  text-align: left !important;
  padding: 6px 14px !important;
}
.dt-button-collection .dt-button.active {
  background: var(--accent-bg) !important;
  color: var(--accent) !important;
}

/* ── Table wrapper ───────────────────────────────────────────── */
.table-wrap {
  flex: 1;
  overflow: auto;
  padding: 0;
}

/* DataTables core overrides */
.dataTables_wrapper { padding: 0; }
.dataTables_filter,
.dataTables_length,
.dataTables_info { display: none; }  /* we handle these ourselves */

.dataTables_paginate {
  padding: 8px 16px;
  display: flex;
  justify-content: flex-end;
  gap: 4px;
  background: var(--surface2);
  border-top: 1px solid var(--border);
  flex-shrink: 0;
}
.paginate_button {
  background: none !important;
  border: 1px solid var(--border) !important;
  border-radius: 4px !important;
  color: var(--text2) !important;
  font-size: 11px !important;
  font-family: 'IBM Plex Mono', monospace !important;
  padding: 3px 8px !important;
  cursor: pointer !important;
  min-width: 28px !important;
  text-align: center;
}
.paginate_button.current {
  background: var(--accent-bg) !important;
  color: var(--accent) !important;
  border-color: var(--accent) !important;
}
.paginate_button:hover:not(.current):not(.disabled) {
  background: var(--surface) !important;
  color: var(--text) !important;
}
.paginate_button.disabled { opacity: .35 !important; cursor: default !important; }

/* ── Table styling ───────────────────────────────────────────── */
table.fin-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 12px;
}
table.fin-table thead th {
  background: var(--header-bg);
  color: var(--header-fg);
  font-family: 'Inter', sans-serif;
  font-size: 11px;
  font-weight: 500;
  letter-spacing: .04em;
  text-transform: uppercase;
  padding: 8px 10px 6px;
  border-right: 1px solid rgba(255,255,255,.06);
  white-space: nowrap;
  position: sticky;
  top: 0;
  z-index: 2;
  cursor: pointer;
  user-select: none;
}
table.fin-table thead th:last-child { border-right: none; }
table.fin-table thead th.dt-ordering-asc::after { content: ' ↑'; opacity: .7; }
table.fin-table thead th.dt-ordering-desc::after { content: ' ↓'; opacity: .7; }

/* Column filter row */
table.fin-table thead tr.filter-row th {
  background: var(--surface2);
  padding: 4px 6px;
  position: sticky;
  top: 33px;
  z-index: 2;
  border-top: 1px solid var(--border);
  border-bottom: 1px solid var(--border);
}
table.fin-table thead tr.filter-row th input {
  width: 100%;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 3px;
  color: var(--text);
  font-size: 11px;
  font-family: 'Inter', sans-serif;
  padding: 3px 6px;
  outline: none;
  min-width: 60px;
}
table.fin-table thead tr.filter-row th input:focus {
  border-color: var(--accent);
}

table.fin-table tbody tr {
  border-bottom: 1px solid var(--border2);
  transition: background .08s;
}
table.fin-table tbody tr:nth-child(odd) { background: var(--surface); }
table.fin-table tbody tr:nth-child(even) { background: var(--row-alt); }
table.fin-table tbody tr:hover { background: var(--accent-bg) !important; }

table.fin-table tbody td {
  padding: 6px 10px;
  color: var(--text);
  border-right: 1px solid var(--border2);
  white-space: nowrap;
  max-width: 320px;
  overflow: hidden;
  text-overflow: ellipsis;
}
table.fin-table tbody td:last-child { border-right: none; }

/* Numbers get mono font */
table.fin-table tbody td.num {
  font-family: 'IBM Plex Mono', monospace;
  font-variant-numeric: tabular-nums;
  text-align: right;
  color: var(--text);
}
table.fin-table tbody td.num-pos {
  font-family: 'IBM Plex Mono', monospace;
  font-variant-numeric: tabular-nums;
  text-align: right;
  color: var(--pos);
}
table.fin-table tbody td.num-neg {
  font-family: 'IBM Plex Mono', monospace;
  font-variant-numeric: tabular-nums;
  text-align: right;
  color: var(--neg);
}

/* Empty state */
.empty-notice {
  padding: 48px 24px;
  text-align: center;
  color: var(--text3);
  font-size: 13px;
}

/* ── Scrollbar ───────────────────────────────────────────────── */
::-webkit-scrollbar { width: 7px; height: 7px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: var(--text3); }
</style>

<body>

<div class="app-header">
  <h1>USPS SLED Financial Report</h1>
  <span class="week-badge">Week of {{DATE}}</span>
  <button class="theme-btn" id="theme-toggle">☾ Dark</button>
</div>

<div class="tab-nav" id="tab-nav"></div>

<div id="panels"></div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.7.1/jquery.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/datatables/1.10.21/js/jquery.dataTables.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/datatables-buttons/2.3.6/js/dataTables.buttons.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/datatables-buttons/2.3.6/js/buttons.colVis.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/datatables-buttons/2.3.6/js/buttons.html5.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>

<script>
const DATA = {{DATA}};

// ── Theme ────────────────────────────────────────────────────────────────────
(function() {
  const saved = localStorage.getItem('fin-theme');
  if (saved) document.documentElement.setAttribute('data-theme', saved);
})();
document.getElementById('theme-toggle').addEventListener('click', function() {
  const cur = document.documentElement.getAttribute('data-theme') ||
    (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  const next = cur === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem('fin-theme', next);
  this.textContent = next === 'dark' ? '☀ Light' : '☾ Dark';
});

// ── Number detection ─────────────────────────────────────────────────────────
function isNumericCol(headers, rows, colIdx) {
  let numCount = 0, total = 0;
  for (let i = 0; i < Math.min(rows.length, 50); i++) {
    const v = rows[i][colIdx];
    if (v === '' || v === null || v === undefined) continue;
    total++;
    if (typeof v === 'number' || (typeof v === 'string' && !isNaN(parseFloat(v)) && !/^\d{4}-\d{2}/.test(v))) numCount++;
  }
  return total > 0 && numCount / total > 0.7;
}

function fmtNum(v) {
  const n = typeof v === 'number' ? v : parseFloat(v);
  if (isNaN(n)) return v;
  if (Math.abs(n) >= 1000) return n.toLocaleString('en-US', {maximumFractionDigits: 2});
  return n.toLocaleString('en-US', {maximumFractionDigits: 4});
}

// ── Build tab + panel ────────────────────────────────────────────────────────
function buildTab(key, cfg, isFirst) {
  const nav = document.getElementById('tab-nav');
  const btn = document.createElement('button');
  btn.className = 'tab-btn' + (isFirst ? ' active' : '');
  btn.dataset.panel = key;
  btn.innerHTML = cfg.label + '<span class="tab-count">' + cfg.rows.length.toLocaleString() + '</span>';
  btn.addEventListener('click', function() {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    this.classList.add('active');
    document.getElementById('panel-' + this.dataset.panel).classList.add('active');
    // Redraw table in case it was hidden
    const dt = $.fn.dataTable.isDataTable('#tbl-' + key);
    if (dt) $('#tbl-' + key).DataTable().columns.adjust();
  });
  nav.appendChild(btn);

  const panelWrap = document.getElementById('panels');
  const panel = document.createElement('div');
  panel.className = 'panel' + (isFirst ? ' active' : '');
  panel.id = 'panel-' + key;
  panel.style.overflow = 'hidden';
  panel.style.display = isFirst ? 'flex' : 'none';
  panel.style.flexDirection = 'column';

  if (!cfg.headers.length || !cfg.rows.length) {
    panel.innerHTML = '<div class="empty-notice">No data available for this tab. The source pivot table may need refreshing.</div>';
    panelWrap.appendChild(panel);
    return;
  }

  // Toolbar
  const toolbar = document.createElement('div');
  toolbar.className = 'panel-toolbar';
  toolbar.innerHTML =
    '<label>Search</label>' +
    '<input type="search" class="search-box" id="search-' + key + '" placeholder="Filter all columns…">' +
    '<div class="dt-buttons" id="btns-' + key + '"></div>' +
    '<span class="row-count" id="rc-' + key + '"></span>';
  panel.appendChild(toolbar);

  // Table wrap
  const wrap = document.createElement('div');
  wrap.className = 'table-wrap';
  wrap.style.flex = '1';
  wrap.style.overflow = 'auto';

  // Detect numeric columns
  const numericCols = cfg.headers.map((_, i) => isNumericCol(cfg.headers, cfg.rows, i));

  // Build table HTML
  let thead = '<thead><tr>' +
    cfg.headers.map(h => '<th>' + escHtml(String(h)) + '</th>').join('') +
    '</tr><tr class="filter-row">' +
    cfg.headers.map(() => '<th><input type="text" placeholder="…"></th>').join('') +
    '</tr></thead>';

  let tbody = '<tbody>';
  for (const row of cfg.rows) {
    tbody += '<tr>';
    for (let c = 0; c < cfg.headers.length; c++) {
      const raw = row[c];
      const v = (raw === null || raw === undefined) ? '' : raw;
      let cls = '';
      let display = escHtml(String(v));
      if (numericCols[c] && v !== '') {
        const n = typeof v === 'number' ? v : parseFloat(v);
        if (!isNaN(n)) {
          display = fmtNum(n);
          cls = n > 0 ? 'num-pos' : (n < 0 ? 'num-neg' : 'num');
        }
      }
      tbody += cls ? '<td class="' + cls + '" data-order="' + escHtml(String(v)) + '">' + display + '</td>'
                   : '<td>' + display + '</td>';
    }
    tbody += '</tr>';
  }
  tbody += '</tbody>';

  const tbl = document.createElement('table');
  tbl.id = 'tbl-' + key;
  tbl.className = 'fin-table';
  tbl.innerHTML = thead + tbody;
  wrap.appendChild(tbl);
  panel.appendChild(wrap);

  // Paginate container
  const pgWrap = document.createElement('div');
  pgWrap.id = 'pg-' + key;
  panel.appendChild(pgWrap);

  panelWrap.appendChild(panel);

  // Init DataTable
  $(function() {
    const colVis = cfg.headers.length > 15 ? [{extend:'colvis', text:'Columns ▾'}] : [];
    const dt = $('#tbl-' + key).DataTable({
      pageLength: 50,
      pagingType: 'simple_numbers',
      dom: '<"top"B>rt<"bottom"p>',
      buttons: [...colVis, {extend:'csv', text:'Export CSV', filename: cfg.label.replace(/\s+/g,'-')}],
      scrollX: true,
      autoWidth: false,
      orderCellsTop: true,
      fixedHeader: false,
      paginate: true,
      language: {
        paginate: { previous: '‹', next: '›' }
      },
    });

    // Move buttons into toolbar
    dt.buttons().container().appendTo('#btns-' + key);

    // Move paginate into bottom wrapper
    $('#tbl-' + key + '_wrapper .dataTables_paginate').appendTo('#pg-' + key);
    $('#pg-' + key).css({background:'var(--surface2)', borderTop:'1px solid var(--border)', padding:'6px 14px', display:'flex', justifyContent:'flex-end'});

    // Per-column filter inputs
    $('#tbl-' + key + ' thead tr.filter-row th input').each(function(i) {
      $(this).on('input', function() {
        dt.column(i).search(this.value).draw();
      });
    });

    // Global search
    $('#search-' + key).on('input', function() {
      dt.search(this.value).draw();
    });

    // Row count
    function updateCount() {
      const info = dt.page.info();
      $('#rc-' + key).text(info.recordsDisplay.toLocaleString() + ' / ' + info.recordsTotal.toLocaleString() + ' rows');
    }
    dt.on('draw', updateCount);
    updateCount();
  });
}

function escHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Render all tabs ──────────────────────────────────────────────────────────
let first = true;
for (const [key, cfg] of Object.entries(DATA)) {
  buildTab(key, cfg, first);
  first = false;
}

// Fix tab visibility on click
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', function() {
    document.querySelectorAll('.panel').forEach(p => {
      p.style.display = 'none';
      p.classList.remove('active');
    });
    const panel = document.getElementById('panel-' + this.dataset.panel);
    if (panel) {
      panel.style.display = 'flex';
      panel.classList.add('active');
    }
  });
});
</script>
</body>
</html>
"""


# ── Main extraction ────────────────────────────────────────────────────────────

def generate(excel_path, output_path=None):
    if output_path is None:
        output_path = excel_path.replace('.xlsx', '.html')

    print(f'Generating HTML dashboard from: {os.path.basename(excel_path)}')

    pythoncom.CoInitialize()
    xl = win32com.client.Dispatch('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(os.path.abspath(excel_path))

    # Unhide all sheets for reading
    for ws in wb.Worksheets:
        ws.Visible = True

    tabs_data = {}

    # PID Details — headers at row 7, data from row 8
    try:
        ws = wb.Worksheets('PID Details')
        headers = read_headers_simple(ws, 7)
        rows = read_data(ws, headers, 8)
        tabs_data['pid-details'] = {'label': 'PID Details', 'headers': headers, 'rows': rows}
        print(f'  PID Details: {len(headers)} cols, {len(rows)} rows')
    except Exception as e:
        print(f'  PID Details: skipped ({e})')
        tabs_data['pid-details'] = {'label': 'PID Details', 'headers': [], 'rows': []}

    # Fin Summary by PID — headers at row 6, data from row 7
    try:
        ws = wb.Worksheets('Fin Summary by PID')
        headers = read_headers_simple(ws, 6)
        rows = read_data(ws, headers, 7)
        tabs_data['fin-summary'] = {'label': 'Fin Summary', 'headers': headers, 'rows': rows}
        print(f'  Fin Summary: {len(headers)} cols, {len(rows)} rows')
    except Exception as e:
        print(f'  Fin Summary: skipped ({e})')
        tabs_data['fin-summary'] = {'label': 'Fin Summary', 'headers': [], 'rows': []}

    # Labor Summary — pivot, headers at rows 7-8 (groups), row 9 (labels)
    try:
        ws = wb.Worksheets('Labor Summary')
        headers = read_headers_pivot(ws, [7, 8], 9)
        rows = read_data(ws, headers, 10)
        tabs_data['labor-summary'] = {'label': 'Labor Summary', 'headers': headers, 'rows': rows}
        print(f'  Labor Summary: {len(headers)} cols, {len(rows)} rows')
    except Exception as e:
        print(f'  Labor Summary: skipped ({e})')
        tabs_data['labor-summary'] = {'label': 'Labor Summary', 'headers': [], 'rows': []}

    # Labor Monthly — pivot, headers at rows 7-10 (groups), row 11 (labels)
    try:
        ws = wb.Worksheets('Labor Monthly')
        headers = read_headers_pivot(ws, [7, 8, 9, 10], 11)
        rows = read_data(ws, headers, 12)
        tabs_data['labor-monthly'] = {'label': 'Labor Monthly', 'headers': headers, 'rows': rows}
        print(f'  Labor Monthly: {len(headers)} cols, {len(rows)} rows')
    except Exception as e:
        print(f'  Labor Monthly: skipped ({e})')
        tabs_data['labor-monthly'] = {'label': 'Labor Monthly', 'headers': [], 'rows': []}

    wb.Close(False)
    xl.Quit()

    # Extract date string from filename
    m = re.search(r'(\d{2}-\d{2}-\d{4})', os.path.basename(excel_path))
    date_str = m.group(1) if m else datetime.today().strftime('%m-%d-%Y')

    html = HTML.replace('{{DATE}}', date_str).replace('{{DATA}}', json.dumps(tabs_data, default=str))

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = os.path.getsize(output_path) // 1024
    print(f'Dashboard saved: {output_path}  ({size_kb:,} KB)')
    return output_path


if __name__ == '__main__':
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    else:
        patterns = [
            os.path.expanduser('~/OneDrive - Cisco/Desktop/Financial Report USPS SLED*.xlsx'),
            os.path.expanduser('~/Desktop/Financial Report USPS SLED*.xlsx'),
        ]
        matches = []
        for p in patterns:
            matches += [m for m in glob.glob(p) if 'BLANK' not in m]
        if not matches:
            print('ERROR: No Financial Report found on Desktop')
            sys.exit(1)
        excel_path = sorted(matches)[-1]

    out = generate(excel_path)
    # Open in default browser
    import webbrowser
    webbrowser.open('file:///' + out.replace('\\', '/'))
