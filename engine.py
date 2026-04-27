"""
engine.py — Crosstab Studio v1.1
Reliable reading of Corporate Reputation, Global Brand Identity, and KP formats.
Wave-by-wave comparison support.
"""

print("CROSSTAB STUDIO ENGINE v1.1")

import io, math, re
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font as XLFont, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from profiles import PROFILES


# ── Value coercion ────────────────────────────────────────────

def _coerce(v):
    """Convert any cell value to float if possible, None if blank, string otherwise."""
    if v is None: return None
    if isinstance(v, float): return None if math.isnan(v) else v
    if isinstance(v, int):   return float(v)
    if isinstance(v, str):
        s = v.strip()
        if s in ('', '-', '\xa0', '\u00a0'): return None
        try:    return float(s)
        except: return v
    return v


# ── Format validator ──────────────────────────────────────────

def validate_format(file_bytes):
    """
    Check uploaded file against all known profiles.
    Returns (matched_profile_name, confidence) or (None, 0).
    """
    try:
        xl  = pd.ExcelFile(io.BytesIO(file_bytes))
        si  = 1 if len(xl.sheet_names) > 1 else 0
        df  = xl.parse(si, header=None, nrows=12, na_values=[''])
        raw = df.values.tolist()
    except Exception:
        return None, 0

    def cell(r, c):
        try:
            v = raw[r][c]
            return str(v).strip() if v and str(v) not in ('nan','None') else ''
        except: return ''

    def row_has_total(r):
        """Check if any cell in a row contains total."""
        try:
            return any('total' in str(v).lower()
                       for v in (raw[r] or [])
                       if v and str(v) not in ('nan','None'))
        except: return False

    # Corporate Reputation (fmt6):
    # Row 2 = descriptor ('Total sample'), row 3 = question, row 4 col 1 = Total
    if len(raw) > 5:
        r2   = cell(2, 0)
        r3   = cell(3, 0)
        r4c1 = cell(4, 1)
        r5c1 = cell(5, 1)
        if (len(r3) > 10
                and ('total sample' in r2.lower() or 'weight' in r2.lower() or len(r2) < 50)
                and ('total' in r4c1.lower() or 'total' in r5c1.lower())):
            return "Corporate Reputation", 95

    # Global Brand Identity (fmt2):
    # Row 2 = question, row 3 col 0 or col 1 = Total, row 4 = sub-labels
    if len(raw) > 5:
        r2   = cell(2, 0)
        r3c0 = cell(3, 0)
        r3c1 = cell(3, 1)
        r4c1 = cell(4, 1)
        if (len(r2) > 10
                and ('total' in r3c0.lower() or 'total' in r3c1.lower())
                and 'total' in r4c1.lower()):
            return "Global Brand Identity", 90

    # KP: row 2 = question, Total somewhere in row 1
    if len(raw) > 3:
        r2 = cell(2, 0)
        if len(r2) > 10 and row_has_total(1):
            # Make sure it's not Corporate Rep or GBI
            r4c1 = cell(4, 1)
            if 'total' not in r4c1.lower():
                return "KP", 85

    return None, 0


# ── Fast scanner ──────────────────────────────────────────────

def fast_scan(file_bytes, profile_name):
    """
    Rapidly read question names from each sheet.
    Returns ordered list of question groups.
    """
    profile = PROFILES[profile_name]
    q_row   = profile["question_row"]
    xl      = pd.ExcelFile(io.BytesIO(file_bytes))
    groups  = {}
    order   = []

    for i, sname in enumerate(xl.sheet_names):
        # Skip sheet 0 if it looks like a TOC
        if i == 0 and profile.get("skip_sheet_0"):
            try:
                df0  = xl.parse(0, header=None, nrows=q_row + 2, na_values=[''])
                raw0 = df0.values.tolist()
                c0   = str(raw0[q_row][0]).strip() if len(raw0) > q_row and raw0[q_row] else ''
                if not c0 or len(c0) < 5: continue
            except Exception:
                continue

        try:
            df  = xl.parse(i, header=None, nrows=q_row + 8, na_values=[''])
            raw = df.values.tolist()
            if len(raw) <= q_row: continue

            # Get question wording — KP may span multiple rows
            wording = str(raw[q_row][0]).strip() if raw[q_row] and raw[q_row][0] else ''
            if not wording or len(wording) < 5: continue
            if any(x in wording.lower() for x in ['sample', 'weight', 'project id', 'table of']): continue

            # For KP: check if wording continues on next rows
            if profile_name == "KP":
                for extra_row in range(q_row + 1, min(q_row + 7, len(raw))):
                    next_cell = str(raw[extra_row][0]).strip() if raw[extra_row] and raw[extra_row][0] else ''
                    if not next_cell or 'base' in next_cell.lower() or '=' in next_cell:
                        break
                    wording += ' ' + next_cell

            # Detect sheet type from wording
            sheet_type = _detect_sheet_type(wording)

            # Extract prefix
            m      = re.match(r'^([A-Za-z0-9_]+)[\.\s]', wording)
            prefix = m.group(1) if m else wording[:8]

            if prefix not in groups:
                groups[prefix] = {
                    'prefix':     prefix,
                    'wording':    wording[:100],
                    'sheets':     [],
                    'types':      set(),
                }
                order.append(prefix)

            groups[prefix]['sheets'].append(i)
            groups[prefix]['types'].add(sheet_type)

        except Exception:
            continue

    # Convert types set to sorted list
    for g in groups.values():
        g['types'] = sorted(g['types'])

    return [groups[p] for p in order]


def _detect_sheet_type(wording):
    """Classify a sheet as standard, t2b, b2b, summary_grid, or mean."""
    wl = wording.lower().replace("'[","[")
    if 't2b' in wl or "top 2 box" in wl:         return 't2b'
    if 'b2b' in wl or "bottom 2 box" in wl:       return 'b2b'
    if 'summary grid' in wl or 'grid' in wl:       return 'summary_grid'
    if 'mean' in wl or 'average' in wl:            return 'mean'
    return 'standard'


# ── Column reader ─────────────────────────────────────────────

def get_columns(file_bytes, profile_name):
    """
    Get available column names from first real data sheet.
    Returns list of (col_index, name, sublabel).
    """
    profile = PROFILES[profile_name]
    xl      = pd.ExcelFile(io.BytesIO(file_bytes))
    start   = 1 if profile.get("skip_sheet_0") else 0
    h_row   = profile["header_row"]
    s_row   = profile.get("sublabel_row", h_row)
    nrows   = max(h_row, s_row) + 2

    for i in range(start, min(start + 10, len(xl.sheet_names))):
        try:
            df   = xl.parse(i, header=None, nrows=nrows, na_values=[''])
            raw  = df.values.tolist()
            hrow = raw[h_row] if len(raw) > h_row else []
            srow = raw[s_row] if len(raw) > s_row and s_row != h_row else []

            col_start = profile.get("column_start", 1)

            # Auto-detect: Total at col 0 or col 1
            if col_start == "auto":
                col0      = hrow[0] if hrow else None
                col_start = 0 if (isinstance(col0, str) and 'total' in col0.lower()) else 1

            # KP: find where 'total' is in the header row
            elif col_start == "find_total":
                col_start = 0
                for ci, v in enumerate(hrow):
                    if isinstance(v, str) and 'total' in v.lower():
                        col_start = ci
                        break

            cols = []
            for j in range(col_start, len(hrow)):
                g = hrow[j]
                s = srow[j] if srow and j < len(srow) else ''
                if isinstance(g, str) and g.strip() and g.strip() not in ('\xa0',):
                    cols.append((j, g.strip(), s.strip() if isinstance(s, str) else ''))
            if cols:
                return cols
        except Exception:
            continue
    return []


# ── Sheet parser ──────────────────────────────────────────────

def parse_sheet(file_bytes, sheet_idx, profile_name, col_indices):
    """
    Parse one sheet. Returns wording, base values, answers, values.
    Handles all three formats reliably.
    """
    profile = PROFILES[profile_name]
    xl      = pd.ExcelFile(io.BytesIO(file_bytes))
    df      = xl.parse(sheet_idx, header=None, na_values=[''])
    raw     = df.values.tolist()

    # ── Question wording ──
    q_row   = profile["question_row"]
    wording = str(raw[q_row][0]).strip() if len(raw) > q_row and raw[q_row] and raw[q_row][0] else ''

    # KP: multi-line question wording
    if profile_name == "KP":
        for extra_row in range(q_row + 1, min(q_row + 7, len(raw))):
            next_cell = str(raw[extra_row][0]).strip() if raw[extra_row] and raw[extra_row][0] else ''
            if not next_cell or 'base' in next_cell.lower() or '=' in next_cell:
                break
            wording += ' ' + next_cell

    # ── Base values ──
    b_row     = profile.get("base_row", 7)
    base_data = raw[b_row] if len(raw) > b_row else []
    base_vals = [_coerce(base_data[j] if j < len(base_data) else None) for j in col_indices]

    # ── Data start ──
    if profile.get("data_start") is not None:
        data_start = profile["data_start"]
    else:
        # Dynamic: find base row then offset
        offset       = profile.get("data_start_base_offset", 2)
        base_row_idx = None
        for ri, row in enumerate(raw):
            if row and isinstance(row[0], str):
                cl = row[0].strip().lower()
                if cl.startswith('base') or cl.startswith('unweighted') or 'base =' in cl:
                    base_row_idx = ri
                    break
        data_start = (base_row_idx + offset) if base_row_idx is not None else b_row + 2

    step    = profile.get("data_step", 3)
    val_off = profile.get("value_row_offset", 1)
    stop_on = set(s.lower() for s in profile.get("stop_on", ["sigma"]))

    answers, values, sig_data = [], [], []
    i = data_start
    while i < len(raw):
        lbl = raw[i][0] if raw[i] else None
        if isinstance(lbl, str) and lbl.strip():
            cl = lbl.strip().lower()
            if cl in stop_on: break
            if any(cl.startswith(s) for s in ('base:', 'unweighted base', 'base: at least',
                                               'base =', 'weighted base')):
                i += 1; continue

            answers.append(lbl.strip())

            val_row = raw[i + val_off] if i + val_off < len(raw) else []
            sig_row = raw[i + 2]       if i + 2       < len(raw) else []

            row_vals = [_coerce(val_row[j] if j < len(val_row) else None) for j in col_indices]
            row_sigs = [sig_row[j]          if j < len(sig_row)  else None for j in col_indices]

            values.append(row_vals)
            sig_data.append(row_sigs)
            i += step
            continue
        i += 1

    return {
        'wording':   wording,
        'base_vals': base_vals,
        'answers':   answers,
        'values':    values,
        'sig_data':  sig_data,
    }


# ── Excel styles ──────────────────────────────────────────────

HDR_FILL    = PatternFill("solid", fgColor="0F2D4A")
WAVE_FILLS  = [
    PatternFill("solid", fgColor="E8F0FE"),
    PatternFill("solid", fgColor="D2E3FC"),
    PatternFill("solid", fgColor="BCCFEF"),
    PatternFill("solid", fgColor="A8C7FA"),
]
HDR_FONT    = XLFont(bold=True, color="FFFFFF", name="Arial", size=10)
WAVE_FONTS  = [
    XLFont(bold=True, color="0F2D4A", name="Arial", size=10),
    XLFont(bold=True, color="0F2D4A", name="Arial", size=10),
    XLFont(bold=True, color="0F2D4A", name="Arial", size=10),
    XLFont(bold=True, color="0F2D4A", name="Arial", size=10),
]
BODY_FONT   = XLFont(name="Arial", size=10)
CTR         = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT        = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN        = Side(style="thin",  color="D0D7DE")
BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ALT_FILL    = PatternFill("solid", fgColor="F6F8FA")


def _fmt_pct(v):
    if v is None: return '—'
    if isinstance(v, float):
        pct = math.floor(v*100) if v*100 - math.floor(v*100) < 0.5 else math.ceil(v*100)
        return f"{pct}%"
    return str(v)


def _write_table(ws, row, wording, col_headers, base_vals,
                 answers, values, wave_idx=None):
    """
    Write one formatted table.
    wave_idx: if set, uses wave-specific header fill instead of dark fill.
    Returns next available row.
    """
    n_cols = len(col_headers)

    # Title row
    c = ws.cell(row=row, column=1, value=wording)
    c.font      = XLFont(bold=True, name="Arial", size=11, color="0F2D4A")
    c.alignment = LEFT
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=n_cols + 1)
    row += 1

    # Header
    use_fill = WAVE_FILLS[wave_idx % len(WAVE_FILLS)] if wave_idx is not None else HDR_FILL
    use_font = WAVE_FONTS[wave_idx % len(WAVE_FONTS)] if wave_idx is not None else HDR_FONT

    ws.cell(row=row, column=1).fill   = use_fill
    ws.cell(row=row, column=1).border = BORDER
    for ci, name in enumerate(col_headers):
        bv   = base_vals[ci] if ci < len(base_vals) else None
        bstr = f"\n(N={int(bv):,})" if isinstance(bv, float) and bv else ''
        cell           = ws.cell(row=row, column=ci + 2, value=name + bstr)
        cell.font      = use_font
        cell.fill      = use_fill
        cell.alignment = CTR
        cell.border    = BORDER
    row += 1

    # Data rows
    for ri, answer in enumerate(answers):
        fill = ALT_FILL if ri % 2 == 0 else PatternFill()
        lc = ws.cell(row=row, column=1, value=str(answer))
        lc.font=BODY_FONT; lc.alignment=LEFT; lc.border=BORDER; lc.fill=fill

        row_vals = values[ri] if ri < len(values) else []
        for ci in range(n_cols):
            v    = row_vals[ci] if ci < len(row_vals) else None
            cell = ws.cell(row=row, column=ci + 2)
            cell.font=BODY_FONT; cell.alignment=CTR
            cell.border=BORDER;  cell.fill=fill
            cell.value = _fmt_pct(v)
        row += 1

    # Column widths
    ws.column_dimensions['A'].width = max(ws.column_dimensions['A'].width, 42)
    for ci in range(n_cols):
        cl = get_column_letter(ci + 2)
        ws.column_dimensions[cl].width = max(ws.column_dimensions[cl].width, 14)

    return row + 2


def _build_toc(wb, toc_entries):
    """Build Contents sheet as first tab with hyperlinks."""
    toc = wb.create_sheet('Contents', 0)
    toc.sheet_view.showGridLines = False
    toc.column_dimensions['A'].width = 6
    toc.column_dimensions['B'].width = 72
    toc.column_dimensions['C'].width = 22

    hf   = PatternFill('solid', fgColor='0F2D4A')
    hfnt = XLFont(bold=True, color='FFFFFF', name='Arial', size=11)
    thin = Side(style='thin', color='D0D7DE')
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, val in [(1,'#'), (2,'Question'), (3,'Sheet')]:
        c = toc.cell(row=1, column=col, value=val)
        c.font=hfnt; c.fill=hf; c.border=brd
        c.alignment=Alignment(horizontal='left', vertical='center')
    toc.row_dimensions[1].height = 24

    lf   = XLFont(color='1A6EBD', underline='single', name='Arial', size=10)
    pf   = XLFont(name='Arial', size=10)
    altf = PatternFill('solid', fgColor='F0F5FA')
    et   = Side(style='thin', color='E8ECF0')
    ebrd = Border(left=et, right=et, top=et, bottom=et)

    seen = set()
    r = 2
    for i, (wording, sheet_name) in enumerate(toc_entries):
        if sheet_name in seen: continue
        seen.add(sheet_name)
        fill = altf if i % 2 == 0 else PatternFill()
        safe = sheet_name.replace("'","''")

        nc = toc.cell(row=r, column=1, value=i+1)
        nc.font=pf; nc.border=ebrd; nc.fill=fill
        nc.alignment=Alignment(horizontal='center', vertical='center')

        qc = toc.cell(row=r, column=2, value=wording[:100])
        qc.hyperlink=f"#{safe}!A1"; qc.font=lf
        qc.border=ebrd; qc.fill=fill
        qc.alignment=Alignment(horizontal='left', vertical='center', wrap_text=True)

        sc = toc.cell(row=r, column=3, value=sheet_name)
        sc.font=pf; sc.border=ebrd; sc.fill=fill
        sc.alignment=Alignment(horizontal='left', vertical='center')

        toc.row_dimensions[r].height = 18
        r += 1

    # Back-links on every data sheet
    bf = XLFont(color='1A6EBD', underline='single', name='Arial', size=9)
    for ws in wb.worksheets:
        if ws.title == 'Contents': continue
        ws.insert_rows(1)
        nav = ws.cell(row=1, column=1, value='← Contents')
        nav.hyperlink='#Contents!A1'; nav.font=bf
        nav.alignment=Alignment(horizontal='left', vertical='center')


# ── Main export ───────────────────────────────────────────────

def generate_excel(selections, files, profile_name, col_indices, col_names):
    """
    Generate formatted Excel output.

    selections: [{prefix, wording, sheets: [idx...]}]
    files:      [{bytes, label}]
    Single file → one table per question.
    Multiple files → one table per file per question (wave comparison),
                     each wave in a different header colour.
    """
    wb          = openpyxl.Workbook()
    wb.remove(wb.active)
    toc_entries = []

    for sel in selections:
        prefix   = sel['prefix']
        wording  = sel['wording']
        sheet_nm = re.sub(r'[\\/*?\[\]:]', '', prefix)[:31]

        if sheet_nm not in wb.sheetnames:
            wb.create_sheet(title=sheet_nm)
            toc_entries.append((wording[:80], sheet_nm))

        ws  = wb[sheet_nm]
        row = 1 if ws.max_row <= 1 else ws.max_row + 2

        if len(files) == 1:
            # ── Single file ──
            p = _find_and_parse(files[0]['bytes'], prefix, profile_name, col_indices)
            if p and p['answers']:
                row = _write_table(ws, row, p['wording'], col_names,
                                   p['base_vals'], p['answers'], p['values'])
        else:
            # ── Multi-file wave comparison ──
            # Write question title once
            title_cell = ws.cell(row=row, column=1, value=wording[:100])
            title_cell.font      = XLFont(bold=True, name="Arial", size=12, color="0F2D4A")
            title_cell.alignment = LEFT
            ws.merge_cells(start_row=row, start_column=1,
                           end_row=row, end_column=len(col_names) + 1)
            row += 2

            for fi, finfo in enumerate(files):
                p = _find_and_parse(finfo['bytes'], prefix, profile_name, col_indices)
                if p is None or not p['answers']:
                    continue
                # Use wave label as table title, wave-coloured header
                row = _write_table(ws, row,
                                   finfo['label'],
                                   col_names,
                                   p['base_vals'],
                                   p['answers'],
                                   p['values'],
                                   wave_idx=fi)

    if toc_entries:
        _build_toc(wb, toc_entries)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _find_and_parse(file_bytes, prefix, profile_name, col_indices):
    """Find the first sheet matching prefix in a file and parse it."""
    groups = fast_scan(file_bytes, profile_name)
    match  = next((g for g in groups if g['prefix'] == prefix), None)
    if match is None: return None
    return parse_sheet(file_bytes, match['sheets'][0], profile_name, col_indices)
