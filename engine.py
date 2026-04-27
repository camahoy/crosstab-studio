"""
engine.py — Crosstab Studio core parsing and output engine
"""

print("LUCIDATA ENGINE v1.0")

import io, math, re
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font as XLFont, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from profiles import PROFILES


# ── Value coercion ────────────────────────────────────────────

def _coerce(v):
    """Convert any cell value to float if possible, None if blank, else as-is."""
    if v is None: return None
    if isinstance(v, float): return None if math.isnan(v) else v
    if isinstance(v, int):   return float(v)
    if isinstance(v, str):
        s = v.strip()
        if s in ('', '-', '\xa0', '\u00a0'): return None
        try:    return float(s)
        except: return v  # sig letters stay as string
    return v


# ── Fast sheet scanner — reads ONLY first few rows ────────────

def fast_scan(file_bytes, profile_name):
    """
    Rapidly scan all sheets using only the question row.
    Returns ordered list of question groups without parsing any data.
    Each group: {prefix, wording, sheets: [sheet_idx...]}
    """
    profile = PROFILES[profile_name]
    q_row   = profile["question_row"]

    xl      = pd.ExcelFile(io.BytesIO(file_bytes))
    groups  = {}
    order   = []

    for i, sname in enumerate(xl.sheet_names):
        if i == 0 and profile.get("skip_sheet_0"):
            # Peek at sheet 0 — skip only if it's a TOC/index
            try:
                df0  = xl.parse(0, header=None, nrows=5, na_values=[''])
                raw0 = df0.values.tolist()
                c0   = str(raw0[q_row][0]).strip() if len(raw0) > q_row and raw0[q_row] else ''
                if not c0 or len(c0) < 5:
                    continue
            except Exception:
                continue

        try:
            # Only read up to question row + 1 — very fast
            df  = xl.parse(i, header=None, nrows=q_row + 2, na_values=[''])
            raw = df.values.tolist()
            if len(raw) <= q_row: continue
            wording = str(raw[q_row][0]).strip() if raw[q_row] and raw[q_row][0] else ''
            if not wording or len(wording) < 5: continue
            if any(x in wording.lower() for x in ['sample', 'weight', 'project id']): continue

            # Extract prefix
            m      = re.match(r'^([A-Za-z0-9_]+)[\.\s]', wording)
            prefix = m.group(1) if m else wording[:8]

            if prefix not in groups:
                groups[prefix] = {'prefix': prefix, 'wording': wording[:80], 'sheets': []}
                order.append(prefix)
            groups[prefix]['sheets'].append(i)

        except Exception:
            continue

    return [groups[p] for p in order]


def get_columns(file_bytes, profile_name):
    """
    Get available column names (subgroups/countries) from first data sheet.
    Returns list of (col_index, name, sublabel) tuples.
    """
    profile = PROFILES[profile_name]
    xl      = pd.ExcelFile(io.BytesIO(file_bytes))

    # Find first real sheet
    start = 1 if profile.get("skip_sheet_0") else 0
    for i in range(start, min(start + 10, len(xl.sheet_names))):
        try:
            df  = xl.parse(i, header=None, nrows=profile.get("sublabel_row", profile["header_row"]) + 2, na_values=[''])
            raw = df.values.tolist()

            hrow = raw[profile["header_row"]] if len(raw) > profile["header_row"] else []
            srow = raw[profile.get("sublabel_row", profile["header_row"])] if "sublabel_row" in profile else []

            # Detect start column
            col_start = profile.get("column_start", 1)
            if col_start == "auto":
                col0 = hrow[0] if hrow else None
                col_start = 0 if (isinstance(col0, str) and 'total' in col0.lower()) else 1

            cols = []
            for j in range(col_start, len(hrow)):
                g = hrow[j]
                s = srow[j] if srow and j < len(srow) else ''
                if isinstance(g, str) and g.strip():
                    cols.append((j, g.strip(), s.strip() if isinstance(s, str) else ''))
            if cols:
                return cols
        except Exception:
            continue
    return []


# ── Full sheet parser ─────────────────────────────────────────

def parse_sheet(file_bytes, sheet_idx, profile_name, col_indices):
    """
    Parse one sheet and return answers + values for selected columns.
    """
    profile = PROFILES[profile_name]
    xl      = pd.ExcelFile(io.BytesIO(file_bytes))

    df  = xl.parse(sheet_idx, header=None, na_values=[''])
    raw = df.values.tolist()

    # Question wording
    q_row   = profile["question_row"]
    wording = str(raw[q_row][0]).strip() if len(raw) > q_row and raw[q_row] and raw[q_row][0] else ''

    # Base values
    b_row     = profile.get("base_row", 8)
    base_data = raw[b_row] if len(raw) > b_row else []
    base_vals = [_coerce(base_data[j] if j < len(base_data) else None) for j in col_indices]

    # Data start
    if "data_start" in profile:
        data_start = profile["data_start"]
    else:
        # Find base row dynamically then offset
        offset = profile.get("data_start_base_offset", 2)
        base_row_idx = None
        for ri, row in enumerate(raw):
            if row and isinstance(row[0], str):
                cl = row[0].strip().lower()
                if cl.startswith('base') or cl.startswith('unweighted'):
                    base_row_idx = ri
                    break
        data_start = (base_row_idx + offset) if base_row_idx is not None else 9

    step      = profile.get("data_step", 3)
    val_off   = profile.get("value_row_offset", 1)
    stop_on   = set(profile.get("stop_on", ["sigma"]))

    answers, values, sig_data = [], [], []
    i = data_start
    while i < len(raw):
        lbl = raw[i][0] if raw[i] else None
        if isinstance(lbl, str) and lbl.strip():
            cl = lbl.strip().lower()
            if cl in stop_on: break
            if any(cl.startswith(s) for s in ('base:', 'unweighted base', 'base: at least')):
                i += 1; continue

            answers.append(lbl.strip())
            val_row = raw[i + val_off] if i + val_off < len(raw) else []
            sig_row = raw[i + 2]       if i + 2       < len(raw) else []

            row_vals = [_coerce(val_row[j] if j < len(val_row) else None) for j in col_indices]
            row_sigs = [sig_row[j] if j < len(sig_row) else None for j in col_indices]
            values.append(row_vals)
            sig_data.append(row_sigs)
            i += step
            continue
        i += 1

    return {
        'wording':    wording,
        'base_vals':  base_vals,
        'answers':    answers,
        'values':     values,
        'sig_data':   sig_data,
    }


# ── Excel output ──────────────────────────────────────────────

HDR_FILL  = PatternFill("solid", fgColor="0F2D4A")
HDR_FONT  = XLFont(bold=True, color="FFFFFF", name="Helvetica Neue", size=10)
BODY_FONT = XLFont(name="Helvetica Neue", size=10)
CTR       = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN      = Side(style="thin", color="D0D7DE")
BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
ALT_FILL  = PatternFill("solid", fgColor="F6F8FA")


def _write_table(ws, row, wording, col_names, base_vals, answers, values):
    """Write one formatted table to worksheet. Returns next available row."""
    n_cols = len(col_names)

    # Question wording title
    c = ws.cell(row=row, column=1, value=wording)
    c.font      = XLFont(bold=True, name="Helvetica Neue", size=11, color="0F2D4A")
    c.alignment = LEFT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols + 1)
    row += 1

    # Header row
    ws.cell(row=row, column=1).fill   = HDR_FILL
    ws.cell(row=row, column=1).border = BORDER
    for ci, name in enumerate(col_names):
        bv   = base_vals[ci] if ci < len(base_vals) else None
        bstr = f"\n(N={int(bv):,})" if bv and isinstance(bv, float) else ''
        cell           = ws.cell(row=row, column=ci + 2, value=name + bstr)
        cell.font      = HDR_FONT
        cell.fill      = HDR_FILL
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
            cell.font=BODY_FONT; cell.alignment=CTR; cell.border=BORDER; cell.fill=fill
            if v is None:
                cell.value = '—'
            elif isinstance(v, float):
                pct = math.floor(v*100) if v*100 - math.floor(v*100) < 0.5 else math.ceil(v*100)
                cell.value = f"{pct}%"
            else:
                cell.value = str(v)
        row += 1

    # Column widths
    ws.column_dimensions['A'].width = max(ws.column_dimensions['A'].width, 40)
    for ci in range(n_cols):
        cl = get_column_letter(ci + 2)
        ws.column_dimensions[cl].width = max(ws.column_dimensions[cl].width, 13)

    return row + 2  # gap


def _build_toc(wb, toc_entries):
    """Build Contents sheet as first sheet."""
    toc = wb.create_sheet('Contents', 0)
    toc.sheet_view.showGridLines = False
    toc.column_dimensions['A'].width = 6
    toc.column_dimensions['B'].width = 70
    toc.column_dimensions['C'].width = 20

    hf   = PatternFill('solid', fgColor='0F2D4A')
    hfnt = XLFont(bold=True, color='FFFFFF', name='Helvetica Neue', size=11)
    thin = Side(style='thin', color='D0D7DE')
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, val in [(1,'#'), (2,'Question'), (3,'Sheet')]:
        c = toc.cell(row=1, column=col, value=val)
        c.font=hfnt; c.fill=hf; c.border=brd
        c.alignment=Alignment(horizontal='left', vertical='center')
    toc.row_dimensions[1].height = 24

    lf   = XLFont(color='1A6EBD', underline='single', name='Helvetica Neue', size=10)
    pf   = XLFont(name='Helvetica Neue', size=10)
    altf = PatternFill('solid', fgColor='F0F5FA')
    ethin= Side(style='thin', color='E8ECF0')
    ebrd = Border(left=ethin, right=ethin, top=ethin, bottom=ethin)

    seen = set()
    r = 2
    for i, (wording, sheet_name) in enumerate(toc_entries):
        if sheet_name in seen: continue
        seen.add(sheet_name)
        fill = altf if i % 2 == 0 else PatternFill()
        safe = sheet_name.replace("'", "''")

        nc = toc.cell(row=r, column=1, value=i+1)
        nc.font=pf; nc.border=ebrd; nc.fill=fill
        nc.alignment=Alignment(horizontal='center', vertical='center')

        qc = toc.cell(row=r, column=2, value=wording[:100])
        qc.hyperlink=f"#{safe}!A1"; qc.font=lf; qc.border=ebrd; qc.fill=fill
        qc.alignment=Alignment(horizontal='left', vertical='center', wrap_text=True)

        sc = toc.cell(row=r, column=3, value=sheet_name)
        sc.font=pf; sc.border=ebrd; sc.fill=fill
        sc.alignment=Alignment(horizontal='left', vertical='center')

        toc.row_dimensions[r].height = 18
        r += 1

    # Back links on every data sheet
    bf = XLFont(color='1A6EBD', underline='single', name='Helvetica Neue', size=9)
    for ws in wb.worksheets:
        if ws.title == 'Contents': continue
        ws.insert_rows(1)
        nav = ws.cell(row=1, column=1, value='← Contents')
        nav.hyperlink='#Contents!A1'; nav.font=bf
        nav.alignment=Alignment(horizontal='left', vertical='center')


def generate_excel(selections, file_bytes, profile_name, col_indices, col_names):
    """
    Generate Excel output for selected questions.
    selections: list of {prefix, wording, sheets: [idx...]}
    Returns bytes.
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
        ws  = wb[sheet_nm]
        row = ws.max_row + 2 if ws.max_row > 1 else 1

        # If question header not yet written
        if ws.max_row <= 1:
            toc_entries.append((wording[:80], sheet_nm))

        # Parse and write each sheet for this question
        for si in sel['sheets']:
            parsed = parse_sheet(file_bytes, si, profile_name, col_indices)
            if not parsed['answers']:
                continue
            row = _write_table(ws, row, parsed['wording'],
                               col_names, parsed['base_vals'],
                               parsed['answers'], parsed['values'])

    if toc_entries:
        _build_toc(wb, toc_entries)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
