"""
manifest_builder.py — Manifest Builder for Crosstab Studio
Takes a slide manifest CSV + W4 banner(s) + optional W3 banner(s),
extracts specified metrics per slide, writes structured Excel output.
"""

import io, re, math
import csv
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font as XLFont, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from engine import (
    read_toc, toc_to_groups, fast_scan, get_columns,
    parse_sheet, _find_and_parse_all,
)
from profiles import get_profile, get_profile_names


# ── Styles ────────────────────────────────────────────────────────────────────

def _f(bold=False, size=10, color="000000", name="Arial"):
    return XLFont(bold=bold, size=size, color=color, name=name)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _border(color="D0D7DE"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

SECTION_FILL  = _fill("0F2D4A")
SECTION_FONT  = _f(bold=True, size=12, color="FFFFFF")
SLIDE_FILL    = _fill("1A6EBD")
SLIDE_FONT    = _f(bold=True, size=10, color="FFFFFF")
HEADER_FILL   = _fill("E8F0FE")
HEADER_FONT   = _f(bold=True, size=9, color="0F2D4A")
W3_FILL       = _fill("F0F9FF")
W3_FONT       = _f(size=9, color="374151")
DELTA_FILL    = _fill("FFFBEB")
DELTA_FONT    = _f(bold=True, size=9, color="7C2D12")
BODY_FONT     = _f(size=9)
LABEL_FONT    = _f(bold=True, size=9)
MISS_FILL     = _fill("FEF2F2")
NEW_FILL      = _fill("F0FDF4")
BORDER        = _border()
THIN_BORDER   = _border("E2E8F0")

CTR = _align("center")
LEFT = _align("left")


# ── CSV parser ────────────────────────────────────────────────────────────────

def parse_manifest_csv(csv_bytes):
    """Parse manifest CSV bytes → list of row dicts."""
    text = csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    rows = []
    for r in reader:
        slide = r.get("Slide", "").strip()
        if not slide:
            continue
        rows.append({
            "slide":        slide,
            "section":      r.get("Section", "").strip(),
            "title":        r.get("Slide_Title", "").strip(),
            "prefix":       r.get("Question_Code", "").strip(),
            "sheet_type":   r.get("Sheet_Type", "").strip(),
            "metric":       r.get("Metric", "").strip(),
            "brands":       [b.strip() for b in r.get("Brands_Entities", "").split("|") if b.strip() and b.strip().upper() != "N/A"],
            "audiences":    [a.strip() for a in r.get("Audiences", "").split("|") if a.strip()],
            "wave_compare": r.get("Wave_Compare", "").strip(),
            "notes":        r.get("Notes", "").strip(),
        })
    return rows


# ── Audience / column matching ────────────────────────────────────────────────

_AUDIENCE_PATTERNS = {
    "Gen Pop":    ["gen pop", "general population", "total", "adult"],
    "Tech Elite": ["tech elite", "technology elite", "tech"],
    "AI Fan":     ["ai fan", "ai fans", "artificial intelligence fan"],
    "Gen Z":      ["gen z", "generation z", "18-29", "18–29", "genz"],
}

def _match_col_for_audience(cols, audience_name):
    """
    Find best column index (from get_columns result) for an audience label.
    cols = [(col_idx, name, sublabel), ...]
    Returns col_idx or None.
    """
    audience_lc = audience_name.lower()
    patterns = _AUDIENCE_PATTERNS.get(audience_name, [audience_lc])

    # First: exact match on col name
    for idx, name, sub in cols:
        if name.strip().lower() == audience_lc:
            return idx

    # Second: pattern substring match
    for pat in patterns:
        for idx, name, sub in cols:
            combined = (name + " " + sub).lower()
            if pat in combined:
                return idx

    # Third: fuzzy — any word from audience in column name
    words = audience_lc.split()
    for idx, name, sub in cols:
        col_lc = (name + " " + sub).lower()
        if any(w in col_lc for w in words if len(w) > 2):
            return idx

    return None


def get_audience_col_map(file_bytes, profile_name, audience_list):
    """
    Return {audience_name: col_idx} for each audience in audience_list.
    Missing audiences get None.
    """
    cols = get_columns(file_bytes, profile_name)
    return {aud: _match_col_for_audience(cols, aud) for aud in audience_list}


# ── TOC inventory ─────────────────────────────────────────────────────────────

def get_all_prefixes(file_bytes, profile_name):
    """Return set of all question prefixes found in the banner."""
    toc = read_toc(file_bytes, profile_name)
    if toc:
        return {e["prefix"] for e in toc}
    groups = fast_scan(file_bytes, profile_name)
    return {g["prefix"] for g in groups}


# ── Metric extraction ─────────────────────────────────────────────────────────

def _fmt(v):
    """Format a float value as percentage string."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, float):
        pct = round(v * 100)
        return pct
    if isinstance(v, str):
        return v
    return v


def _extract_metric(parsed, metric_type, brands_filter=None):
    """
    From a parsed sheet dict, extract the metric.
    Returns list of (row_label, value_list) where value_list aligns to col_indices.

    metric_type: "T2B", "T4B", "All options", "Full scale", "MEAN RANK",
                 "PCT correct", "Weighted rank score", "RANK", "T2B optimistic + T2B pessimistic"
    brands_filter: list of brand strings to keep (case-insensitive substring match), or None=all
    """
    answers = parsed.get("answers", [])
    values  = parsed.get("values", [])
    metric_lc = metric_type.lower()

    def _brand_match(label, brands):
        if not brands:
            return True
        ll = label.lower()
        return any(b.lower() in ll or ll in b.lower() for b in brands)

    # Full scale / All options → return all rows (filtered by brand if needed)
    if any(k in metric_lc for k in ["all options", "full scale"]):
        rows = []
        for i, ans in enumerate(answers):
            if _brand_match(ans, brands_filter):
                rows.append((ans, values[i] if i < len(values) else []))
        return rows

    # T2B → find row labeled "Top 2 Box" / "T2B"; fallback: sum top 2 non-net rows
    if "t2b" in metric_lc and "t4b" not in metric_lc:
        # Check for explicit T2B row
        for i, ans in enumerate(answers):
            al = ans.lower()
            if "top 2 box" in al or al.startswith("t2b") or al == "top 2":
                if brands_filter:
                    # The T2B row represents the whole question, not a brand
                    # In summary sheets, rows ARE brands → match brand
                    if not _brand_match(ans, brands_filter):
                        continue
                return [(ans, values[i] if i < len(values) else [])]

        # If no explicit T2B row, this might be a summary sheet where rows = brands
        # Return all rows that match brand filter
        rows = []
        for i, ans in enumerate(answers):
            if _brand_match(ans, brands_filter):
                rows.append((ans, values[i] if i < len(values) else []))
        if rows:
            return rows

        # Last resort: sum top 2 numeric rows
        numeric_rows = [(a, v) for a, v in zip(answers, values)
                        if v and any(isinstance(x, float) for x in v)]
        if len(numeric_rows) >= 2:
            a1, v1 = numeric_rows[0]
            a2, v2 = numeric_rows[1]
            n_c = max(len(v1), len(v2))
            combined = []
            for ci in range(n_c):
                x1 = v1[ci] if ci < len(v1) and isinstance(v1[ci], float) else 0.0
                x2 = v2[ci] if ci < len(v2) and isinstance(v2[ci], float) else 0.0
                combined.append(x1 + x2)
            return [("Top 2 Box (computed)", combined)]
        return []

    # T4B → find "Top 4 Box" row; fallback: sum top 4
    if "t4b" in metric_lc:
        for i, ans in enumerate(answers):
            al = ans.lower()
            if "top 4 box" in al or al.startswith("t4b") or al == "top 4":
                return [(ans, values[i] if i < len(values) else [])]
        numeric_rows = [(a, v) for a, v in zip(answers, values)
                        if v and any(isinstance(x, float) for x in v)]
        if len(numeric_rows) >= 4:
            combined = []
            n_c = max(len(r[1]) for r in numeric_rows[:4])
            for ci in range(n_c):
                total = sum(
                    r[1][ci] if ci < len(r[1]) and isinstance(r[1][ci], float) else 0.0
                    for r in numeric_rows[:4]
                )
                combined.append(total)
            return [("Top 4 Box (computed)", combined)]
        return []

    # Mean / rank / weighted rank score
    if any(k in metric_lc for k in ["mean", "rank", "pct correct", "weighted rank"]):
        for i, ans in enumerate(answers):
            al = ans.lower()
            if any(k in al for k in ["mean", "average", "weighted rank", "rank score"]):
                return [(ans, values[i] if i < len(values) else [])]
        # Fallback: return all rows filtered by brand
        rows = []
        for i, ans in enumerate(answers):
            if _brand_match(ans, brands_filter):
                rows.append((ans, values[i] if i < len(values) else []))
        return rows

    # Default: return all rows filtered by brand
    rows = []
    for i, ans in enumerate(answers):
        if _brand_match(ans, brands_filter):
            rows.append((ans, values[i] if i < len(values) else []))
    return rows


# ── Per-brand sheet lookup ────────────────────────────────────────────────────

def _find_sheets_for_prefix(file_bytes, profile_name, prefix):
    """Return list of (sheet_idx, wording, sheet_type) for a prefix."""
    toc = read_toc(file_bytes, profile_name)
    if toc:
        return [(e["sheet_idx"], e["wording"], e["sheet_type"])
                for e in toc if e["prefix"] == prefix]
    groups = fast_scan(file_bytes, profile_name)
    match = next((g for g in groups if g["prefix"] == prefix), None)
    if not match:
        return []
    return [(si, match["wording"], t)
            for si, t in zip(match["sheets"], match.get("types", ["standard"]*len(match["sheets"])))]


def _find_brand_in_wording(wording, brands):
    """Return which brand from the list appears in this sheet's wording, or None."""
    wl = wording.lower()
    for b in brands:
        if b.lower() in wl:
            return b
    return None


# ── Slide data extraction ─────────────────────────────────────────────────────

def extract_slide_data(manifest_row, file_bytes, profile_name, col_map):
    """
    Extract data for one manifest row from a banner file.
    col_map = {audience: col_idx}

    Returns:
        {
            "found": bool,
            "rows": [(row_label, {audience: value}), ...]
            "base":  {audience: base_n}
            "wording": str
        }
    """
    prefix     = manifest_row["prefix"]
    metric     = manifest_row["metric"]
    brands     = manifest_row["brands"]
    audiences  = manifest_row["audiences"]
    sheet_type = manifest_row["sheet_type"]

    # Build col_indices list in audience order
    col_indices = []
    aud_order   = []
    for aud in audiences:
        ci = col_map.get(aud)
        if ci is not None:
            col_indices.append(ci)
            aud_order.append(aud)

    if not col_indices:
        return {"found": False, "rows": [], "base": {}, "wording": ""}

    # Detect if this is a per-brand/per-entity sheet type
    per_entity = any(k in sheet_type.lower() for k in ["per brand", "per ceo", "per tool",
                                                         "per assistant", "per statement",
                                                         "per medium", "per outlet"])

    if per_entity and brands:
        # Parse each brand's sheet individually
        all_sheets = _find_sheets_for_prefix(file_bytes, profile_name, prefix)
        result_rows = []
        base_vals   = {}
        wording_seen = ""

        for brand in brands:
            # Find the sheet for this brand
            matched_sheet = None
            for si, word, stype in all_sheets:
                if _find_brand_in_wording(word, [brand]):
                    matched_sheet = si
                    wording_seen  = word
                    break

            if matched_sheet is None:
                result_rows.append((brand, {aud: None for aud in aud_order}))
                continue

            p = parse_sheet(file_bytes, matched_sheet, profile_name, col_indices)
            if not p:
                result_rows.append((brand, {aud: None for aud in aud_order}))
                continue

            if not wording_seen:
                wording_seen = p.get("wording", "")

            # Update base
            for ai, aud in enumerate(aud_order):
                if aud not in base_vals:
                    bv = p["base_vals"][ai] if ai < len(p["base_vals"]) else None
                    if bv is not None:
                        base_vals[aud] = int(bv) if isinstance(bv, float) else bv

            rows = _extract_metric(p, metric, brands_filter=None)
            if rows:
                # For per-brand sheets, the metric row label doesn't matter much
                # We want just one value per brand per audience
                _, vals = rows[0]
                row_dict = {}
                for ai, aud in enumerate(aud_order):
                    row_dict[aud] = _fmt(vals[ai]) if ai < len(vals) else None
                result_rows.append((brand, row_dict))
            else:
                result_rows.append((brand, {aud: None for aud in aud_order}))

        return {"found": bool(result_rows), "rows": result_rows,
                "base": base_vals, "wording": wording_seen}

    else:
        # T2B Summary or Full scale — one or a few sheets contain all the data
        parsed_list = _find_and_parse_all(file_bytes, prefix, profile_name, col_indices)
        if not parsed_list:
            return {"found": False, "rows": [], "base": {}, "wording": ""}

        # Use first parsed sheet
        p = parsed_list[0]
        wording = p.get("wording", "")

        base_vals = {}
        for ai, aud in enumerate(aud_order):
            bv = p["base_vals"][ai] if ai < len(p["base_vals"]) else None
            if bv is not None:
                base_vals[aud] = int(bv) if isinstance(bv, float) else bv

        rows = _extract_metric(p, metric, brands_filter=brands or None)
        result_rows = []
        for label, vals in rows:
            row_dict = {}
            for ai, aud in enumerate(aud_order):
                row_dict[aud] = _fmt(vals[ai]) if ai < len(vals) else None
            result_rows.append((label, row_dict))

        return {"found": True, "rows": result_rows, "base": base_vals, "wording": wording}


# ── Excel builder ─────────────────────────────────────────────────────────────

def build_manifest_excel(manifest_rows, w4_files, w3_files, profile_name,
                         missing_prefixes, new_prefixes):
    """
    Build the output Excel.
    w4_files / w3_files: list of {bytes, label, name}
    missing_prefixes: set of prefixes in manifest not found in W4
    new_prefixes: set of prefixes in W4 not in manifest
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Collect all unique audiences to build column maps
    all_audiences = []
    seen_aud = set()
    for row in manifest_rows:
        for a in row["audiences"]:
            if a not in seen_aud:
                all_audiences.append(a)
                seen_aud.add(a)

    # Build col maps for each file
    def _col_maps(files):
        maps = []
        for finfo in files:
            maps.append(get_audience_col_map(finfo["bytes"], profile_name, all_audiences))
        return maps

    w4_col_maps = _col_maps(w4_files)
    w3_col_maps = _col_maps(w3_files) if w3_files else []

    # Merge col maps across files: prefer first file that has a mapping
    def _merged_col_map(col_maps, aud_list):
        result = {}
        for aud in aud_list:
            for cm in col_maps:
                if cm.get(aud) is not None:
                    result[aud] = cm[aud]
                    break
        return result

    # Group manifest rows by section
    sections = {}
    section_order = []
    for row in manifest_rows:
        sec = row["section"] or "Other"
        if sec not in sections:
            sections[sec] = []
            section_order.append(sec)
        sections[sec].append(row)

    # Create tabs: one per section + a "Flags" tab
    for sec in section_order:
        tab_name = re.sub(r'[\\/*?\[\]:]', '', sec)[:31]
        ws = wb.create_sheet(title=tab_name)
        ws.column_dimensions["A"].width = 42
        current_row = 1

        # Section header
        for ci in range(1, 20):
            ws.cell(row=current_row, column=ci).fill = SECTION_FILL
        c = ws.cell(row=current_row, column=1, value=sec.upper())
        c.font = SECTION_FONT
        c.alignment = LEFT
        current_row += 2

        for mrow in sections[sec]:
            prefix     = mrow["prefix"]
            audiences  = mrow["audiences"]
            do_wave    = mrow["wave_compare"].lower().startswith("w") and bool(w3_files)
            is_missing = prefix in missing_prefixes
            is_new_q   = mrow["wave_compare"].lower() == "n/a - new"

            # Build effective col maps for this slide
            w4_map = _merged_col_map(w4_col_maps, audiences)
            w3_map = _merged_col_map(w3_col_maps, audiences) if w3_files else {}

            # Find best W4 file for this prefix
            w4_data = {"found": False, "rows": [], "base": {}, "wording": ""}
            for fi, finfo in enumerate(w4_files):
                cm = w4_col_maps[fi] if fi < len(w4_col_maps) else {}
                data = extract_slide_data(mrow, finfo["bytes"], profile_name, cm)
                if data["found"]:
                    w4_data = data
                    break

            # Find best W3 file for this prefix
            w3_data = {"found": False, "rows": [], "base": {}, "wording": ""}
            if do_wave and w3_files:
                for fi, finfo in enumerate(w3_files):
                    cm = w3_col_maps[fi] if fi < len(w3_col_maps) else {}
                    data = extract_slide_data(mrow, finfo["bytes"], profile_name, cm)
                    if data["found"]:
                        w3_data = data
                        break

            # ── Slide header ──────────────────────────────────────────
            slide_text = f"Slide {mrow['slide']}: {mrow['title']}"
            if is_missing:
                slide_text += "  ⚠ NOT FOUND IN BANNER"
            elif is_new_q:
                slide_text += "  ★ NEW THIS WAVE"

            fill = MISS_FILL if is_missing else (NEW_FILL if is_new_q else None)
            c = ws.cell(row=current_row, column=1, value=slide_text)
            c.font = SLIDE_FONT if not is_missing else _f(bold=True, size=10, color="991B1B")
            c.fill = _fill("FEF2F2") if is_missing else (_fill("F0FDF4") if is_new_q else SLIDE_FILL)
            c.alignment = LEFT
            # Fill rest of header row
            n_header_cols = 1 + len(audiences) * (3 if do_wave else 1)
            ws.merge_cells(start_row=current_row, start_column=1,
                           end_row=current_row, end_column=max(n_header_cols, 2))
            current_row += 1

            if mrow["notes"]:
                nc = ws.cell(row=current_row, column=1, value=mrow["notes"])
                nc.font = _f(size=8, color="6B7280")
                nc.alignment = LEFT
                current_row += 1

            if is_missing:
                current_row += 1
                continue

            # ── Column headers ────────────────────────────────────────
            ws.cell(row=current_row, column=1, value="").fill = HEADER_FILL

            col_cursor = 2
            col_map_pos = {}  # audience → (w4_col, w3_col, delta_col) in worksheet

            for aud in audiences:
                if do_wave:
                    c_w4 = ws.cell(row=current_row, column=col_cursor, value=f"{aud} W4")
                    c_w4.font = HEADER_FONT
                    c_w4.fill = HEADER_FILL
                    c_w4.alignment = CTR
                    ws.column_dimensions[get_column_letter(col_cursor)].width = 12

                    c_w3 = ws.cell(row=current_row, column=col_cursor + 1, value=f"{aud} W3")
                    c_w3.font = _f(bold=True, size=9, color="374151")
                    c_w3.fill = W3_FILL
                    c_w3.alignment = CTR
                    ws.column_dimensions[get_column_letter(col_cursor + 1)].width = 12

                    c_d = ws.cell(row=current_row, column=col_cursor + 2, value="Δ")
                    c_d.font = DELTA_FONT
                    c_d.fill = DELTA_FILL
                    c_d.alignment = CTR
                    ws.column_dimensions[get_column_letter(col_cursor + 2)].width = 8

                    col_map_pos[aud] = (col_cursor, col_cursor + 1, col_cursor + 2)
                    col_cursor += 3
                else:
                    c_w4 = ws.cell(row=current_row, column=col_cursor, value=aud)
                    c_w4.font = HEADER_FONT
                    c_w4.fill = HEADER_FILL
                    c_w4.alignment = CTR
                    ws.column_dimensions[get_column_letter(col_cursor)].width = 14

                    col_map_pos[aud] = (col_cursor, None, None)
                    col_cursor += 1

            current_row += 1

            # ── Base row ──────────────────────────────────────────────
            if w4_data["base"]:
                bc = ws.cell(row=current_row, column=1, value="Base (n)")
                bc.font = _f(bold=True, size=8, color="6B7280")
                bc.alignment = LEFT
                for aud in audiences:
                    pos = col_map_pos.get(aud)
                    if not pos:
                        continue
                    w4_col, w3_col, d_col = pos
                    bv4 = w4_data["base"].get(aud)
                    if bv4 is not None:
                        cell = ws.cell(row=current_row, column=w4_col, value=bv4)
                        cell.font = _f(size=8, color="6B7280")
                        cell.alignment = CTR
                    if do_wave and w3_col:
                        bv3 = w3_data["base"].get(aud)
                        if bv3 is not None:
                            cell = ws.cell(row=current_row, column=w3_col, value=bv3)
                            cell.font = _f(size=8, color="6B7280")
                            cell.alignment = CTR
                current_row += 1

            # ── Data rows ─────────────────────────────────────────────
            # Build lookup for W3 rows
            w3_lookup = {label: vals for label, vals in w3_data.get("rows", [])}

            alt = False
            for label, w4_vals in w4_data.get("rows", []):
                row_fill = _fill("F6F8FA") if alt else _fill("FFFFFF")
                alt = not alt

                lc = ws.cell(row=current_row, column=1, value=label)
                lc.font = LABEL_FONT
                lc.fill = row_fill
                lc.alignment = LEFT
                lc.border = THIN_BORDER

                for aud in audiences:
                    pos = col_map_pos.get(aud)
                    if not pos:
                        continue
                    w4_col, w3_col, d_col = pos

                    v4 = w4_vals.get(aud)
                    v3 = w3_lookup.get(label, {}).get(aud) if do_wave else None

                    # W4 value
                    cell4 = ws.cell(row=current_row, column=w4_col,
                                    value=f"{v4}%" if isinstance(v4, int) else (v4 or "—"))
                    cell4.font = BODY_FONT
                    cell4.fill = row_fill
                    cell4.alignment = CTR
                    cell4.border = THIN_BORDER

                    if do_wave and w3_col:
                        cell3 = ws.cell(row=current_row, column=w3_col,
                                        value=f"{v3}%" if isinstance(v3, int) else (v3 or "—"))
                        cell3.font = W3_FONT
                        cell3.fill = _fill("F0F9FF")
                        cell3.alignment = CTR
                        cell3.border = THIN_BORDER

                        if d_col and isinstance(v4, int) and isinstance(v3, int):
                            delta = v4 - v3
                            delta_str = f"+{delta}pp" if delta > 0 else (f"{delta}pp" if delta != 0 else "—")
                            dc = ws.cell(row=current_row, column=d_col, value=delta_str)
                            dc.font = _f(bold=True, size=9,
                                         color="16A34A" if delta > 0 else ("DC2626" if delta < 0 else "6B7280"))
                            dc.fill = DELTA_FILL
                            dc.alignment = CTR
                            dc.border = THIN_BORDER
                        elif d_col:
                            dc = ws.cell(row=current_row, column=d_col, value="—")
                            dc.font = _f(size=9, color="9CA3AF")
                            dc.fill = DELTA_FILL
                            dc.alignment = CTR

                current_row += 1

            if not w4_data.get("rows"):
                nc = ws.cell(row=current_row, column=1, value="(no data extracted)")
                nc.font = _f(size=9, color="9CA3AF")
                current_row += 1

            current_row += 1  # spacer between slides

    # ── Flags tab ─────────────────────────────────────────────────────────────
    ws_f = wb.create_sheet(title="⚑ Flags")
    ws_f.column_dimensions["A"].width = 20
    ws_f.column_dimensions["B"].width = 60

    r = 1
    hdr = ws_f.cell(row=r, column=1, value="TYPE")
    hdr.font = HEADER_FONT; hdr.fill = HEADER_FILL
    hdr2 = ws_f.cell(row=r, column=2, value="PREFIX / NOTE")
    hdr2.font = HEADER_FONT; hdr2.fill = HEADER_FILL
    r += 1

    for p in sorted(missing_prefixes):
        c1 = ws_f.cell(row=r, column=1, value="MISSING")
        c1.font = _f(bold=True, size=9, color="991B1B")
        c1.fill = MISS_FILL
        c2 = ws_f.cell(row=r, column=2, value=p)
        c2.font = _f(size=9)
        r += 1

    for p in sorted(new_prefixes):
        c1 = ws_f.cell(row=r, column=1, value="NEW IN W4")
        c1.font = _f(bold=True, size=9, color="166534")
        c1.fill = NEW_FILL
        c2 = ws_f.cell(row=r, column=2, value=p)
        c2.font = _f(size=9)
        r += 1

    if not missing_prefixes and not new_prefixes:
        ws_f.cell(row=r, column=1, value="No flags — all prefixes matched.")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ── Streamlit UI ──────────────────────────────────────────────────────────────

def show_manifest_builder():
    st.markdown(
        "<div style='font-size:0.82rem;color:#374151;margin-bottom:12px'>"
        "Upload a slide manifest CSV and your W4 banner file(s). "
        "The builder will extract the specified metric per slide and write a structured Excel output."
        "</div>",
        unsafe_allow_html=True,
    )

    col_l, col_r = st.columns([1, 1])

    with col_l:
        st.markdown('<div class="step-label">Manifest CSV</div>', unsafe_allow_html=True)
        manifest_file = st.file_uploader(
            "Upload manifest CSV",
            type=["csv"],
            key="mb_manifest",
            label_visibility="collapsed",
        )

    with col_r:
        st.markdown('<div class="step-label">Format profile</div>', unsafe_allow_html=True)
        profile_names = [p for p in get_profile_names() if p != "+ Add new format"]
        mb_profile = st.selectbox(
            "Profile",
            profile_names,
            key="mb_profile",
            label_visibility="collapsed",
            help="Select the banner format profile. Use 'Corporate Reputation' for GQR-style banners.",
        )

    st.markdown('<div class="step-label" style="margin-top:12px">W4 Banner file(s)</div>', unsafe_allow_html=True)
    st.markdown(
        "<div style='font-size:0.76rem;color:#6B7280;margin-bottom:6px'>"
        "Upload one or more W4 banner Excel files. If audiences are split across files, upload all of them."
        "</div>",
        unsafe_allow_html=True,
    )
    w4_uploads = st.file_uploader(
        "W4 banners",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="mb_w4",
        label_visibility="collapsed",
    )

    with st.expander("W3 banner files (optional — for wave comparison)", expanded=False):
        st.markdown(
            "<div style='font-size:0.76rem;color:#6B7280;margin-bottom:6px'>"
            "Upload W3 banner file(s) to compute delta columns (W4 – W3)."
            "</div>",
            unsafe_allow_html=True,
        )
        w3_uploads = st.file_uploader(
            "W3 banners",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="mb_w3",
            label_visibility="collapsed",
        )

    if not manifest_file or not w4_uploads:
        st.info("Upload a manifest CSV and at least one W4 banner to continue.")
        return

    # Parse manifest
    manifest_rows = parse_manifest_csv(manifest_file.getvalue())
    if not manifest_rows:
        st.error("Could not parse the manifest CSV — check the file format.")
        return

    st.markdown(
        f'<span class="stat-pill">{len(manifest_rows)} slides</span>'
        f'<span class="stat-pill">{len({r["section"] for r in manifest_rows})} sections</span>'
        f'<span class="stat-pill">{len(w4_uploads)} W4 file(s)</span>'
        + (f'<span class="stat-pill">{len(w3_uploads)} W3 file(s)</span>' if w3_uploads else ""),
        unsafe_allow_html=True,
    )
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # Preview manifest
    with st.expander("Preview manifest", expanded=False):
        df_prev = pd.DataFrame([{
            "Slide": r["slide"],
            "Section": r["section"],
            "Title": r["title"][:50],
            "Prefix": r["prefix"],
            "Metric": r["metric"],
            "Audiences": ", ".join(r["audiences"]),
            "Wave": r["wave_compare"],
        } for r in manifest_rows])
        st.dataframe(df_prev, use_container_width=True, hide_index=True)

    if st.button("◈  Build manifest Excel", key="mb_run"):
        w4_files = [{"bytes": f.getvalue(), "label": f.name, "name": f.name}
                    for f in w4_uploads]
        w3_files = [{"bytes": f.getvalue(), "label": f.name, "name": f.name}
                    for f in w3_uploads] if w3_uploads else []

        with st.spinner("Scanning banners and extracting metrics…"):
            try:
                # Inventory all W4 prefixes
                all_w4_prefixes = set()
                for finfo in w4_files:
                    all_w4_prefixes |= get_all_prefixes(finfo["bytes"], mb_profile)

                # Determine flags
                manifest_prefixes = {r["prefix"] for r in manifest_rows if r["prefix"]}
                missing_prefixes  = manifest_prefixes - all_w4_prefixes
                new_prefixes      = all_w4_prefixes - manifest_prefixes

                if missing_prefixes:
                    st.warning(
                        f"⚠ {len(missing_prefixes)} prefix(es) in the manifest not found in W4 banner: "
                        + ", ".join(sorted(missing_prefixes))
                    )
                if new_prefixes:
                    st.info(
                        f"★ {len(new_prefixes)} prefix(es) in W4 banner not in manifest (new this wave): "
                        + ", ".join(sorted(new_prefixes))
                    )

                # Build Excel
                result_bytes = build_manifest_excel(
                    manifest_rows, w4_files, w3_files, mb_profile,
                    missing_prefixes, new_prefixes,
                )

                st.success(
                    f"Done — {len(manifest_rows)} slides across "
                    f"{len({r['section'] for r in manifest_rows})} sections"
                )
                st.download_button(
                    label="⬇  Download manifest Excel",
                    data=result_bytes,
                    file_name="manifest_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="mb_download",
                )

            except Exception as e:
                import traceback
                st.error(f"Build failed: {e}")
                st.code(traceback.format_exc())
