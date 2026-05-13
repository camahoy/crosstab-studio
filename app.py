"""
app.py — Crosstab Studio v1.2
Upload-first flow: file detection drives profile selection.
"""

import streamlit as st
import pandas as pd
from engine import (fast_scan, get_columns, generate_excel, generate_word,
                    detect_and_describe, read_toc, toc_to_groups,
                    deep_scan_file, save_user_profile)
from profiles import get_profile_names, get_profile

st.set_page_config(
    page_title="Crosstab Studio",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Sora:wght@300;400;600;700&display=swap');
html, body, [class*="css"] { font-family:'Sora',sans-serif; background:#F7F9FC; }

.cs-header { display:flex; align-items:baseline; gap:14px; margin-bottom:2rem;
             padding-bottom:1.25rem; border-bottom:2px solid #0F2D4A; }
.cs-logo { font-family:'Sora',sans-serif; font-weight:700; font-size:1.9rem;
           color:#0F2D4A; letter-spacing:-0.03em; }
.cs-logo span { color:#1A6EBD; }
.cs-sub { font-family:'DM Mono',monospace; font-size:0.72rem; color:#9CA3AF;
          letter-spacing:0.08em; text-transform:uppercase; }

.step-label { font-family:'DM Mono',monospace; font-size:0.68rem; font-weight:500;
              letter-spacing:0.1em; text-transform:uppercase; color:#1A6EBD; margin-bottom:6px; }

.detect-card { background:white; border:1.5px solid #E2E8F0; border-radius:10px;
               padding:1.25rem 1.5rem; margin-bottom:1rem; }
.detect-card.matched { border-color:#86EFAC; }
.detect-card.unmatched { border-color:#FECACA; }

.detect-title { font-weight:700; font-size:1rem; color:#0F2D4A; margin-bottom:0.75rem; }
.detect-row { display:flex; gap:10px; align-items:flex-start; padding:5px 0;
              border-bottom:1px solid #F1F5F9; font-size:0.82rem; }
.detect-key { color:#6B7280; min-width:180px; font-family:'DM Mono',monospace;
              font-size:0.73rem; padding-top:1px; }
.detect-val { color:#0F2D4A; font-weight:500; flex:1; }
.detect-val.ok  { color:#16A34A; }
.detect-val.warn { color:#D97706; }

.stat-pill { display:inline-block; background:#EFF6FF; color:#1A6EBD;
             font-family:'DM Mono',monospace; font-size:0.7rem;
             padding:2px 10px; border-radius:99px; margin-right:4px; }

.type-badge { display:inline-block; font-family:'DM Mono',monospace; font-size:0.65rem;
              padding:1px 6px; border-radius:4px; margin-left:4px; }
.type-standard     { background:#DCFCE7; color:#166534; }
.type-t2b          { background:#DBEAFE; color:#1D4ED8; }
.type-b2b          { background:#FEF3C7; color:#B45309; }
.type-summary_grid { background:#F3E8FF; color:#7C3AED; }
.type-mean         { background:#FFE4E6; color:#BE123C; }

.unmatched-box { background:#FEF2F2; border:1px solid #FECACA; border-radius:8px;
                 padding:1rem 1.25rem; color:#991B1B; font-size:0.88rem; margin:0.5rem 0; }
.coming-soon { background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px;
               padding:1rem 1.25rem; color:#713F12; font-size:0.88rem; }

.stButton > button { background:#0F2D4A !important; color:white !important;
                     border:none !important; border-radius:6px !important;
                     font-family:'Sora',sans-serif !important; font-weight:600 !important;
                     font-size:0.88rem !important; padding:0.5rem 1.5rem !important; }
.stButton > button:hover { background:#1A6EBD !important; }
hr { border:none; border-top:1px solid #E2E8F0; margin:1.5rem 0; }
#MainMenu, footer, header { visibility:hidden; }
.block-container { padding-top:2rem; max-width:1200px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="cs-header">
    <div class="cs-logo">Crosstab<span> Studio</span></div>
    <div class="cs-sub">Research output formatter</div>
</div>
""", unsafe_allow_html=True)

# ── State ─────────────────────────────────────────────────────
for k, v in {
    'files': [], 'detected': None, 'confirmed_profile': None,
    'question_groups': [], 'columns': [], 'selected_qs': set(),
    'selected_cols': [], 'scan_done': False,
    'include_types': {'standard', 't2b', 'b2b', 'summary_grid', 'mean'},
    'deep_scan_result': None,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

TYPE_LABELS = {
    'standard':     ('Standard', 'standard'),
    't2b':          ('T2B',      't2b'),
    'b2b':          ('B2B',      'b2b'),
    'summary_grid': ('Grid',     'summary_grid'),
    'mean':         ('Mean',     'mean'),
}

# ═══════════════════════════════════════════════════
# STEP 1 — Upload
# ═══════════════════════════════════════════════════
st.markdown('<div class="step-label">Step 1 — Upload your file</div>', unsafe_allow_html=True)
st.markdown(
    "<div style='font-size:0.8rem;color:#6B7280;margin-bottom:10px'>"
    "Upload one or more files. If comparing waves or subgroups, upload all files here "
    "and label each one (e.g. W1, W2, W3 or Gen Pop, Tech Elite)."
    "</div>",
    unsafe_allow_html=True
)

uploaded_files = st.file_uploader(
    "Upload", type=["xlsx","xls"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    files_changed = len(uploaded_files) != len(st.session_state.files)
    if files_changed:
        st.session_state.files             = []
        st.session_state.detected          = None
        st.session_state.confirmed_profile = None
        st.session_state.scan_done         = False
        st.session_state.selected_qs       = set()
        st.session_state.deep_scan_result  = None

    new_files = []
    label_col, _ = st.columns([1.5, 1])
    with label_col:
        for i, uf in enumerate(uploaded_files):
            default = f"W{i+1}" if len(uploaded_files) > 1 else uf.name.replace('.xlsx','')[:25]
            label   = st.text_input(
                f"Label — {uf.name[:40]}",
                value=default,
                key=f"lbl_{i}_{uf.name}",
            )
            new_files.append({'bytes': uf.getvalue(), 'label': label, 'name': uf.name})
    st.session_state.files = new_files

    # ── Auto-detect on first file ──────────────────
    if st.session_state.detected is None and new_files:
        with st.spinner("Detecting format..."):
            st.session_state.detected = detect_and_describe(new_files[0]['bytes'])

st.markdown('<hr>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════
# STEP 2 — Format detection result
# ═══════════════════════════════════════════════════
if st.session_state.files and st.session_state.detected is not None:
    det = st.session_state.detected
    st.markdown('<div class="step-label">Step 2 — Confirm format</div>', unsafe_allow_html=True)

    matched_profile = det.get('matched_profile')
    findings        = det.get('findings', [])
    sample_cols     = det.get('sample_columns', [])
    n_sheets        = det.get('n_sheets', 0)

    card_class = "matched" if matched_profile else "unmatched"

    st.markdown(f'<div class="detect-card {card_class}">', unsafe_allow_html=True)

    if matched_profile:
        p = get_profile(matched_profile)
        st.markdown(
            f'<div class="detect-title">✓ Format detected</div>',
            unsafe_allow_html=True
        )
        # Show what was found
        for key, val, status in findings:
            cls = 'ok' if status == 'ok' else 'warn' if status == 'warn' else ''
            st.markdown(
                f'<div class="detect-row">'
                f'<span class="detect-key">{key}</span>'
                f'<span class="detect-val {cls}">{val}</span>'
                f'</div>',
                unsafe_allow_html=True
            )
        if sample_cols:
            st.markdown(
                f'<div class="detect-row">'
                f'<span class="detect-key">Columns found</span>'
                f'<span class="detect-val">{" · ".join(sample_cols[:8])}'
                f'{"  +" + str(len(sample_cols)-8) + " more" if len(sample_cols) > 8 else ""}'
                f'</span></div>',
                unsafe_allow_html=True
            )
        st.markdown('</div>', unsafe_allow_html=True)

        col_confirm, col_override = st.columns([2, 1])
        with col_confirm:
            if st.session_state.confirmed_profile != matched_profile:
                if st.button(f"✓  Confirm — use this format"):
                    st.session_state.confirmed_profile = matched_profile
                    st.session_state.scan_done         = False
                    st.session_state.selected_qs       = set()
                    st.rerun()
        with col_override:
            other = st.selectbox(
                "Override profile",
                ["— use detected —"] + [p for p in get_profile_names() if p != matched_profile and p != "+ Add new format"],
                label_visibility="collapsed",
                key="profile_override",
            )
            if other != "— use detected —" and st.button("Apply override"):
                st.session_state.confirmed_profile = other
                st.session_state.scan_done         = False
                st.session_state.selected_qs       = set()
                st.rerun()

    else:
        st.markdown(
            '<div class="detect-title">✗ Format not recognised — running full file scan…</div>',
            unsafe_allow_html=True
        )
        for key, val, status in findings:
            st.markdown(
                f'<div class="detect-row">'
                f'<span class="detect-key">{key}</span>'
                f'<span class="detect-val">{val}</span>'
                f'</div>',
                unsafe_allow_html=True
            )
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Deep scan (runs once, result cached in session state) ──
        if st.session_state.deep_scan_result is None:
            with st.spinner("Scanning all sheets to learn the file structure…"):
                st.session_state.deep_scan_result = deep_scan_file(
                    st.session_state.files[0]['bytes']
                )

        scan = st.session_state.deep_scan_result or {}

        if scan.get('error'):
            st.error(f"Scan error: {scan['error']}")
        elif scan.get('dominant'):
            dom = scan['dominant']

            # ── Scan summary ──────────────────────────────────────
            sc1, sc2 = st.columns(2)
            with sc1:
                st.markdown(
                    f"<div style='font-size:0.82rem;color:#374151;margin-top:8px'>"
                    f"<b>Scan complete</b> — {scan['sheet_count']} sheets total, "
                    f"{len(scan['sampled'])} sampled"
                    f"{'  ·  ' + scan['toc_summary'] if scan['has_toc'] else ''}"
                    f"</div>",
                    unsafe_allow_html=True
                )
            with sc2:
                if scan['variant_count'] > 0:
                    st.markdown(
                        f"<div style='font-size:0.82rem;color:#D97706;margin-top:8px'>"
                        f"⚠ {scan['variant_count']} variant layout(s) found in "
                        f"{len(scan['outlier_sheets'])} sheet(s) — "
                        f"dominant pattern covers the rest"
                        f"</div>",
                        unsafe_allow_html=True
                    )

            # ── Row preview ───────────────────────────────────────
            if scan.get('row_preview'):
                with st.expander("Row preview — first data sheet (rows 1–15)", expanded=False):
                    preview = scan['row_preview']
                    df_prev = pd.DataFrame(
                        preview,
                        index=[f"Row {i+1}" for i in range(len(preview))],
                        columns=[f"Col {j+1}" for j in range(len(preview[0]) if preview else 0)],
                    )
                    st.dataframe(df_prev, use_container_width=True)

            # ── Format wizard ─────────────────────────────────────
            st.markdown(
                "<div style='font-size:0.88rem;font-weight:600;color:#0F2D4A;"
                "margin:18px 0 6px'>Confirm file structure</div>"
                "<div style='font-size:0.78rem;color:#6B7280;margin-bottom:12px'>"
                "Row numbers as shown in Excel (Row 1 = first row of spreadsheet). "
                "Pre-filled from scan — adjust if anything looks off.</div>",
                unsafe_allow_html=True
            )

            with st.form("format_wizard"):
                wc1, wc2 = st.columns(2)

                # Display 1-indexed (Excel rows), store 0-indexed in profile
                with wc1:
                    q_row_in = st.number_input(
                        "Question wording row",
                        min_value=1, max_value=60,
                        value=(dom['question_row'] or 2) + 1,
                        help="Row containing the question text",
                    )
                    h_row_in = st.number_input(
                        "Column headers row (where 'Total' appears)",
                        min_value=1, max_value=60,
                        value=(dom['header_row'] or 4) + 1,
                    )
                    b_row_in = st.number_input(
                        "Base / N row",
                        min_value=1, max_value=60,
                        value=(dom['base_row'] or 7) + 1,
                    )

                with wc2:
                    d_start_in = st.number_input(
                        "Data starts at row",
                        min_value=1, max_value=60,
                        value=(dom['data_start'] or 9) + 1,
                    )
                    d_step_in = st.number_input(
                        "Rows per answer choice",
                        min_value=1, max_value=10,
                        value=dom.get('data_step', 3),
                        help="Usually 3 (label / value / sig-code)",
                    )
                    col_start_in = st.number_input(
                        "'Total' column (1 = column A)",
                        min_value=1, max_value=20,
                        value=(dom.get('col_start', 1) or 1) + 1,
                    )

                stop_raw  = st.text_input(
                    "Stop parsing when this text appears in column A",
                    value=", ".join(dom.get('stop_on') or ['sigma']),
                    help="Comma-separated. E.g.: sigma, back to top",
                )
                skip_toc  = st.checkbox(
                    "Skip first sheet (it's a table of contents)",
                    value=scan['has_toc'],
                )

                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                pname_in  = st.text_input(
                    "Save this format as",
                    placeholder="e.g.  MyClient Tracker",
                    help="Saved to user_profiles.json — available in all future sessions",
                )
                submitted = st.form_submit_button("◈  Save format and use it")

                if submitted:
                    pname = pname_in.strip()
                    if not pname:
                        st.warning("Please give this format a name before saving.")
                    else:
                        stop_list = [s.strip().lower() for s in stop_raw.split(',') if s.strip()]
                        new_profile = {
                            "description":      f"User-defined — {pname}",
                            "specs":            [],
                            "question_row":     q_row_in - 1,
                            "header_row":       h_row_in - 1,
                            "base_row":         b_row_in - 1,
                            "data_start":       d_start_in - 1,
                            "data_step":        d_step_in,
                            "value_row_offset": 1,
                            "skip_sheet_0":     skip_toc,
                            "column_start":     col_start_in - 1,
                            "stop_on":          stop_list or ['sigma'],
                            "coerce_strings":   False,
                            "multi_file_mode":  "waves",
                        }
                        save_user_profile(pname, new_profile)
                        st.session_state.confirmed_profile = pname
                        st.session_state.scan_done         = False
                        st.session_state.selected_qs       = set()
                        st.success(f"Profile '{pname}' saved — continuing with scan.")
                        st.rerun()

        # ── Manual override (always available) ────────────────
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        manual = st.selectbox(
            "Or try an existing profile:",
            ["— select —"] + [p for p in get_profile_names()
                              if p != "+ Add new format"],
            label_visibility="collapsed",
        )
        if manual != "— select —" and st.button("Try this profile"):
            st.session_state.confirmed_profile = manual
            st.session_state.scan_done         = False
            st.rerun()

    st.markdown('<hr>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════
# STEP 3 — Scan
# ═══════════════════════════════════════════════════
confirmed = st.session_state.confirmed_profile
if confirmed and st.session_state.files and not st.session_state.scan_done:
    p    = get_profile(confirmed)
    mode = p.get("multi_file_mode", "waves") if p else "waves"
    st.markdown(
        f"<div style='font-size:0.82rem;color:#374151;margin-bottom:12px'>"
        f"Ready to scan using <strong>{confirmed}</strong> profile. "
        f"{'Multiple files will be compared as waves.' if len(st.session_state.files) > 1 else ''}"
        f"</div>",
        unsafe_allow_html=True
    )
    if st.button("◈  Scan file"):
        with st.spinner("Scanning..."):
            try:
                files = st.session_state.files

                # Build question list from first file
                ref = files[0]['bytes']
                toc = read_toc(ref, confirmed)
                groups = toc_to_groups(toc) if toc else fast_scan(ref, confirmed)
                cols   = get_columns(ref, confirmed)

                # For multi-file: merge questions from all other files
                # so questions that only exist in wave 2+ are still shown
                if len(files) > 1:
                    seen = {g['prefix'] for g in groups}
                    for finfo in files[1:]:
                        extra_toc = read_toc(finfo['bytes'], confirmed)
                        extra = (toc_to_groups(extra_toc) if extra_toc
                                 else fast_scan(finfo['bytes'], confirmed))
                        for g in extra:
                            if g['prefix'] not in seen:
                                groups.append(g)
                                seen.add(g['prefix'])

                st.session_state.question_groups = groups
                st.session_state.columns         = cols
                st.session_state.scan_done       = True
                st.session_state.selected_cols   = [j for j, g, s in cols]
                st.rerun()
            except Exception as e:
                st.error(f"Scan failed: {e}")
                import traceback; st.code(traceback.format_exc())

# ═══════════════════════════════════════════════════
# STEP 4 — Select questions + columns
# ═══════════════════════════════════════════════════
if st.session_state.scan_done:
    groups = st.session_state.question_groups
    cols   = st.session_state.columns
    n_f    = len(st.session_state.files)

    st.markdown(
        f'<span class="stat-pill">{len(groups)} questions</span>'
        f'<span class="stat-pill">{sum(len(g["sheets"]) for g in groups)} sheets</span>'
        f'<span class="stat-pill">{n_f} file{"s" if n_f>1 else ""}</span>'
        f'<span class="stat-pill">{len(cols)} columns</span>',
        unsafe_allow_html=True
    )

    # Sheet-type filter
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.markdown('<div class="step-label">Sheet types to include</div>', unsafe_allow_html=True)
    type_opts = [
        ('standard',     'Standard'),
        ('t2b',          'T2B Summary'),
        ('b2b',          'B2B Summary'),
        ('summary_grid', 'Grid'),
        ('mean',         'Mean'),
    ]
    type_cols = st.columns(len(type_opts))
    include_types = set()
    for col_el, (key, label) in zip(type_cols, type_opts):
        with col_el:
            if st.checkbox(label, value=(key in st.session_state.include_types), key=f"type_{key}"):
                include_types.add(key)
    st.session_state.include_types = include_types
    st.markdown("<br>", unsafe_allow_html=True)

    left, right = st.columns([1.8, 1])

    with left:
        st.markdown('<div class="step-label">Step 3 — Select questions</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1,1,2])
        with c1:
            if st.button("Select all"):
                st.session_state.selected_qs = {g['prefix'] for g in groups}
                # Clear checkbox widget keys so value= takes effect on rerun
                for g in groups:
                    st.session_state.pop(f"q_{g['prefix']}", None)
                st.rerun()
        with c2:
            if st.button("Clear all"):
                st.session_state.selected_qs = set()
                for g in groups:
                    st.session_state.pop(f"q_{g['prefix']}", None)
                st.rerun()
        with c3:
            search = st.text_input("Search", placeholder="Filter...", label_visibility="collapsed")

        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

        filtered = [g for g in groups if not search
                    or search.lower() in g['wording'].lower()
                    or search.lower() in g['prefix'].lower()]

        for g in filtered:
            prefix  = g['prefix']
            wording = g['wording']
            n_s     = len(g['sheets'])
            types   = g.get('types', [])
            checked = prefix in st.session_state.selected_qs

            new_val = st.checkbox(
                f"**{prefix}** — {wording[:65]}{'…' if len(wording)>65 else ''} *({n_s})*",
                value=checked, key=f"q_{prefix}",
            )
            if types:
                badges = ''.join(
                    f'<span class="type-badge type-{TYPE_LABELS.get(t,("",t))[1]}">'
                    f'{TYPE_LABELS.get(t,(t,t))[0]}</span>'
                    for t in types
                )
                st.markdown(
                    f'<div style="margin-top:-12px;margin-bottom:4px;padding-left:28px">{badges}</div>',
                    unsafe_allow_html=True
                )
            if new_val != checked:
                if new_val: st.session_state.selected_qs.add(prefix)
                else:       st.session_state.selected_qs.discard(prefix)

    with right:
        st.markdown('<div class="step-label">Step 4 — Select columns</div>', unsafe_allow_html=True)
        st.markdown("<div style='font-size:0.78rem;color:#6B7280;margin-bottom:8px'>Subgroups / countries to include</div>", unsafe_allow_html=True)
        sel_cols = []
        for j, g, s in cols:
            label   = g if not s or s.lower() == g.lower() else f"{g} — {s}"
            checked = j in st.session_state.selected_cols
            if st.checkbox(label, value=checked, key=f"col_{j}"):
                sel_cols.append(j)
        st.session_state.selected_cols = sel_cols

    # ═══════════════════════════════════════════════
    # STEP 5 — Export
    # ═══════════════════════════════════════════════
    st.markdown('<hr>', unsafe_allow_html=True)
    st.markdown('<div class="step-label">Step 5 — Export</div>', unsafe_allow_html=True)

    n_sel     = len(st.session_state.selected_qs)
    n_col_sel = len(st.session_state.selected_cols)

    if n_f > 1:
        st.markdown(
            f"<div style='font-size:0.82rem;color:#374151;margin-bottom:8px'>"
            f"Wave comparison: {', '.join(f['label'] for f in st.session_state.files)}"
            f" — each wave as a separate colour-coded table per question."
            f"</div>",
            unsafe_allow_html=True
        )

    if n_sel == 0:
        st.info("Select at least one question above.")
    elif n_col_sel == 0:
        st.info("Select at least one column.")
    else:
        st.markdown(
            f'<span class="stat-pill">{n_sel} questions selected</span>'
            f'<span class="stat-pill">{n_col_sel} columns selected</span>',
            unsafe_allow_html=True
        )
        st.markdown("<br>", unsafe_allow_html=True)

        # Export format choice
        export_fmt = st.radio(
            "Export format",
            ["Excel", "Media Release Template (Word)"],
            horizontal=True,
            label_visibility="collapsed",
        )

        survey_title = ''
        if export_fmt == "Media Release Template (Word)":
            survey_title = st.text_input(
                "Survey title (appears in document header)",
                placeholder="e.g. College Student Fall Mental Wellness Survey",
            )

        btn_label = (
            f"◈  Generate Excel ({n_sel} questions)"
            if export_fmt == "Excel"
            else f"◈  Generate Word ({n_sel} questions)"
        )

        if st.button(btn_label):
            selected_groups = [g for g in groups if g['prefix'] in st.session_state.selected_qs]
            col_indices     = st.session_state.selected_cols
            col_names       = [g for j,g,s in cols if j in col_indices]

            with st.spinner("Building tables..."):
                try:
                    if export_fmt == "Excel":
                        result_bytes = generate_excel(
                            selected_groups,
                            st.session_state.files,
                            confirmed,
                            col_indices,
                            col_names,
                            include_types=st.session_state.include_types or None,
                        )
                        st.success(f"Done — {n_sel} questions exported")
                        st.download_button(
                            label="⬇  Download Excel",
                            data=result_bytes,
                            file_name="crosstab_studio_export.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    else:
                        result_bytes, err = generate_word(
                            selected_groups,
                            st.session_state.files,
                            confirmed,
                            col_indices,
                            col_names,
                            survey_title=survey_title,
                            include_types=st.session_state.include_types or None,
                        )
                        if err:
                            st.error(f"Word export failed: {err}")
                            st.info("Make sure template_doc.docx is in the app root directory.")
                        else:
                            st.success(f"Done — {n_sel} questions exported")
                            st.download_button(
                                label="⬇  Download Word",
                                data=result_bytes,
                                file_name="media_release.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            )
                except Exception as e:
                    st.error(f"Export failed: {e}")
                    import traceback; st.code(traceback.format_exc())
