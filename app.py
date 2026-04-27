"""
app.py — Crosstab Studio v1.1
"""

import streamlit as st
from engine import fast_scan, get_columns, generate_excel, validate_format
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
html, body, [class*="css"] { font-family: 'Sora', sans-serif; background: #F7F9FC; }

.cs-header { display:flex; align-items:baseline; gap:14px; margin-bottom:2rem;
             padding-bottom:1.25rem; border-bottom:2px solid #0F2D4A; }
.cs-logo { font-family:'Sora',sans-serif; font-weight:700; font-size:1.9rem;
           color:#0F2D4A; letter-spacing:-0.03em; }
.cs-logo span { color:#1A6EBD; }
.cs-sub { font-family:'DM Mono',monospace; font-size:0.72rem; color:#9CA3AF;
          letter-spacing:0.08em; text-transform:uppercase; }

.step-label { font-family:'DM Mono',monospace; font-size:0.68rem; font-weight:500;
              letter-spacing:0.1em; text-transform:uppercase; color:#1A6EBD; margin-bottom:6px; }

.spec-row { display:flex; gap:8px; align-items:center; padding:4px 0;
            border-bottom:1px solid #F1F5F9; font-size:0.82rem; }
.spec-key { color:#6B7280; min-width:140px; font-family:'DM Mono',monospace; font-size:0.75rem; }
.spec-val { color:#0F2D4A; font-weight:500; }

.stat-pill { display:inline-block; background:#EFF6FF; color:#1A6EBD;
             font-family:'DM Mono',monospace; font-size:0.7rem;
             padding:2px 10px; border-radius:99px; margin-right:4px; }

.type-badge { display:inline-block; font-family:'DM Mono',monospace; font-size:0.65rem;
              padding:1px 6px; border-radius:4px; margin-left:4px; }
.type-standard  { background:#DCFCE7; color:#166534; }
.type-t2b       { background:#DBEAFE; color:#1D4ED8; }
.type-b2b       { background:#FEF3C7; color:#B45309; }
.type-summary_grid { background:#F3E8FF; color:#7C3AED; }
.type-mean      { background:#FFE4E6; color:#BE123C; }

.valid-ok   { background:#DCFCE7; border:1px solid #86EFAC; border-radius:8px;
              padding:0.75rem 1rem; color:#166534; font-size:0.88rem; }
.valid-fail { background:#FEF2F2; border:1px solid #FECACA; border-radius:8px;
              padding:0.75rem 1rem; color:#991B1B; font-size:0.88rem; }
.valid-warn { background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px;
              padding:0.75rem 1rem; color:#92400E; font-size:0.88rem; }

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

# ── Header ────────────────────────────────────────────────────
st.markdown("""
<div class="cs-header">
    <div class="cs-logo">Crosstab<span> Studio</span></div>
    <div class="cs-sub">Research output formatter</div>
</div>
""", unsafe_allow_html=True)

# ── State ─────────────────────────────────────────────────────
for k, v in {
    'profile_name': None, 'files': [], 'question_groups': [],
    'columns': [], 'selected_qs': set(), 'selected_cols': [],
    'scan_done': False, 'validation': None,
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
# STEP 1 — Profile + Specs
# ═══════════════════════════════════════════════════
col_profile, col_specs = st.columns([1, 1.4])

with col_profile:
    st.markdown('<div class="step-label">Step 1 — Format profile</div>', unsafe_allow_html=True)
    chosen = st.selectbox("Profile", get_profile_names(), label_visibility="collapsed")

    if chosen != st.session_state.profile_name:
        st.session_state.profile_name = chosen
        st.session_state.scan_done    = False
        st.session_state.selected_qs  = set()
        st.session_state.files        = []
        st.session_state.validation   = None

with col_specs:
    if chosen and chosen != "+ Add new format":
        p = get_profile(chosen)
        st.markdown('<div class="step-label">Profile specs</div>', unsafe_allow_html=True)
        for key, val in p.get("specs", []):
            st.markdown(
                f'<div class="spec-row"><span class="spec-key">{key}</span>'
                f'<span class="spec-val">{val}</span></div>',
                unsafe_allow_html=True
            )
        mode = p.get("multi_file_mode", "waves")
        st.markdown(
            f'<div class="spec-row"><span class="spec-key">Multi-file mode</span>'
            f'<span class="spec-val">{"Wave comparison" if mode == "waves" else "Subgroup comparison"}</span></div>',
            unsafe_allow_html=True
        )

if chosen == "+ Add new format":
    st.markdown("""
    <div class="coming-soon">
        <strong>Custom format builder — coming soon</strong><br>
        Define your own row structure, save it with a name, and reuse it across sessions.
    </div>""", unsafe_allow_html=True)
    st.stop()

st.markdown('<hr>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════
# STEP 2 — Upload
# ═══════════════════════════════════════════════════
profile    = get_profile(chosen)
mode       = profile.get("multi_file_mode", "waves")
mode_label = "wave" if mode == "waves" else "subgroup"

st.markdown('<div class="step-label">Step 2 — Upload files</div>', unsafe_allow_html=True)
st.markdown(
    f"<div style='font-size:0.8rem;color:#6B7280;margin-bottom:10px'>"
    f"Upload one or more files. Each file = one {mode_label}. "
    f"Label each one ({'W1, W2, W3' if mode == 'waves' else 'Gen Pop, Tech Elite'} etc)."
    f"</div>",
    unsafe_allow_html=True
)

uploaded_files = st.file_uploader(
    "Upload", type=["xlsx","xls"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    if len(uploaded_files) != len(st.session_state.files):
        st.session_state.files     = []
        st.session_state.scan_done = False
        st.session_state.validation = None

    labels_col, _ = st.columns([1.5, 1])
    with labels_col:
        new_files = []
        for i, uf in enumerate(uploaded_files):
            default = f"W{i+1}" if mode == "waves" else uf.name.replace('.xlsx','')[:20]
            label   = st.text_input(
                f"Label — {uf.name[:40]}",
                value=default,
                key=f"lbl_{i}_{uf.name}",
            )
            new_files.append({'bytes': uf.getvalue(), 'label': label})
        st.session_state.files = new_files

    # ── Format validation ──────────────────────────
    if st.session_state.validation is None and st.session_state.files:
        matched, confidence = validate_format(st.session_state.files[0]['bytes'])
        st.session_state.validation = (matched, confidence)

    matched, confidence = st.session_state.validation or (None, 0)

    if matched == chosen:
        st.markdown(
            f'<div class="valid-ok">✓ File recognised as <strong>{matched}</strong> '
            f'(confidence {confidence}%)</div>',
            unsafe_allow_html=True
        )
    elif matched and matched != chosen:
        st.markdown(
            f'<div class="valid-warn">⚠ File looks like <strong>{matched}</strong> '
            f'but <strong>{chosen}</strong> is selected. '
            f'Consider switching profile.</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            '<div class="valid-fail">✗ Format not recognised. If this is a new banner type, '
            'share the file structure with your admin to add support.</div>',
            unsafe_allow_html=True
        )

# ── Scan button ────────────────────────────────────
if st.session_state.files and not st.session_state.scan_done:
    st.markdown('<hr>', unsafe_allow_html=True)
    if st.button("◈  Scan files"):
        with st.spinner("Scanning..."):
            try:
                ref   = st.session_state.files[0]['bytes']
                groups = fast_scan(ref, chosen)
                cols   = get_columns(ref, chosen)
                st.session_state.question_groups = groups
                st.session_state.columns         = cols
                st.session_state.scan_done       = True
                st.session_state.selected_cols   = [j for j,g,s in cols]
                st.rerun()
            except Exception as e:
                st.error(f"Scan failed: {e}")
                import traceback; st.code(traceback.format_exc())

# ═══════════════════════════════════════════════════
# STEP 3 — Select questions + columns
# ═══════════════════════════════════════════════════
if st.session_state.scan_done:
    groups = st.session_state.question_groups
    cols   = st.session_state.columns
    n_f    = len(st.session_state.files)

    st.markdown('<hr>', unsafe_allow_html=True)
    st.markdown(
        f'<span class="stat-pill">{len(groups)} questions</span>'
        f'<span class="stat-pill">{sum(len(g["sheets"]) for g in groups)} sheets</span>'
        f'<span class="stat-pill">{n_f} file{"s" if n_f>1 else ""}</span>'
        f'<span class="stat-pill">{len(cols)} columns</span>',
        unsafe_allow_html=True
    )
    st.markdown("<br>", unsafe_allow_html=True)

    left, right = st.columns([1.8, 1])

    with left:
        st.markdown('<div class="step-label">Step 3 — Select questions</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1,1,2])
        with c1:
            if st.button("Select all"):
                st.session_state.selected_qs = {g['prefix'] for g in groups}
                st.rerun()
        with c2:
            if st.button("Clear all"):
                st.session_state.selected_qs = set()
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

            # Build type badges HTML
            badges = ''
            for t in types:
                label, cls = TYPE_LABELS.get(t, (t, 'standard'))
                badges += f'<span class="type-badge type-{cls}">{label}</span>'

            # Checkbox label
            short   = f"{wording[:65]}{'…' if len(wording)>65 else ''}"
            new_val = st.checkbox(
                f"**{prefix}** — {short} *({n_s})*",
                value=checked,
                key=f"q_{prefix}",
            )
            if badges:
                st.markdown(
                    f'<div style="margin-top:-12px;margin-bottom:4px;padding-left:28px">'
                    f'{badges}</div>',
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
    # STEP 5 — Generate
    # ═══════════════════════════════════════════════
    st.markdown('<hr>', unsafe_allow_html=True)
    st.markdown('<div class="step-label">Step 5 — Export</div>', unsafe_allow_html=True)

    n_sel     = len(st.session_state.selected_qs)
    n_col_sel = len(st.session_state.selected_cols)
    n_f       = len(st.session_state.files)

    if n_f > 1:
        st.markdown(
            f"<div style='font-size:0.82rem;color:#374151;margin-bottom:8px'>"
            f"Wave comparison: {', '.join(f['label'] for f in st.session_state.files)}"
            f" — each wave output as a separate colour-coded table per question."
            f"</div>",
            unsafe_allow_html=True
        )

    if n_sel == 0:
        st.info("Select at least one question above.")
    elif n_col_sel == 0:
        st.info("Select at least one column.")
    else:
        st.markdown(
            f'<span class="stat-pill">{n_sel} questions</span>'
            f'<span class="stat-pill">{n_col_sel} columns</span>',
            unsafe_allow_html=True
        )
        st.markdown("<br>", unsafe_allow_html=True)

        if st.button(f"◈  Generate Excel ({n_sel} questions)"):
            selected_groups = [g for g in groups if g['prefix'] in st.session_state.selected_qs]
            col_indices     = st.session_state.selected_cols
            col_names       = [g for j,g,s in cols if j in col_indices]

            with st.spinner(f"Building tables..."):
                try:
                    excel_bytes = generate_excel(
                        selected_groups,
                        st.session_state.files,
                        chosen,
                        col_indices,
                        col_names,
                    )
                    st.success(f"Done — {n_sel} questions exported")
                    st.download_button(
                        label="⬇  Download Excel",
                        data=excel_bytes,
                        file_name="crosstab_studio_export.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"Export failed: {e}")
                    import traceback; st.code(traceback.format_exc())
