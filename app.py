"""
app.py — Crosstab Studio
A workflow tool for reformatting market research exports into presentation-ready tables.
"""

import io
import streamlit as st
import pandas as pd
from engine import fast_scan, get_columns, generate_excel
from profiles import get_profile_names, get_profile

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="Crosstab Studio",
    page_icon="◈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Styles ────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Sora:wght@300;400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Sora', sans-serif;
    background-color: #F7F9FC;
}

/* Header */
.crossstudio-header {
    display: flex;
    align-items: baseline;
    gap: 12px;
    margin-bottom: 2rem;
    padding-bottom: 1.5rem;
    border-bottom: 2px solid #0F2D4A;
}
.crossstudio-logo {
    font-family: 'Sora', sans-serif;
    font-weight: 700;
    font-size: 2rem;
    color: #0F2D4A;
    letter-spacing: -0.04em;
}
.crossstudio-logo span {
    color: #1A6EBD;
}
.crossstudio-sub {
    font-family: 'DM Mono', monospace;
    font-size: 0.75rem;
    color: #6B7280;
    letter-spacing: 0.05em;
    text-transform: uppercase;
}

/* Steps */
.step-label {
    font-family: 'DM Mono', monospace;
    font-size: 0.7rem;
    font-weight: 500;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #1A6EBD;
    margin-bottom: 0.4rem;
}

/* Profile card */
.profile-card {
    background: white;
    border: 1.5px solid #E2E8F0;
    border-radius: 8px;
    padding: 1rem 1.25rem;
    margin-bottom: 0.5rem;
    cursor: pointer;
    transition: border-color 0.15s;
}
.profile-card:hover { border-color: #1A6EBD; }
.profile-name { font-weight: 600; color: #0F2D4A; font-size: 0.95rem; }
.profile-desc { font-size: 0.8rem; color: #6B7280; margin-top: 0.2rem; }

/* Question groups */
.q-group {
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 6px;
    padding: 0.6rem 1rem;
    margin-bottom: 4px;
    display: flex;
    align-items: center;
    gap: 10px;
}
.q-prefix {
    font-family: 'DM Mono', monospace;
    font-size: 0.8rem;
    font-weight: 500;
    color: #1A6EBD;
    min-width: 60px;
}
.q-wording { font-size: 0.88rem; color: #374151; flex: 1; }
.q-count {
    font-family: 'DM Mono', monospace;
    font-size: 0.72rem;
    color: #9CA3AF;
    white-space: nowrap;
}

/* Stat pill */
.stat-pill {
    display: inline-block;
    background: #EFF6FF;
    color: #1A6EBD;
    font-family: 'DM Mono', monospace;
    font-size: 0.72rem;
    padding: 2px 8px;
    border-radius: 99px;
    margin-right: 4px;
}

/* Buttons */
.stButton > button {
    background: #0F2D4A;
    color: white;
    border: none;
    border-radius: 6px;
    font-family: 'Sora', sans-serif;
    font-weight: 600;
    font-size: 0.88rem;
    padding: 0.5rem 1.5rem;
    transition: background 0.15s;
}
.stButton > button:hover { background: #1A6EBD; }

/* Divider */
hr { border: none; border-top: 1px solid #E2E8F0; margin: 1.5rem 0; }

/* Hide streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 2rem; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────
st.markdown("""
<div class="crossstudio-header">
    <div class="crossstudio-logo">Cross<span>tab Studio</span></div>
    <div class="crossstudio-sub">Research output formatter</div>
</div>
""", unsafe_allow_html=True)

# ── State init ────────────────────────────────────────────────
for key, default in [
    ('file_bytes',    None),
    ('profile_name',  None),
    ('question_groups', []),
    ('columns',       []),
    ('selected_qs',   set()),
    ('selected_cols', []),
    ('scan_done',     False),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ═══════════════════════════════════════════════════════════
# STEP 1 — Upload + Profile
# ═══════════════════════════════════════════════════════════
col1, col2 = st.columns([1.2, 1])

with col1:
    st.markdown('<div class="step-label">Step 1 — Upload file</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Upload your Excel banner or crosstab",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
    )
    if uploaded:
        fb = uploaded.read()
        if fb != st.session_state.file_bytes:
            st.session_state.file_bytes   = fb
            st.session_state.scan_done    = False
            st.session_state.selected_qs  = set()
            st.session_state.selected_cols = []

with col2:
    st.markdown('<div class="step-label">Step 2 — Select format profile</div>', unsafe_allow_html=True)
    profile_names = get_profile_names()
    chosen = st.radio(
        "Profile",
        profile_names,
        label_visibility="collapsed",
    )
    if chosen != st.session_state.profile_name:
        st.session_state.profile_name = chosen
        st.session_state.scan_done    = False
        st.session_state.selected_qs  = set()
        st.session_state.selected_cols = []

    if chosen:
        p = get_profile(chosen)
        st.markdown(
            f'<div class="profile-desc">◈ {p["description"]}</div>',
            unsafe_allow_html=True
        )

# ── Scan button ───────────────────────────────────────────────
st.markdown('<hr>', unsafe_allow_html=True)

if st.session_state.file_bytes and st.session_state.profile_name:
    if not st.session_state.scan_done:
        if st.button("◈  Scan file"):
            with st.spinner("Scanning..."):
                try:
                    groups = fast_scan(
                        st.session_state.file_bytes,
                        st.session_state.profile_name,
                    )
                    cols = get_columns(
                        st.session_state.file_bytes,
                        st.session_state.profile_name,
                    )
                    st.session_state.question_groups = groups
                    st.session_state.columns         = cols
                    st.session_state.scan_done       = True
                    st.session_state.selected_cols   = [j for j,g,s in cols]
                    st.rerun()
                except Exception as e:
                    st.error(f"Scan failed: {e}")

# ═══════════════════════════════════════════════════════════
# STEP 3 — Select questions + columns
# ═══════════════════════════════════════════════════════════
if st.session_state.scan_done:
    groups = st.session_state.question_groups
    cols   = st.session_state.columns

    n_q  = len(groups)
    n_sh = sum(len(g['sheets']) for g in groups)
    st.markdown(
        f'<span class="stat-pill">{n_q} questions</span>'
        f'<span class="stat-pill">{n_sh} sheets</span>'
        f'<span class="stat-pill">{len(cols)} columns</span>',
        unsafe_allow_html=True
    )
    st.markdown("<br>", unsafe_allow_html=True)

    left, right = st.columns([1.8, 1])

    # ── Question selector ──────────────────────────────────
    with left:
        st.markdown('<div class="step-label">Step 3 — Select questions</div>', unsafe_allow_html=True)

        ctrl1, ctrl2, ctrl3 = st.columns([1,1,2])
        with ctrl1:
            if st.button("Select all"):
                st.session_state.selected_qs = {g['prefix'] for g in groups}
                st.rerun()
        with ctrl2:
            if st.button("Clear all"):
                st.session_state.selected_qs = set()
                st.rerun()
        with ctrl3:
            search = st.text_input("Search", placeholder="Filter questions...", label_visibility="collapsed")

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        filtered = [g for g in groups if not search or search.lower() in g['wording'].lower() or search.lower() in g['prefix'].lower()]

        for g in filtered:
            prefix  = g['prefix']
            wording = g['wording']
            n_s     = len(g['sheets'])
            checked = prefix in st.session_state.selected_qs

            new_val = st.checkbox(
                f"**{prefix}** — {wording[:70]}{'…' if len(wording)>70 else ''} *({n_s} sheet{'s' if n_s>1 else ''})*",
                value=checked,
                key=f"q_{prefix}",
            )
            if new_val != checked:
                if new_val:
                    st.session_state.selected_qs.add(prefix)
                else:
                    st.session_state.selected_qs.discard(prefix)

    # ── Column selector ────────────────────────────────────
    with right:
        st.markdown('<div class="step-label">Step 4 — Select columns</div>', unsafe_allow_html=True)
        st.markdown("<div style='font-size:0.8rem;color:#6B7280;margin-bottom:8px'>Choose which subgroups / countries to include</div>", unsafe_allow_html=True)

        selected_col_indices = []
        for j, g, s in cols:
            label   = g if not s or s.lower() == g.lower() else f"{g} — {s}"
            checked = j in st.session_state.selected_cols
            new_val = st.checkbox(label, value=checked, key=f"col_{j}")
            if new_val:
                selected_col_indices.append(j)
        st.session_state.selected_cols = selected_col_indices

    # ═══════════════════════════════════════════════════════
    # STEP 5 — Export
    # ═══════════════════════════════════════════════════════
    st.markdown('<hr>', unsafe_allow_html=True)
    st.markdown('<div class="step-label">Step 5 — Export</div>', unsafe_allow_html=True)

    n_selected = len(st.session_state.selected_qs)
    n_cols_sel = len(st.session_state.selected_cols)

    if n_selected == 0:
        st.info("Select at least one question to export.")
    elif n_cols_sel == 0:
        st.info("Select at least one column to export.")
    else:
        st.markdown(
            f'<span class="stat-pill">{n_selected} questions selected</span>'
            f'<span class="stat-pill">{n_cols_sel} columns selected</span>',
            unsafe_allow_html=True
        )
        st.markdown("<br>", unsafe_allow_html=True)

        if st.button(f"◈  Generate Excel ({n_selected} questions)"):
            selected_groups = [g for g in groups if g['prefix'] in st.session_state.selected_qs]
            col_indices     = st.session_state.selected_cols
            col_names       = [g for j,g,s in cols if j in col_indices]

            with st.spinner(f"Building {n_selected} tables..."):
                try:
                    excel_bytes = generate_excel(
                        selected_groups,
                        st.session_state.file_bytes,
                        st.session_state.profile_name,
                        col_indices,
                        col_names,
                    )
                    st.success(f"Done — {n_selected} questions exported")
                    st.download_button(
                        label="⬇  Download Excel",
                        data=excel_bytes,
                        file_name="crossstudio_export.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"Export failed: {e}")
                    import traceback
                    st.code(traceback.format_exc())
