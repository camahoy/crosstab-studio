"""
profiles.py — Crosstab Studio format profiles
"""

PROFILES = {
    "Corporate Reputation": {
        "description": "GQR weighted banner — question at row 3, categories at row 5",
        "specs": [
            ("Question row",   "Row 3"),
            ("Column headers", "Row 5"),
            ("Base row",       "Row 8"),
            ("Data start",     "Row 12"),
            ("Data step",      "Every 3 rows"),
            ("Value row",      "Row offset +1"),
            ("Stop on",        "Sigma"),
        ],
        "question_row":         3,
        "header_row":           5,
        "base_row":             8,
        "data_start":           12,
        "data_step":            3,
        "value_row_offset":     1,
        "skip_sheet_0":         True,
        "column_start":         1,
        "stop_on":              ["sigma"],
        "coerce_strings":       False,
        "multi_file_mode":      "waves",
    },
    "Global Brand Identity": {
        "description": "Multi-country banner — question at row 2, countries at row 3",
        "specs": [
            ("Question row",   "Row 2"),
            ("Column headers", "Row 3 (auto-detect col 0 or 1)"),
            ("Sub-labels",     "Row 4"),
            ("Base row",       "Row 7"),
            ("Data start",     "Base row + 2"),
            ("Data step",      "Every 3 rows"),
            ("Value row",      "Row offset +1"),
            ("Stop on",        "Sigma"),
        ],
        "question_row":              2,
        "header_row":                3,
        "sublabel_row":              4,
        "base_row":                  7,
        "data_start_base_offset":    2,
        "data_step":                 3,
        "value_row_offset":          1,
        "skip_sheet_0":              True,
        "column_start":              "auto",
        "stop_on":                   ["sigma"],
        "coerce_strings":            True,
        "multi_file_mode":           "subgroups",
    },
    "KP": {
        "description": "KP topline banner — question at row 2, subgroups at row 4",
        "specs": [
            ("Question row",   "Row 2 (multi-line support)"),
            ("Column headers", "Row 4 (Total, subgroup names)"),
            ("Base row",       "Row 7"),
            ("Data start",     "Dynamic (after base row)"),
            ("Data step",      "Every 3 rows"),
            ("Value row",      "Row offset +1"),
            ("Stop on",        "Total Mentions / Back to Top / Sigma"),
        ],
        "question_row":         2,
        "header_row":           4,
        "base_row":             7,
        "data_start":           None,
        "data_step":            3,
        "value_row_offset":     1,
        "skip_sheet_0":         True,
        "column_start":         1,
        "stop_on":              ["total mentions", "back to top", "sigma",
                                 "overlap formula used"],
        "coerce_strings":       False,
        "multi_file_mode":      "waves",
    },
    "KP Banner 2": {
        "description": "KP Banner 2 — question at row 4, group/subgroup headers at rows 8–9",
        "specs": [
            ("Detection",      "Row 0 col 1 = 'Back to Contents'"),
            ("Question row",   "Row 4"),
            ("Group headers",  "Row 8 (Total at col 1, group names from col 2)"),
            ("Column headers", "Row 9 (subgroup names from col 2)"),
            ("Base row",       "Row 13 (weighted base counts)"),
            ("Data start",     "Row 15"),
            ("Data step",      "Every 3 rows"),
            ("Value row",      "Row offset +1 (percentages, already ×100)"),
            ("Stop on",        "Total Mentions / Back to Top / Sigma"),
        ],
        "question_row":         4,
        "header_row":           9,
        "base_row":             13,
        "data_start":           15,
        "data_step":            3,
        "value_row_offset":     1,
        "skip_sheet_0":         True,
        "column_start":         1,
        "stop_on":              ["total mentions", "back to top", "sigma",
                                 "overlap formula used"],
        "coerce_strings":       True,
        "multi_file_mode":      "waves",
        "kp_banner2":           True,   # flag: special column detection
        "values_already_pct":   True,   # values stored as e.g. 19.0, not 0.19
    },
    "KP Omni": {
        "description": "KP Omni banner — question at row 4, dynamic base/header detection",
        "specs": [
            ("Question row",   "Row 4"),
            ("Statement row",  "Row 5 (if present, shifts data down by 1)"),
            ("Column headers", "Dynamic (2 rows above Base Weighted)"),
            ("Base row",       "Dynamic (finds 'Base Weighted' row)"),
            ("Data start",     "Dynamic (Base Weighted row + 1)"),
            ("Data step",      "Every 3 rows"),
            ("Value row",      "Row offset +1"),
            ("Stop on",        "Back to Top / Field Dates / Sigma"),
        ],
        "question_row":         4,
        "header_row":           None,    # dynamic
        "base_row":             None,    # dynamic — finds 'Base Weighted'
        "data_start":           None,    # dynamic
        "data_step":            3,
        "value_row_offset":     1,
        "skip_sheet_0":         True,
        "column_start":         1,
        "stop_on":              ["back to top", "field dates", "sigma",
                                 "overlap formula used"],
        "coerce_strings":       False,
        "multi_file_mode":      "waves",
        "dynamic_base":         True,    # flag: use dynamic row detection
    },
    "IData": {
        "description": "IData banner — question at row 2, subgroups at row 4",
        "specs": [
            ("Question row",   "Row 2"),
            ("Column headers", "Row 4 (Total at col 1)"),
            ("Base row",       "Row 7"),
            ("Data start",     "Row 9"),
            ("Data step",      "Every 3 rows"),
            ("Value row",      "Row offset +1"),
            ("Stop on",        "Sigma / Overlap formula used"),
        ],
        "question_row":         2,
        "header_row":           4,
        "base_row":             7,
        "data_start":           9,
        "data_step":            3,
        "value_row_offset":     1,
        "skip_sheet_0":         True,
        "column_start":         1,
        "stop_on":              ["sigma", "overlap formula used"],
        "coerce_strings":       False,
        "multi_file_mode":      "waves",
    },
    "+ Add new format": {
        "description": "coming_soon",
        "specs": [],
        "question_row": 0, "header_row": 0, "base_row": 0,
        "data_start": 0,   "data_step": 3,  "value_row_offset": 1,
        "skip_sheet_0": False, "column_start": 1,
        "stop_on": [], "coerce_strings": False, "multi_file_mode": "waves",
    },
}

def _load_user():
    import json, os
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'user_profiles.json')
    try:
        if os.path.exists(path):
            with open(path, 'r') as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def get_profile_names():
    names = list(PROFILES.keys())
    user  = _load_user()
    # Insert user profiles just before "+ Add new format"
    insert_at = len(names) - 1
    for uname in user:
        if uname not in names:
            names.insert(insert_at, uname)
            insert_at += 1
    return names


def get_profile(name):
    if name in PROFILES:
        return PROFILES[name]
    return _load_user().get(name)
