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
        "description": "KP topline banner — question at row 2, subgroups at row 1",
        "specs": [
            ("Question row",   "Row 2 (multi-line support)"),
            ("Column headers", "Row 1 (finds Total dynamically)"),
            ("Base row",       "Row 7"),
            ("Data start",     "Dynamic (after base row)"),
            ("Data step",      "Every 3 rows"),
            ("Value row",      "Row offset +1"),
            ("Stop on",        "Total Mentions / Back to Top / Sigma"),
        ],
        "question_row":         2,
        "header_row":           1,
        "base_row":             7,
        "data_start":           None,   # dynamic
        "data_step":            3,
        "value_row_offset":     1,
        "skip_sheet_0":         True,
        "column_start":         "find_total",   # find 'total' in header row
        "stop_on":              ["total mentions", "total mentions", "back to top", "sigma",
                                 "overlap formula used"],
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

def get_profile_names():
    return list(PROFILES.keys())

def get_profile(name):
    return PROFILES.get(name)
