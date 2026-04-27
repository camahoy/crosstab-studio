"""
profiles.py — Crosstab Studio saved format profiles
Each profile defines how to read a specific banner/crosstab format.
"""

PROFILES = {
    "GQR Corporate Reputation": {
        "description": "GQR weighted banner — question at row 3, categories at row 5, data every 3 rows",
        "fmt_codes": ["fmt5", "fmt6"],
        "question_row": 3,
        "header_row": 5,
        "base_row": 8,
        "data_start": 12,
        "data_step": 3,
        "value_row_offset": 1,
        "sig_row_offset": 2,
        "skip_sheet_0": True,
        "sheet_classifier": "gqr",
        "column_start": 1,
        "stop_on": ["sigma"],
    },
    "Mastercard Multi-Country": {
        "description": "Multi-country banner — question at row 2, countries at row 3, data every 3 rows",
        "fmt_codes": ["fmt2", "fmt3"],
        "question_row": 2,
        "header_row": 3,
        "sublabel_row": 4,
        "base_row": 7,
        "data_start_base_offset": 2,
        "data_step": 3,
        "value_row_offset": 1,
        "sig_row_offset": 2,
        "skip_sheet_0": True,
        "sheet_classifier": "mastercard",
        "column_start": "auto",  # auto-detect col0 or col1 for Total
        "stop_on": ["sigma"],
        "coerce_strings": True,  # values stored as strings
    },
}

def get_profile_names():
    return list(PROFILES.keys())

def get_profile(name):
    return PROFILES.get(name)
