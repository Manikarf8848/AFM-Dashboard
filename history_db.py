"""
history_db.py  —  in-memory upload history (Streamlit Cloud safe)
No SQLite / filesystem dependency — resets on each server restart.
"""

import pandas as pd

_HISTORY = []


def record_upload(file_name: str, df: pd.DataFrame):
    weeks = []
    if "Week" in df.columns:
        weeks = sorted([int(w) for w in df["Week"].dropna().unique().tolist()])

    date_min = date_max = upload_ts = ""
    if "Time Created" in df.columns:
        tc = pd.to_datetime(df["Time Created"], errors="coerce").dropna()
        if not tc.empty:
            date_min  = str(tc.min().date())
            date_max  = str(tc.max().date())
            upload_ts = str(tc.max())

    _HISTORY.insert(0, {
        "file_name":    file_name,
        "upload_ts":    upload_ts,
        "total_andons": len(df),
        "week_numbers": weeks,
        "date_min":     date_min,
        "date_max":     date_max,
    })


def get_history(n: int = 20):
    return _HISTORY[:n]


def clear_history():
    _HISTORY.clear()
