"""
Parse scheduler "Reschedule Into" column (single text field).

Examples from production exports:
  Time: 12:00p -> 10:00a
  Date: 2026-03-16 -> 2026-03-19 Time: 06:00p -> 02:30p
  Date: 3/16/2026 -> 3/19/2026 Time: 06:00p -> 02:30p
"""
from __future__ import annotations

import re
from typing import Any

import pandas as pd


def _time_12h_to_24h(s: str) -> str | None:
    """Convert '04:30p' / '11:30a' / '12:00P' to 'HH:MM:00'."""
    s = str(s).strip().upper().replace(" ", "")
    m = re.match(r"^(\d{1,2}):(\d{2})([AP])$", s)
    if not m:
        return None
    h, mi, ap = int(m.group(1)), int(m.group(2)), m.group(3)
    if ap == "P" and h != 12:
        h += 12
    elif ap == "A" and h == 12:
        h = 0
    return f"{h:02d}:{mi:02d}:00"


def normalize_time_value(val: Any) -> str:
    """Normalize Excel time to 'HH:MM:SS' (24h). Handles 12h suffix, datetime, Excel time."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "09:00:00"
    if hasattr(val, "hour") and hasattr(val, "minute") and not isinstance(val, str):
        try:
            return f"{val.hour:02d}:{val.minute:02d}:{getattr(val, 'second', 0):02d}"
        except Exception:
            pass
    s = str(val).strip()
    if not s:
        return "09:00:00"
    t24 = _time_12h_to_24h(s)
    if t24:
        return t24
    try:
        dt = pd.to_datetime(s)
        return dt.strftime("%H:%M:%S")
    except Exception:
        if ":" in s and len(s) >= 5:
            parts = s.split(":")
            if len(parts) >= 2:
                h, m = int(parts[0]), int(parts[1])
                sec = int(parts[2]) if len(parts) > 2 else 0
                return f"{h:02d}:{m:02d}:{sec:02d}"
        return "09:00:00"


def parse_reschedule_into(text: Any) -> tuple[str | None, str | None]:
    """
    Return (new_date_yyyy_mm_dd, new_time_hh_mm_ss) from Reschedule Into text.
    Either or both may be None if not present in the string.
    """
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return None, None
    s = str(text).strip()
    if not s:
        return None, None

    new_date = None
    new_time = None

    # Date: YYYY-MM-DD -> YYYY-MM-DD
    m = re.search(
        r"Date:\s*(\d{4}-\d{2}-\d{2})\s*->\s*(\d{4}-\d{2}-\d{2})",
        s,
        re.I,
    )
    if m:
        new_date = m.group(2)

    # Date: M/D/YYYY -> M/D/YYYY
    if new_date is None:
        m2 = re.search(
            r"Date:\s*(\d{1,2}/\d{1,2}/\d{4})\s*->\s*(\d{1,2}/\d{1,2}/\d{4})",
            s,
            re.I,
        )
        if m2:
            try:
                new_date = pd.to_datetime(m2.group(2)).strftime("%Y-%m-%d")
            except Exception:
                new_date = None

    # Time: 06:00p -> 02:30p (take right side as new)
    m3 = re.search(
        r"Time:\s*([\d:]+[apAP])\s*->\s*([\d:]+[apAP])",
        s,
    )
    if m3:
        t24 = _time_12h_to_24h(m3.group(2))
        if t24:
            new_time = t24

    return new_date, new_time
