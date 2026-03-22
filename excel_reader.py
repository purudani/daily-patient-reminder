"""
Read action report, actual report, and Mailchimp export; apply business rules;
output list of { action, record } for the daily job.

Data flow:
- Scheduler-style Action sheet: Date, Time (12h), Type, Location, Action (CREATE/RESCHEDULE/…),
  Reschedule Into (single text), Has newer.
- Reschedule Into is parsed for new date/time; falls back to legacy Reschedule Date/Time columns.
- Actions like "CANCEL w. remove" normalize to cancel.
- Location codes → full address via config.LOCATION_MAP.
- Mailchimp: PN + Email + optional First/Last (or Name).
"""
from __future__ import annotations

import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

from config import (
    ACTION_REPORT_PATH,
    ACTION_SHEET_NAME,
    ACTUAL_COL_DATE,
    ACTUAL_COL_LOCATION,
    ACTUAL_COL_TIME,
    ACTUAL_COL_TYPE,
    ACTUAL_REPORT_PATH,
    ACTUAL_SHEET_NAME,
    COL_ACTION,
    COL_APPT_DATE,
    COL_APPT_TIME,
    COL_APPT_TYPE,
    COL_HAS_NEWER_ACTION,
    COL_LOCATION,
    COL_MAILCHIMP_EMAIL,
    COL_MAILCHIMP_FIRST,
    COL_MAILCHIMP_LAST,
    COL_MAILCHIMP_NAME,
    COL_MAILCHIMP_PN,
    COL_PN,
    COL_RESCHEDULE_DATE,
    COL_RESCHEDULE_INTO,
    COL_RESCHEDULE_INTO_ALIASES,
    COL_RESCHEDULE_TIME,
    LOCATION_MAP,
    MAILCHIMP_EXPORT_PATH,
    MAILCHIMP_SHEET_NAME,
    MULTIPLE_ACTIONS_USE_LAST,
    SKIP_BLANK_PN,
    SKIP_FIRST_N_ROWS,
    SKIP_NEXT_DAY,
    SKIP_SAME_DAY,
)
from reschedule_parse import normalize_time_value, parse_reschedule_into

logger = logging.getLogger(__name__)

# Only these four actions are processed; any other value in the Action column is skipped.
ACTIONS_CREATE = ("create",)
ACTIONS_RESCHEDULE = ("reschedule",)
ACTIONS_CANCEL = ("cancel",)
ACTIONS_DELETE = ("delete",)
VALID_ACTIONS = (*ACTIONS_CREATE, *ACTIONS_RESCHEDULE, *ACTIONS_CANCEL, *ACTIONS_DELETE)

# Alternate headers for "Has newer" column in different exports
_HAS_NEWER_ALIASES = (
    "Has newer",
    "Has newer Date",
    "Has Newer Action",
    "Has newer actions?",
)


def _normalize_action(raw: str) -> str | None:
    """Map scheduler Action values to internal create|reschedule|cancel|delete."""
    s = raw.strip().lower()
    if s == "create":
        return "create"
    if s == "reschedule":
        return "reschedule"
    if s == "delete":
        return "delete"
    if s.startswith("cancel"):
        return "cancel"
    return None


def _actual_column(date_col: str | None, fallback: str) -> str:
    return date_col if date_col else fallback


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace from column names and lowercase for flexible matching."""
    df = df.rename(columns=lambda c: (c or "").strip())
    return df


def _cell_value(row: Any, col: str, df_columns: list) -> Any:
    """Get value from row by column name; try exact then case-insensitive."""
    if col in df_columns:
        return row[col]
    for c in df_columns:
        if c and str(c).strip().lower() == str(col).strip().lower():
            return row[c]
    return None


def _has_newer_value(row: Any, cols: list) -> Any:
    for alias in _HAS_NEWER_ALIASES:
        v = _cell_value(row, alias, cols)
        if v is not None and not (isinstance(v, float) and pd.isna(v)) and str(v).strip():
            return v
    return None


def _reschedule_into_value(row: Any, cols: list) -> Any:
    """Reschedule text from Action row — tries primary column name then aliases (e.g. Reschedule Info)."""
    for name in (COL_RESCHEDULE_INTO, *COL_RESCHEDULE_INTO_ALIASES):
        v = _cell_value(row, name, cols)
        if v is not None and not (isinstance(v, float) and pd.isna(v)) and str(v).strip():
            return v
    return None


def _parse_date(val: Any) -> str | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s:
        return None
    try:
        dt = pd.to_datetime(val)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return s


def _parse_time(val: Any) -> str:
    """Normalize to HH:MM:SS (handles 04:30p, Excel times, datetimes)."""
    return normalize_time_value(val)


def _normalize_pn(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if isinstance(val, float) and val != int(val):
            return str(val).strip()
        return str(int(val))
    return str(val).strip()


def load_action_df() -> pd.DataFrame:
    """Load action report; use row SKIP_FIRST_N_ROWS as header (so first N rows are skipped)."""
    path = Path(ACTION_REPORT_PATH).expanduser()
    if not path.exists():
        raise FileNotFoundError(f"Action report not found: {path}")
    df = pd.read_excel(path, sheet_name=ACTION_SHEET_NAME, header=SKIP_FIRST_N_ROWS)
    df = _normalize_columns(df)
    return df


def load_actual_df() -> pd.DataFrame:
    """Load actual appointment data (same file or separate)."""
    path = Path(ACTUAL_REPORT_PATH).expanduser()
    if not path.exists():
        logger.warning("Actual report not found: %s", path)
        return pd.DataFrame()
    df = pd.read_excel(path, sheet_name=ACTUAL_SHEET_NAME, header=SKIP_FIRST_N_ROWS)
    df = _normalize_columns(df)
    return df


def load_mailchimp_lookup() -> dict[str, dict[str, str]]:
    """Load Mailchimp / audience export: PN -> { email, name }."""
    path = Path(MAILCHIMP_EXPORT_PATH).expanduser()
    if not path.exists():
        logger.warning("Mailchimp export not found: %s", path)
        return {}
    df = pd.read_excel(path, sheet_name=MAILCHIMP_SHEET_NAME)
    df = _normalize_columns(df)
    cols = list(df.columns)
    lookup = {}
    for _, row in df.iterrows():
        pn_val = _cell_value(row, COL_MAILCHIMP_PN, cols)
        if pn_val is None or (isinstance(pn_val, float) and pd.isna(pn_val)):
            continue
        pn = _normalize_pn(pn_val)
        if not pn:
            continue
        email = _cell_value(row, COL_MAILCHIMP_EMAIL, cols)
        if email is None or (isinstance(email, float) and pd.isna(email)):
            continue
        email = str(email).strip()
        if not email or "@" not in email:
            continue
        name = _cell_value(row, COL_MAILCHIMP_NAME, cols)
        if name is not None and not (isinstance(name, float) and pd.isna(name)) and str(name).strip():
            display_name = str(name).strip()
        else:
            last = _cell_value(row, COL_MAILCHIMP_LAST, cols)
            first = _cell_value(row, COL_MAILCHIMP_FIRST, cols)
            parts = []
            if last is not None and not (isinstance(last, float) and pd.isna(last)):
                parts.append(str(last).strip())
            if first is not None and not (isinstance(first, float) and pd.isna(first)):
                parts.append(str(first).strip())
            display_name = " ".join(parts) if parts else email
        lookup[pn] = {"email": email, "name": display_name}
    return lookup


def _is_same_day(date_val: Any) -> bool:
    if date_val is None:
        return True
    try:
        dt = pd.to_datetime(date_val)
        return dt.date() == datetime.now().date()
    except Exception:
        return False


def _is_next_day(date_val: Any) -> bool:
    if date_val is None:
        return False
    try:
        dt = pd.to_datetime(date_val)
        return dt.date() == (datetime.now() + timedelta(days=1)).date()
    except Exception:
        return False


def _action_row_to_record(
    row: Any,
    cols: list,
    mailchimp: dict[str, dict[str, str]],
    actual_row: Any | None,
    use_reschedule_columns: bool,
    actual_cols: tuple[str, str, str, str],
    actual_df_columns: list | None,
) -> dict[str, Any] | None:
    """Build one record dict for calendar_actions from action row (and optional actual row)."""
    pn_val = _cell_value(row, COL_PN, cols)
    pn = _normalize_pn(pn_val)
    if not pn and SKIP_BLANK_PN:
        return None

    info = mailchimp.get(pn, {})
    if not info:
        for k, v in mailchimp.items():
            if _normalize_pn(k) == pn:
                info = v
                break
    email = info.get("email") or ""
    name = info.get("name") or ""

    ad_col, at_col, al_col, aty_col = actual_cols

    if actual_row is not None:
        acols = actual_df_columns if actual_df_columns is not None else cols
        date_str = _parse_date(_cell_value(actual_row, ad_col, acols))
        time_str = _parse_time(_cell_value(actual_row, at_col, acols))
        loc = _cell_value(actual_row, al_col, acols)
        appt_type = _cell_value(actual_row, aty_col, acols)
    else:
        base_date = _parse_date(_cell_value(row, COL_APPT_DATE, cols))
        base_time = _parse_time(_cell_value(row, COL_APPT_TIME, cols))
        if use_reschedule_columns:
            rinfo = _reschedule_into_value(row, cols)
            nd, nt = parse_reschedule_into(rinfo)
            date_str = nd or _parse_date(_cell_value(row, COL_RESCHEDULE_DATE, cols)) or base_date
            time_str = nt or _parse_time(_cell_value(row, COL_RESCHEDULE_TIME, cols)) or base_time
        else:
            date_str = base_date
            time_str = base_time
        loc = _cell_value(row, COL_LOCATION, cols)
        appt_type = _cell_value(row, COL_APPT_TYPE, cols)

    if not date_str:
        return None

    loc_code = (str(loc).strip().upper() if loc is not None and not (isinstance(loc, float) and pd.isna(loc)) else "LIB")
    location_address = LOCATION_MAP.get(loc_code) or loc_code

    return {
        "pn": pn,
        "email": email,
        "patient_name": name,
        "appt_date": date_str,
        "appt_time": time_str,
        "location": loc_code,
        "location_address": location_address,
        "appt_type": str(appt_type).strip().upper() if appt_type is not None else "",
    }


def _find_actual_row_for_pn(
    actual_df: pd.DataFrame,
    pn: str,
    *,
    date_hint: str | None = None,
    time_hint: str | None = None,
) -> Any | None:
    """
    Find one row in the Actual sheet for this PN.

    With multiple rows for the same PN, match **Date + Time** on Actual to the Action row’s
    scheduled **Date** / **Time**. If no match, use the first PN row and log a warning.
    """
    if actual_df is None or actual_df.empty:
        return None
    cols = list(actual_df.columns)
    ad_col = _actual_column(ACTUAL_COL_DATE, COL_APPT_DATE)
    at_col = _actual_column(ACTUAL_COL_TIME, COL_APPT_TIME)

    matching: list[Any] = []
    for _, row in actual_df.iterrows():
        pn_val = _cell_value(row, COL_PN, cols)
        if pn_val is None:
            continue
        if _normalize_pn(pn_val) != pn:
            continue
        matching.append(row)

    if not matching:
        return None
    if len(matching) == 1:
        return matching[0]

    # Multiple rows for this PN — match on scheduled date/time
    if date_hint and time_hint:
        for actual_row in matching:
            d = _parse_date(_cell_value(actual_row, ad_col, cols))
            t = _parse_time(_cell_value(actual_row, at_col, cols))
            if d == date_hint and t == time_hint:
                return actual_row

    logger.warning(
        "Actual sheet: %d rows for PN=%s; could not match Date+Time to Action row. "
        "Using the first row. Prefer one current-appointment row per patient on Actual, or unique Date+Time per visit.",
        len(matching),
        pn,
    )
    return matching[0]


def get_actions_to_process() -> list[dict[str, Any]]:
    """
    Load all data, apply rules, return list of:
    { "action": "Create"|"Reschedule"|"Cancel"|"Delete", "record": { ... } }

    Rules:
    - Skip first N rows of action sheet.
    - Skip blank PN (if SKIP_BLANK_PN).
    - Skip same-day (and optionally next-day) appointments from action row when not using Actual.
    - If multiple actions for same appointment, keep last.
    - Has Newer Action = Yes:
      - Create or Delete -> skip (won't show in Actual sheet).
      - Reschedule -> get new appt data from Actual; if not found in Actual -> skip (deleted);
        if new appt date is today -> skip; else use Actual data.
      - Cancel -> must find appt in Actual; if not found -> skip; else process cancel.
    - Reschedule without Actual -> parse Reschedule Into text; else legacy Reschedule Date/Time columns.
    - Location column codes are expanded to full address in record.location_address.
    """
    action_df = load_action_df()
    if action_df.empty:
        logger.info("Action sheet is empty")
        return []

    actual_df = load_actual_df()
    mailchimp = load_mailchimp_lookup()
    cols = list(action_df.columns)
    actual_cols_list = list(actual_df.columns) if not actual_df.empty else None
    actual_field_tuple = (
        _actual_column(ACTUAL_COL_DATE, COL_APPT_DATE),
        _actual_column(ACTUAL_COL_TIME, COL_APPT_TIME),
        _actual_column(ACTUAL_COL_LOCATION, COL_LOCATION),
        _actual_column(ACTUAL_COL_TYPE, COL_APPT_TYPE),
    )

    rows_with_action = []
    for idx, row in action_df.iterrows():
        action_val = _cell_value(row, COL_ACTION, cols)
        if action_val is None or (isinstance(action_val, float) and pd.isna(action_val)):
            continue
        action = _normalize_action(str(action_val).strip())
        if action is None or action not in VALID_ACTIONS:
            continue

        pn_val = _cell_value(row, COL_PN, cols)
        pn = _normalize_pn(pn_val)
        if not pn and SKIP_BLANK_PN:
            continue

        hn = _has_newer_value(row, cols)
        has_newer = hn is not None and str(hn).strip().lower() in ("true", "1", "yes", "y")

        if has_newer and action in (*ACTIONS_CREATE, *ACTIONS_DELETE):
            continue  # create & delete won't show in actual appt data sheet, hence skip

        # When Has newer action = Yes, we must resolve from Actual (except Create/Delete already skipped)
        need_actual_lookup = has_newer and action in (*ACTIONS_RESCHEDULE, *ACTIONS_CANCEL)
        actual_row = None
        if need_actual_lookup:
            date_hint = _parse_date(_cell_value(row, COL_APPT_DATE, cols))
            time_hint = _parse_time(_cell_value(row, COL_APPT_TIME, cols))
            actual_row = _find_actual_row_for_pn(
                actual_df,
                pn,
                date_hint=date_hint,
                time_hint=time_hint,
            )
            if actual_row is None:
                continue  # cannot find appt in actual sheet -> skip (e.g. deleted)

        use_actual_for_record = has_newer and action in ACTIONS_RESCHEDULE
        use_reschedule_cols = action in ACTIONS_RESCHEDULE and not use_actual_for_record

        # Same-day skip: for rows not using Actual, use action row date; for Actual we check after building record
        if not use_actual_for_record:
            date_val = _cell_value(row, COL_APPT_DATE, cols)
            if SKIP_SAME_DAY and _is_same_day(date_val):
                continue
            if SKIP_NEXT_DAY and _is_next_day(date_val):
                continue

        record = _action_row_to_record(
            row,
            cols,
            mailchimp,
            actual_row if use_actual_for_record else None,
            use_reschedule_cols,
            actual_field_tuple,
            actual_cols_list if use_actual_for_record else None,
        )
        if record is None:
            continue

        # When we used Actual data for Reschedule: skip if new appt date is today
        if use_actual_for_record and SKIP_SAME_DAY and _is_same_day(record.get("appt_date")):
            continue

        # For reschedule, keep original date/time so we can look up the existing event by key
        if action in ACTIONS_RESCHEDULE:
            record["original_appt_date"] = _parse_date(_cell_value(row, COL_APPT_DATE, cols))
            record["original_appt_time"] = _parse_time(_cell_value(row, COL_APPT_TIME, cols))
        if SKIP_BLANK_PN and not record.get("email"):
            logger.debug("Skipping PN %s: no email in Mailchimp", pn)
            continue

        rows_with_action.append({
            "action": action,
            "record": record,
            "sort_key": (pn, _cell_value(row, COL_APPT_DATE, cols), idx),
        })

    # If multiple actions for same appointment, keep last
    if MULTIPLE_ACTIONS_USE_LAST and rows_with_action:
        by_key = {}
        for item in rows_with_action:
            r = item["record"]
            a = item["action"]
            if a == "reschedule" and r.get("original_appt_date") and r.get("original_appt_time"):
                key = f"{r.get('pn')}_{r['original_appt_date']}_{r['original_appt_time']}"
            else:
                key = f"{r.get('pn')}_{r.get('appt_date')}_{r.get('appt_time')}"
            by_key[key] = item
        rows_with_action = list(by_key.values())

    # Sort for stable order
    rows_with_action.sort(key=lambda x: (str(x["sort_key"][0]), str(x["sort_key"][1]), x["sort_key"][2]))

    return [{"action": item["action"], "record": item["record"]} for item in rows_with_action]
