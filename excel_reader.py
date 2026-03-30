"""
Read action report, actual report, and Mailchimp export; apply business rules;
output list of { action, record } for the daily job.

Data flow:
- Scheduler-style Action sheet: Date, Time (12h), Type, Location, Action (CREATE/RESCHEDULE/…),
  Reschedule Into (single text), Has newer.
- Has newer = Yes: match one row in Actual by PN + Action Date/Time on Actual;
  if several rows match, narrow by appointment Date then appointment Time.
  The outgoing slot uses the matched Actual row; action still comes from Action. See BUSINESS_LOGIC.md.
- Reschedule Into is parsed for new date/time when not using that has_newer path; falls back to legacy columns.
- Actions like "CANCEL w. remove" normalize to cancel.
- Location codes → full address via config.LOCATION_MAP.
- Mailchimp: PN + Email + optional First/Last (or Name).
"""
from __future__ import annotations

import logging
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

from config import (
    ACTION_REPORT_PATH,
    ACTUAL_COL_ACTION_DATE,
    ACTUAL_COL_ACTION_TIME,
    ACTUAL_COL_DATE,
    ACTUAL_COL_LOCATION,
    ACTUAL_COL_TIME,
    ACTUAL_COL_TYPE,
    ACTUAL_REPORT_PATH,
    COL_ACTION,
    COL_ACTION_DATE,
    COL_ACTION_TIME,
    COL_APPT_DATE,
    COL_APPT_TIME,
    COL_APPT_TYPE,
    COL_HAS_NEWER_ACTION,
    COL_LOCATION,
    COL_PATIENT_NAME,
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
    REFERENCE_DATE,
    SKIP_BLANK_PN,
    SKIP_FIRST_N_ROWS,
    SKIP_NEXT_DAY,
    SKIP_SAME_DAY,
)
from reschedule_parse import normalize_time_value, parse_reschedule_into

logger = logging.getLogger(__name__)


_ACTIVITY_BETWEEN_RE = re.compile(
    r"Activity\s+between:\s*([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{2,4})",
    re.IGNORECASE,
)


def _parse_header_date_token(token: str) -> date | None:
    raw = (token or "").strip()
    if not raw:
        return None
    for fmt in ("%m-%d-%y", "%m/%d/%y", "%m-%d-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    try:
        return pd.to_datetime(raw).date()
    except Exception:
        return None


def _report_activity_date_from_action_header() -> date | None:
    """
    Parse "Activity between: <date> and <date>" from the first rows of Action sheet.
    Scheduler exports place this text before the table header.
    """
    path = Path(ACTION_REPORT_PATH).expanduser()
    if not path.exists():
        return None
    try:
        preface = pd.read_excel(path, sheet_name=0, header=None, nrows=3)
    except Exception:
        return None
    for _, row in preface.iterrows():
        for v in row.tolist():
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            m = _ACTIVITY_BETWEEN_RE.search(str(v))
            if not m:
                continue
            d = _parse_header_date_token(m.group(1))
            if d:
                return d
    return None


def _reference_today() -> date:
    """
    Date used by skip logic.
    - If config.REFERENCE_DATE is set (YYYY-MM-DD), use that.
    - Else use the machine's current local date.
    """
    raw = (REFERENCE_DATE or "").strip()
    if raw:
        try:
            return datetime.strptime(raw, "%Y-%m-%d").date()
        except ValueError:
            logger.warning("Invalid REFERENCE_DATE=%s; expected YYYY-MM-DD. Falling back to system date.", raw)
    inferred = _report_activity_date_from_action_header()
    if inferred is not None:
        return inferred
    return datetime.now().date()

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


def _is_uncpt(appt_type: Any) -> bool:
    if appt_type is None or (isinstance(appt_type, float) and pd.isna(appt_type)):
        return False
    return str(appt_type).strip().upper() == "UNCPT"


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


def _parse_time_optional(val: Any) -> str | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if not s:
        return None
    return _parse_time(val)


def _normalize_pn(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if isinstance(val, float) and val != int(val):
            return str(val).strip()
        return str(int(val))
    return str(val).strip()


def load_action_df() -> pd.DataFrame:
    """Load action report from the first sheet; header starts at SKIP_FIRST_N_ROWS."""
    path = Path(ACTION_REPORT_PATH).expanduser()
    if not path.exists():
        raise FileNotFoundError(f"Action report not found: {path}")
    # Scheduler export tab names change daily; always consume the first tab.
    df = pd.read_excel(path, sheet_name=0, header=SKIP_FIRST_N_ROWS)
    df = _normalize_columns(df)
    return df


def load_actual_df() -> pd.DataFrame:
    """Load actual appointment data from the first sheet (tab name may vary daily)."""
    path = Path(ACTUAL_REPORT_PATH).expanduser()
    if not path.exists():
        logger.warning("Actual report not found: %s", path)
        return pd.DataFrame()
    df = pd.read_excel(path, sheet_name=0, header=SKIP_FIRST_N_ROWS)
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
        return dt.date() == _reference_today()
    except Exception:
        return False


def _is_next_day(date_val: Any) -> bool:
    if date_val is None:
        return False
    try:
        dt = pd.to_datetime(date_val)
        return dt.date() == (_reference_today() + timedelta(days=1))
    except Exception:
        return False


def _is_future_beyond_next_day(date_val: Any) -> bool:
    if date_val is None:
        return False
    try:
        dt = pd.to_datetime(date_val)
        return dt.date() > (_reference_today() + timedelta(days=1))
    except Exception:
        return False


def _appointment_group_key(action: str, record: dict[str, Any]) -> str:
    """
    Group rows that refer to the same underlying appointment slot.
    Reschedules key off the original slot so updates replace the prior invite.
    """
    if action == "reschedule" and record.get("original_appt_date") and record.get("original_appt_time"):
        return f"{record.get('pn')}_{record['original_appt_date']}_{record['original_appt_time']}"
    return f"{record.get('pn')}_{record.get('appt_date')}_{record.get('appt_time')}"


def _action_sort_tuple(item: dict[str, Any]) -> tuple[str, str, str, str, int]:
    record = item["record"]
    return (
        str(record.get("pn") or ""),
        str(record.get("appt_date") or ""),
        str(record.get("appt_time") or ""),
        str(item.get("action_time") or ""),
        int(item.get("row_index") or 0),
    )


def _resolve_multiple_actions(rows_with_action: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Sort actions by PN, appointment date/time, then action time.
    If a same-appointment group contains only create/delete in either order, skip it entirely.
    Otherwise keep the final action after sorting.
    """
    if not rows_with_action:
        return []

    grouped: dict[str, list[dict[str, Any]]] = {}
    for item in rows_with_action:
        grouped.setdefault(item["appointment_group_key"], []).append(item)

    resolved: list[dict[str, Any]] = []
    for key, group in grouped.items():
        group.sort(key=_action_sort_tuple)
        unique_actions = {g["action"] for g in group}
        if len(group) == 2 and unique_actions == {"create", "delete"}:
            logger.info("Skipping appointment %s: same-day actions collapse to create/delete pair.", key)
            continue
        resolved.append(group[-1])

    resolved.sort(key=_action_sort_tuple)
    return resolved


def _action_row_to_record(
    row: Any,
    cols: list,
    mailchimp: dict[str, dict[str, str]],
    actual_row: Any | None,
    use_reschedule_columns: bool,
    actual_cols: tuple[str, str, str, str],
    actual_df_columns: list | None,
) -> dict[str, Any] | None:
    """Build one record dict for calendar_actions from action row and optional Actual row."""
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
        row_name = _cell_value(actual_row, COL_PATIENT_NAME, acols)
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
        row_name = _cell_value(row, COL_PATIENT_NAME, cols)

    if not date_str:
        return None

    loc_code = (str(loc).strip().upper() if loc is not None and not (isinstance(loc, float) and pd.isna(loc)) else "LIB")
    location_address = LOCATION_MAP.get(loc_code) or loc_code
    if not name:
        if row_name is not None and not (isinstance(row_name, float) and pd.isna(row_name)):
            name = str(row_name).strip()

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


def _find_actual_row_for_has_newer(
    actual_df: pd.DataFrame,
    pn: str,
    action_row: Any,
    action_cols: list,
) -> Any | None:
    """
    Find one Actual row for has_newer processing.

    1. Candidates: same PN, and Actual's Action Date/Time match the Action row's
       Action Date/Time (scheduler columns on the Actual export).
    2. Exactly one match → return it.
    3. Several matches → keep rows whose Actual *appointment* Date equals the Action
       row's appointment Date; if still several, narrow by Actual appointment Time vs
       Action appointment Time.
    4. Zero or ambiguous after narrowing → None (caller skips).
    """
    if actual_df is None or actual_df.empty:
        return None
    cols = list(actual_df.columns)
    ad_col = _actual_column(ACTUAL_COL_DATE, COL_APPT_DATE)
    at_col = _actual_column(ACTUAL_COL_TIME, COL_APPT_TIME)
    aad_col = _actual_column(ACTUAL_COL_ACTION_DATE, COL_ACTION_DATE)
    aat_col = _actual_column(ACTUAL_COL_ACTION_TIME, COL_ACTION_TIME)

    action_date_hint = _parse_date(_cell_value(action_row, COL_ACTION_DATE, action_cols))
    action_time_hint = _parse_time_optional(_cell_value(action_row, COL_ACTION_TIME, action_cols))
    orig_date = _parse_date(_cell_value(action_row, COL_APPT_DATE, action_cols))
    orig_time = _parse_time_optional(_cell_value(action_row, COL_APPT_TIME, action_cols))

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

    if action_date_hint and action_time_hint:
        matching = [
            r for r in matching
            if _parse_date(_cell_value(r, aad_col, cols)) == action_date_hint
            and _parse_time_optional(_cell_value(r, aat_col, cols)) == action_time_hint
        ]
    else:
        matching = []

    if not matching:
        return None
    if len(matching) == 1:
        return matching[0]

    if orig_date:
        by_date = [
            r for r in matching
            if _parse_date(_cell_value(r, ad_col, cols)) == orig_date
        ]
        if len(by_date) == 1:
            return by_date[0]
        if len(by_date) == 0:
            logger.warning(
                "Actual sheet: %d rows for PN=%s match Action Date/Time but none match Action appt date %s; skip.",
                len(matching),
                pn,
                orig_date,
            )
            return None
        matching = by_date

    if len(matching) == 1:
        return matching[0]

    if orig_time:
        by_time = [
            r for r in matching
            if _parse_time_optional(_cell_value(r, at_col, cols)) == orig_time
        ]
        if len(by_time) == 1:
            return by_time[0]
        if len(by_time) == 0:
            logger.warning(
                "Actual sheet: %d rows for PN=%s after appt-date narrow but none match Action appt time; skip.",
                len(matching),
                pn,
            )
            return None
        matching = by_time

    if len(matching) == 1:
        return matching[0]

    logger.warning(
        "Actual sheet: %d rows for PN=%s remain ambiguous after appt date/time narrow; skip.",
        len(matching),
        pn,
    )
    return None


def get_actions_to_process() -> list[dict[str, Any]]:
    """
    Load all data, apply rules, return list of:
    { "action": "Create"|"Reschedule"|"Cancel"|"Delete", "record": { ... } }

    Rules:
    - Skip first N rows of action sheet.
    - Skip blank PN (if SKIP_BLANK_PN).
    - If multiple actions for same appointment, keep last.
    - Without Has Newer:
      - Skip same-day/next-day using Action appointment Date.
      - Skip UNCPT using Action/record type.
      - Reschedule parses Reschedule Into first, then legacy Reschedule Date/Time.
    - With Has Newer:
      - Match Actual by PN + Action Date + Action Time on Actual.
      - If several rows match, narrow by appointment Date, then appointment Time.
      - Skip only if no usable Actual row remains, resolved type is UNCPT, or resolved
        appointment date is same-day/next-day per flags.
      - Outgoing appointment date/time/location/type come from Actual; action stays from Action.
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

        original_appt_date = _parse_date(_cell_value(row, COL_APPT_DATE, cols))

        if not has_newer:
            action_appt_type = _cell_value(row, COL_APPT_TYPE, cols)
            if _is_uncpt(action_appt_type):
                continue
            if SKIP_SAME_DAY and _is_same_day(original_appt_date):
                continue

        need_actual_lookup = has_newer
        actual_row = None
        if need_actual_lookup:
            actual_row = _find_actual_row_for_has_newer(actual_df, pn, row, cols)
            if actual_row is None:
                continue

        use_reschedule_cols = action in ACTIONS_RESCHEDULE and not need_actual_lookup

        record = _action_row_to_record(
            row,
            cols,
            mailchimp,
            actual_row if need_actual_lookup else None,
            use_reschedule_cols,
            actual_field_tuple,
            actual_cols_list if need_actual_lookup else None,
        )
        if record is None:
            continue

        moved_future_to_next_day = (
            action in ACTIONS_RESCHEDULE
            and _is_future_beyond_next_day(original_appt_date)
            and _is_next_day(record.get("appt_date"))
        )

        if _is_uncpt(record.get("appt_type")):
            continue
        if SKIP_SAME_DAY and _is_same_day(record.get("appt_date")):
            continue
        if SKIP_NEXT_DAY and _is_next_day(record.get("appt_date")) and not moved_future_to_next_day:
            continue

        # For reschedule, keep original date/time so we can look up the existing event by key
        if action in ACTIONS_RESCHEDULE:
            record["original_appt_date"] = original_appt_date
            record["original_appt_time"] = _parse_time(_cell_value(row, COL_APPT_TIME, cols))
        elif need_actual_lookup and action in (*ACTIONS_CANCEL, *ACTIONS_DELETE):
            record["original_appt_date"] = original_appt_date
            record["original_appt_time"] = _parse_time(_cell_value(row, COL_APPT_TIME, cols))
        if SKIP_BLANK_PN and not record.get("email"):
            logger.debug("Skipping PN %s: no email in Mailchimp", pn)
            continue

        rows_with_action.append({
            "action": action,
            "record": record,
            "row_index": idx,
            "action_time": _parse_time_optional(_cell_value(row, COL_ACTION_TIME, cols)) or "",
            "appointment_group_key": _appointment_group_key(action, record),
        })

    # If multiple actions for same appointment, keep last
    if MULTIPLE_ACTIONS_USE_LAST and rows_with_action:
        rows_with_action = _resolve_multiple_actions(rows_with_action)

    # Sort for stable order
    rows_with_action.sort(key=_action_sort_tuple)

    return [{"action": item["action"], "record": item["record"]} for item in rows_with_action]
