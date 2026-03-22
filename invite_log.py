"""
Append one row per outbound patient email to an Excel log (audit trail).

Thread-safe append; creates the file with headers on first write.
"""
from __future__ import annotations

import logging
import threading
from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import pandas as pd

logger = logging.getLogger(__name__)

_lock = threading.Lock()

LOG_COLUMNS: list[str] = [
    "sent_at_eastern",
    "from_email",
    "to_email",
    "to_name",
    "action",
    "subject",
    "pn",
    "appointment_id",
    "appointment_key",
    "appt_date",
    "appt_time",
    "duration_minutes",
    "location_code",
    "location_address",
    "ics_uid",
    "ics_sequence",
    "ics_method",
    "ics_attachment",
]


def _log_path() -> Path | None:
    from config import INVITE_LOG_PATH

    raw = (INVITE_LOG_PATH or "").strip()
    if not raw:
        return None
    return Path(raw).expanduser()


def append_invite_row(
    *,
    from_email: str,
    to_email: str,
    to_name: str | None,
    action: str,
    subject: str,
    record: dict[str, Any],
    appointment_key: str,
    ical_uid: str,
    ics_sequence: int,
    ics_method: str,
    ics_attachment: str,
    duration_minutes: int | None = None,
) -> None:
    """Append a single row to the invite log Excel file."""
    path = _log_path()
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)

    from config import TIMEZONE

    tz = ZoneInfo((TIMEZONE or "America/New_York").strip() or "America/New_York")
    sent_at = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S %Z")

    row = {
        "sent_at_eastern": sent_at,
        "from_email": (from_email or "").strip(),
        "to_email": (to_email or "").strip(),
        "to_name": (to_name or "").strip() if to_name else "",
        "action": (action or "").strip().lower(),
        "subject": (subject or "")[:500],
        "pn": str(record.get("pn") or record.get("PN") or ""),
        "appointment_id": str(record.get("appointment_id") or record.get("AppointmentID") or ""),
        "appointment_key": str(appointment_key),
        "appt_date": str(record.get("appt_date") or record.get("Date") or ""),
        "appt_time": str(record.get("appt_time") or record.get("Time") or ""),
        "duration_minutes": duration_minutes if duration_minutes is not None else "",
        "location_code": str(record.get("location") or record.get("Location") or ""),
        "location_address": str(record.get("location_address") or ""),
        "ics_uid": str(ical_uid),
        "ics_sequence": int(ics_sequence),
        "ics_method": str(ics_method).upper(),
        "ics_attachment": str(ics_attachment),
    }

    with _lock:
        if path.is_file():
            try:
                existing = pd.read_excel(path, sheet_name=0, engine="openpyxl")
                if "sent_at_utc" in existing.columns and "sent_at_eastern" not in existing.columns:
                    existing = existing.rename(columns={"sent_at_utc": "sent_at_eastern"})
            except Exception as e:
                logger.warning("Could not read invite log %s (%s); recreating.", path, e)
                existing = pd.DataFrame(columns=LOG_COLUMNS)
        else:
            existing = pd.DataFrame(columns=LOG_COLUMNS)

        combined = pd.concat(
            [existing, pd.DataFrame([row])],
            ignore_index=True,
        )
        # Ensure column order / any new columns
        for c in LOG_COLUMNS:
            if c not in combined.columns:
                combined[c] = ""
        combined = combined[[c for c in LOG_COLUMNS if c in combined.columns]]
        combined.to_excel(path, sheet_name="Invites", index=False, engine="openpyxl")

    logger.debug("Invite log append: %s -> %s (%s)", from_email, to_email, action)
