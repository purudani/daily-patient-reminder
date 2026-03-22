"""
Persistent store for appointment → ICS identity (UID + SEQUENCE).

Reschedule/cancel reuse the same UID with a higher SEQUENCE (RFC 5545).
Legacy rows may include event_id / i_cal_uid from older Graph-calendar flows;
get_invite_state still reads i_cal_uid as a fallback UID for cancel.
"""
from __future__ import annotations

import json
import os
import threading
from pathlib import Path
from typing import Any


def _default_path() -> str:
    try:
        from config import EVENT_ID_STORE_PATH
        return EVENT_ID_STORE_PATH
    except Exception:
        return os.path.join(os.path.dirname(__file__), "event_id_store.json")


_path: str | None = None
_lock = threading.Lock()


def get_store_path() -> str:
    global _path
    if _path is None:
        _path = _default_path()
    return _path


def _load_unsafe() -> dict[str, Any]:
    path = get_store_path()
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except json.JSONDecodeError:
        # Recover from truncated/empty file instead of crashing the run.
        return {}


def _save_unsafe(data: dict[str, Any]) -> None:
    path = get_store_path()
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def get_event_id(appointment_key: str) -> str | None:
    """Legacy: stored Graph event ID (unused for ICS-only flow)."""
    with _lock:
        data = _load_unsafe()
    record = data.get(appointment_key)
    if not record or not isinstance(record, dict):
        return None
    eid = record.get("event_id")
    return str(eid).strip() if eid else None


def get_i_cal_uid(appointment_key: str) -> str | None:
    """Legacy: Graph iCalUId (may be used as UID when migrating from Graph calendar)."""
    with _lock:
        data = _load_unsafe()
    record = data.get(appointment_key)
    if not record or not isinstance(record, dict):
        return None
    uid = record.get("i_cal_uid")
    return str(uid).strip() if uid else None


def get_invite_state(appointment_key: str) -> dict[str, Any] | None:
    """
    Return { "ical_uid": str, "sequence": int } if we have a UID to update/cancel, else None.

    Prefers ical_uid; falls back to legacy i_cal_uid from Graph so a cancel after
    switching to ICS can still reference the old invite when that UID was stored.
    """
    with _lock:
        data = _load_unsafe()
    record = data.get(appointment_key)
    if not record or not isinstance(record, dict):
        return None
    uid = (record.get("ical_uid") or record.get("i_cal_uid") or "").strip()
    if not uid:
        return None
    seq = record.get("sequence", 0)
    if isinstance(seq, str) and seq.strip().isdigit():
        seq = int(seq.strip())
    elif not isinstance(seq, int):
        seq = 0
    return {"ical_uid": uid, "sequence": seq}


def set_invite_state(appointment_key: str, ical_uid: str, sequence: int) -> None:
    """Persist UID and SEQUENCE after sending an ICS invite or update."""
    with _lock:
        data = _load_unsafe()
        data[appointment_key] = {
            "ical_uid": ical_uid,
            "sequence": int(sequence),
        }
        _save_unsafe(data)


def set_event(appointment_key: str, event_id: str, i_cal_uid: str | None = None) -> None:
    """
    Legacy Graph helper (kept for compatibility). Prefer set_invite_state for ICS.

    If i_cal_uid is set, mirrors it into ical_uid and preserves sequence when possible.
    """
    with _lock:
        data = _load_unsafe()
        uid = (i_cal_uid or "").strip()
        rec: dict[str, Any] = {
            "event_id": event_id,
            "i_cal_uid": uid,
        }
        if uid:
            prev = data.get(appointment_key)
            seq = 0
            if isinstance(prev, dict):
                ps = prev.get("sequence", 0)
                if isinstance(ps, str) and ps.strip().isdigit():
                    seq = int(ps.strip())
                elif isinstance(ps, int):
                    seq = ps
            rec["ical_uid"] = uid
            rec["sequence"] = seq
        data[appointment_key] = rec
        _save_unsafe(data)


def remove_event(appointment_key: str) -> None:
    """Remove stored event ID after deleting/cancelling the event."""
    with _lock:
        data = _load_unsafe()
        if appointment_key in data:
            del data[appointment_key]
            _save_unsafe(data)


def get_all_keys() -> list[str]:
    """Return all stored appointment keys (for debugging)."""
    with _lock:
        data = _load_unsafe()
    return list(data.keys())
