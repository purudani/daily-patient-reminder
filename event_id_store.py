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
    Return invite metadata if we have a UID to update/cancel, else None.

    Includes optional last_appt_date / last_appt_time / last_duration_minutes so a
    cancellation uses the same DTSTART/DTEND as the last REQUEST (critical after reschedules).

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
    out: dict[str, Any] = {"ical_uid": uid, "sequence": seq}
    ld = record.get("last_appt_date")
    lt = record.get("last_appt_time")
    if ld and lt:
        out["last_appt_date"] = str(ld).strip()
        out["last_appt_time"] = str(lt).strip()
    dur = record.get("last_duration_minutes")
    if dur is not None:
        try:
            out["last_duration_minutes"] = int(dur)
        except (TypeError, ValueError):
            pass
    return out


def set_invite_state(
    appointment_key: str,
    ical_uid: str,
    sequence: int,
    *,
    last_appt_date: str | None = None,
    last_appt_time: str | None = None,
    last_duration_minutes: int | None = None,
) -> None:
    """Persist UID and SEQUENCE after sending an ICS invite or update."""
    with _lock:
        data = _load_unsafe()
        rec: dict[str, Any] = {
            "ical_uid": ical_uid,
            "sequence": int(sequence),
        }
        if last_appt_date and last_appt_time:
            rec["last_appt_date"] = str(last_appt_date).strip()
            rec["last_appt_time"] = str(last_appt_time).strip()
        if last_duration_minutes is not None:
            rec["last_duration_minutes"] = int(last_duration_minutes)
        data[appointment_key] = rec
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


def resolve_store_key_for_cancel(pn: str, primary_key: str) -> str | None:
    """
    Resolve which store key to use for cancel/delete.

    1. If ``primary_key`` (from Action Date/Time) exists in the store, use it.
    2. Else, if exactly **one** key exists for this PN, use it (common when the Action row
       shows the *new* time after a reschedule but the store is still keyed by the *original* slot).

    If multiple keys exist for the same PN, returns None (ambiguous — align Date/Time on the
    cancel row with the key in ``event_id_store.json``, or cancel one visit per run).
    """
    pn_s = str(pn or "").strip()
    if get_invite_state(primary_key):
        return primary_key
    if not pn_s:
        return None
    prefix = f"{pn_s}_"
    with _lock:
        data = _load_unsafe()
    matches = [k for k in data.keys() if str(k).startswith(prefix)]
    if len(matches) == 1:
        return matches[0]
    return None
