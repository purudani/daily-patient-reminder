#!/usr/bin/env python3
"""
One-time launch sender for a simple workbook with these columns:

- Date
- Time
- Patient #
- Patient First Name
- Patient Last Name
- Location
- App Type
- Email

Default mode is dry-run. Use --send to actually send confirmations.
"""
from __future__ import annotations

import argparse
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

try:
    from dotenv import load_dotenv

    _env_file = Path(__file__).resolve().parent / ".env"
    if _env_file.exists():
        load_dotenv(_env_file)
except ImportError:
    pass

from calendar_actions import do_create
from config import LOCATION_MAP, LOG_FOLDER, SKIP_NEXT_DAY, SKIP_SAME_DAY
from event_id_store import get_invite_state
from excel_reader import _is_uncpt, _normalize_pn, _parse_date, _parse_time
from graph_auth import get_access_token

DATE_COL = "Date"
TIME_COL = "Time"
PN_COL = "Patient #"
FIRST_COL = "Patient First Name"
LAST_COL = "Patient Last Name"
LOCATION_COL = "Location"
TYPE_COL = "App Type"
EMAIL_COL = "Email"


def _decision(
    idx: int,
    pn: str,
    patient_name: str,
    decision: str,
    reason: str,
    *,
    appt_date: str = "",
    appt_time: str = "",
    appt_type: str = "",
    email: str = "",
    key: str = "",
    sent: bool = False,
    reference_today: str = "",
) -> dict[str, Any]:
    return {
        "row_index": idx,
        "pn": pn,
        "patient_name": patient_name,
        "action": "launch_create",
        "decision": decision,
        "reason": reason,
        "appt_date": appt_date,
        "appt_time": appt_time,
        "appt_type": appt_type,
        "email": email,
        "appointment_key": key,
        "sent": sent,
        "evaluated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "reference_today": reference_today,
    }


def _reference_today(raw: str) -> date:
    raw = (raw or "").strip()
    if raw:
        return datetime.strptime(raw, "%Y-%m-%d").date()
    return datetime.now().date()


def _is_same_day_for(date_val: Any, ref_today: date) -> bool:
    if not date_val:
        return True
    try:
        return pd.to_datetime(date_val).date() == ref_today
    except Exception:
        return False


def _is_next_day_for(date_val: Any, ref_today: date) -> bool:
    if not date_val:
        return False
    try:
        return pd.to_datetime(date_val).date() == (ref_today + timedelta(days=1))
    except Exception:
        return False


def _load_report(path_arg: str) -> pd.DataFrame:
    path = Path(path_arg).expanduser()
    if not path.exists():
        raise FileNotFoundError(f"Launch report not found: {path}")
    df = pd.read_excel(path, sheet_name=0)
    df = df.rename(columns=lambda c: (c or "").strip())
    missing = [c for c in (DATE_COL, TIME_COL, PN_COL, FIRST_COL, LAST_COL, LOCATION_COL, TYPE_COL, EMAIL_COL) if c not in df.columns]
    if missing:
        raise ValueError(f"Launch report is missing required columns: {', '.join(missing)}")
    return df


def main() -> int:
    ap = argparse.ArgumentParser(description="Send one-time launch confirmations from a simple launch workbook.")
    ap.add_argument("--report-path", required=True, help="Path to the launch workbook (.xlsx). First sheet is used.")
    ap.add_argument("--send", action="store_true", help="Actually send confirmations. Default is dry-run.")
    ap.add_argument("--limit", type=int, default=0, help="Optional max rows to process (0 = no limit).")
    ap.add_argument("--reference-date", default="", help="Optional YYYY-MM-DD override for same-day/next-day checks.")
    args = ap.parse_args()

    df = _load_report(args.report_path)
    ref_today = _reference_today(args.reference_date)
    ref_today_iso = ref_today.isoformat()
    out: list[dict[str, Any]] = []

    token: str | None = None
    if args.send:
        token = get_access_token()

    processed = 0
    for idx, row in df.iterrows():
        if args.limit > 0 and processed >= args.limit:
            break

        pn = _normalize_pn(row.get(PN_COL))
        first = str(row.get(FIRST_COL) or "").strip()
        last = str(row.get(LAST_COL) or "").strip()
        patient_name = " ".join(part for part in (first, last) if part).strip()
        appt_type = str(row.get(TYPE_COL) or "").strip().upper()
        appt_date = _parse_date(row.get(DATE_COL)) or ""
        appt_time = _parse_time(row.get(TIME_COL))
        email = str(row.get(EMAIL_COL) or "").strip()
        loc_code = str(row.get(LOCATION_COL) or "LIB").strip().upper() or "LIB"
        location_address = LOCATION_MAP.get(loc_code) or loc_code
        key = f"{pn}_{appt_date}_{appt_time}"

        if not pn:
            out.append(_decision(idx, pn, patient_name, "skip", "blank PN", reference_today=ref_today_iso))
            continue
        if _is_uncpt(appt_type):
            out.append(_decision(idx, pn, patient_name, "skip", "UNCPT", appt_date=appt_date, appt_time=appt_time, appt_type=appt_type, reference_today=ref_today_iso))
            continue
        if not appt_date:
            out.append(_decision(idx, pn, patient_name, "skip", "missing appointment date", appt_time=appt_time, appt_type=appt_type, reference_today=ref_today_iso))
            continue
        if SKIP_SAME_DAY and _is_same_day_for(appt_date, ref_today):
            out.append(_decision(idx, pn, patient_name, "skip", "same-day", appt_date=appt_date, appt_time=appt_time, appt_type=appt_type, email=email, reference_today=ref_today_iso))
            continue
        if SKIP_NEXT_DAY and _is_next_day_for(appt_date, ref_today):
            out.append(_decision(idx, pn, patient_name, "skip", "next-day", appt_date=appt_date, appt_time=appt_time, appt_type=appt_type, email=email, reference_today=ref_today_iso))
            continue
        if get_invite_state(key):
            out.append(_decision(idx, pn, patient_name, "skip", "already in event_id_store", appt_date=appt_date, appt_time=appt_time, appt_type=appt_type, email=email, key=key, reference_today=ref_today_iso))
            continue
        if not email:
            out.append(_decision(idx, pn, patient_name, "skip", "missing email", appt_date=appt_date, appt_time=appt_time, appt_type=appt_type, key=key, reference_today=ref_today_iso))
            continue

        rec = {
            "pn": pn,
            "email": email,
            "patient_name": patient_name or email,
            "appt_date": appt_date,
            "appt_time": appt_time,
            "location": loc_code,
            "location_address": location_address,
            "appt_type": appt_type,
        }

        processed += 1
        if not args.send:
            out.append(_decision(idx, pn, rec["patient_name"], "process", "would send (launch dry-run)", appt_date=appt_date, appt_time=appt_time, appt_type=appt_type, email=email, key=key, sent=False, reference_today=ref_today_iso))
            continue

        try:
            sent = do_create(token, rec, LOCATION_MAP)  # type: ignore[arg-type]
            out.append(_decision(idx, pn, rec["patient_name"], "process" if sent else "skip", "sent" if sent else "do_create returned False", appt_date=appt_date, appt_time=appt_time, appt_type=appt_type, email=email, key=key, sent=bool(sent), reference_today=ref_today_iso))
        except Exception as e:
            out.append(_decision(idx, pn, rec["patient_name"], "error", f"exception: {e}", appt_date=appt_date, appt_time=appt_time, appt_type=appt_type, email=email, key=key, sent=False, reference_today=ref_today_iso))

    out_dir = Path(LOG_FOLDER).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"one_time_send_{stamp}.xlsx"
    pd.DataFrame(out).to_excel(out_path, sheet_name="Launch", index=False)
    print(f"Wrote one-time log: {out_path}")
    print(f"Mode: {'SEND' if args.send else 'DRY-RUN'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
