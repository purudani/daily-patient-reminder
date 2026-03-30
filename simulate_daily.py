#!/usr/bin/env python3
"""
Dry-run simulator for the daily reminder job.

It does NOT send emails. It evaluates every Action row and writes a decision log
showing what the automation would do and why.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

# Load .env before config import so REFERENCE_DATE and paths apply in simulation too.
try:
    from dotenv import load_dotenv

    _env_file = Path(__file__).resolve().parent / ".env"
    if _env_file.exists():
        load_dotenv(_env_file)
except ImportError:
    pass

from config import (
    COL_ACTION,
    COL_ACTION_TIME,
    COL_APPT_DATE,
    COL_APPT_TIME,
    COL_APPT_TYPE,
    COL_PATIENT_NAME,
    COL_PN,
    LOG_FOLDER,
)
from excel_reader import (
    ACTIONS_CANCEL,
    ACTIONS_CREATE,
    ACTIONS_DELETE,
    ACTIONS_RESCHEDULE,
    SKIP_BLANK_PN,
    SKIP_NEXT_DAY,
    SKIP_SAME_DAY,
    _action_sort_tuple,
    _action_row_to_record,
    _actual_column,
    _appointment_group_key,
    _cell_value,
    _find_actual_row_for_has_newer,
    _has_newer_value,
    _is_future_beyond_next_day,
    _is_next_day,
    _is_same_day,
    _is_uncpt,
    _normalize_action,
    _normalize_pn,
    _parse_date,
    _parse_time,
    _parse_time_optional,
    load_action_df,
    load_actual_df,
    load_mailchimp_lookup,
)


def _decision_row(
    idx: int,
    action_raw: Any,
    action_norm: str | None,
    pn: str,
    has_newer: bool,
    decision: str,
    reason: str,
    *,
    used_actual: bool = False,
    would_send: bool = False,
    send_kind: str = "",
    appt_date: str = "",
    appt_time: str = "",
    appt_type: str = "",
    email: str = "",
    patient_name: str = "",
) -> dict[str, Any]:
    return {
        "row_index": idx,
        "pn": pn,
        "patient_name": patient_name,
        "raw_action": str(action_raw or ""),
        "normalized_action": action_norm or "",
        "has_newer": has_newer,
        "decision": decision,
        "reason": reason,
        "used_actual": used_actual,
        "would_send": would_send,
        "send_kind": send_kind,
        "appt_date": appt_date,
        "appt_time": appt_time,
        "appt_type": appt_type,
        "email": email,
    }


def main() -> int:
    simulation_default_email = "simulation@example.com"

    action_df = load_action_df()
    actual_df = load_actual_df()
    mailchimp = load_mailchimp_lookup()
    cols = list(action_df.columns)
    actual_cols_list = list(actual_df.columns) if not actual_df.empty else None
    from config import ACTUAL_COL_DATE, ACTUAL_COL_LOCATION, ACTUAL_COL_TIME, ACTUAL_COL_TYPE

    actual_field_tuple = (
        _actual_column(ACTUAL_COL_DATE, COL_APPT_DATE),
        _actual_column(ACTUAL_COL_TIME, COL_APPT_TIME),
        _actual_column(ACTUAL_COL_LOCATION, "Location"),
        _actual_column(ACTUAL_COL_TYPE, COL_APPT_TYPE),
    )

    out: list[dict[str, Any]] = []
    candidates: list[dict[str, Any]] = []
    for idx, row in action_df.iterrows():
        row_patient_name = str(_cell_value(row, COL_PATIENT_NAME, cols) or "").strip()
        pn = _normalize_pn(_cell_value(row, COL_PN, cols))
        action_val = _cell_value(row, COL_ACTION, cols)
        action = _normalize_action(str(action_val).strip()) if action_val is not None else None
        if action is None:
            raw = str(action_val or "").strip()
            reason = (
                f"unsupported action: {raw}"
                if raw
                else "unsupported/blank action"
            )
            out.append(_decision_row(idx, action_val, action, pn, False, "skip", reason, patient_name=row_patient_name))
            continue

        if SKIP_BLANK_PN and not pn:
            out.append(_decision_row(idx, action_val, action, pn, False, "skip", "blank PN", patient_name=row_patient_name))
            continue

        hn = _has_newer_value(row, cols)
        has_newer = hn is not None and str(hn).strip().lower() in ("true", "1", "yes", "y")
        original_appt_date = _parse_date(_cell_value(row, COL_APPT_DATE, cols))

        if not has_newer:
            if _is_uncpt(_cell_value(row, COL_APPT_TYPE, cols)):
                out.append(_decision_row(idx, action_val, action, pn, False, "skip", "UNCPT on action row", patient_name=row_patient_name))
                continue
            if SKIP_SAME_DAY and _is_same_day(original_appt_date):
                out.append(_decision_row(idx, action_val, action, pn, has_newer, "skip", "same-day (Action appt date)", patient_name=row_patient_name))
                continue

        need_actual_lookup = has_newer
        actual_row = None
        if need_actual_lookup:
            actual_row = _find_actual_row_for_has_newer(actual_df, pn, row, cols)
            if actual_row is None:
                out.append(
                    _decision_row(
                        idx,
                        action_val,
                        action,
                        pn,
                        has_newer,
                        "skip",
                        "has_newer: no usable Actual row after PN + Action Date/Time match and appointment date/time narrowing",
                        patient_name=row_patient_name,
                    )
                )
                continue

        rec = _action_row_to_record(
            row,
            cols,
            mailchimp,
            actual_row if need_actual_lookup else None,
            action in ACTIONS_RESCHEDULE and not need_actual_lookup,
            actual_field_tuple,
            actual_cols_list if need_actual_lookup else None,
        )
        if rec is None:
            out.append(_decision_row(idx, action_val, action, pn, has_newer, "skip", "record build failed", patient_name=row_patient_name))
            continue

        if _is_uncpt(rec.get("appt_type")):
            out.append(_decision_row(idx, action_val, action, pn, has_newer, "skip", "UNCPT on actual/resolved record", used_actual=need_actual_lookup, patient_name=row_patient_name))
            continue

        moved_future_to_next_day = (
            action in ACTIONS_RESCHEDULE
            and _is_future_beyond_next_day(original_appt_date)
            and _is_next_day(rec.get("appt_date"))
        )

        if SKIP_SAME_DAY and _is_same_day(rec.get("appt_date")):
            reason = "same-day (Actual appt date)" if need_actual_lookup else "same-day (resolved appt date)"
            out.append(_decision_row(idx, action_val, action, pn, has_newer, "skip", reason, used_actual=need_actual_lookup, patient_name=row_patient_name))
            continue
        if SKIP_NEXT_DAY and _is_next_day(rec.get("appt_date")) and not moved_future_to_next_day:
            reason = "next-day (Actual appt date)" if need_actual_lookup else "next-day (resolved appt date)"
            out.append(_decision_row(idx, action_val, action, pn, has_newer, "skip", reason, used_actual=need_actual_lookup, patient_name=row_patient_name))
            continue

        if action in ACTIONS_RESCHEDULE:
            rec["original_appt_date"] = original_appt_date
            rec["original_appt_time"] = _parse_time(_cell_value(row, COL_APPT_TIME, cols))
        elif need_actual_lookup and action in (*ACTIONS_CANCEL, *ACTIONS_DELETE):
            rec["original_appt_date"] = original_appt_date
            rec["original_appt_time"] = _parse_time(_cell_value(row, COL_APPT_TIME, cols))

        email_missing = not rec.get("email")
        if email_missing:
            rec["email"] = simulation_default_email

        send_kind = "invite.ics" if action in (*ACTIONS_CREATE, *ACTIONS_RESCHEDULE) else "cancel.ics"
        out.append(_decision_row(
            idx,
            action_val,
            action,
            pn,
            has_newer,
            "process",
            "candidate before grouped-action resolution",
            used_actual=need_actual_lookup,
            would_send=True,
            send_kind=send_kind,
            appt_date=str(rec.get("appt_date") or ""),
            appt_time=str(rec.get("appt_time") or ""),
            appt_type=str(rec.get("appt_type") or ""),
            email=str(rec.get("email") or ""),
            patient_name=str(rec.get("patient_name") or row_patient_name or ""),
        ))
        candidates.append({
            "row_index": idx,
            "action": action,
            "record": rec,
            "action_time": _parse_time_optional(_cell_value(row, COL_ACTION_TIME, cols)) or "",
            "appointment_group_key": _appointment_group_key(action, rec),
            "email_missing": email_missing,
        })

    decisions_by_index = {row["row_index"]: row for row in out}
    grouped: dict[str, list[dict[str, Any]]] = {}
    for item in candidates:
        grouped.setdefault(item["appointment_group_key"], []).append(item)

    kept_indexes: set[int] = set()
    for _, group in grouped.items():
        group.sort(key=_action_sort_tuple)
        actions = {item["action"] for item in group}
        if len(group) == 2 and actions == {"create", "delete"}:
            for item in group:
                decision = decisions_by_index[item["row_index"]]
                decision["decision"] = "skip"
                decision["would_send"] = False
                decision["send_kind"] = ""
                decision["reason"] = "same appointment has create/delete pair in one run"
            continue
        survivor = group[-1]
        kept_indexes.add(survivor["row_index"])
        for item in group[:-1]:
            decision = decisions_by_index[item["row_index"]]
            decision["decision"] = "skip"
            decision["would_send"] = False
            decision["send_kind"] = ""
            decision["reason"] = "superseded by later action for same appointment"

    for item in candidates:
        if item["row_index"] in kept_indexes:
            decisions_by_index[item["row_index"]]["reason"] = (
                "would send after grouped-action resolution"
                + (" (using placeholder email)" if item["email_missing"] else "")
            )

    df = pd.DataFrame(out)
    # Ensure patient_name is always present and appears right after PN.
    if not df.empty and "patient_name" in df.columns:
        if COL_PATIENT_NAME in action_df.columns:
            names_by_idx = action_df[COL_PATIENT_NAME].fillna("").astype(str).to_dict()
            df["patient_name"] = df["patient_name"].where(
                df["patient_name"].astype(str).str.strip() != "",
                df["row_index"].map(names_by_idx).fillna(""),
            )
        ordered = ["row_index", "pn", "patient_name"] + [
            c for c in df.columns if c not in {"row_index", "pn", "patient_name"}
        ]
        df = df[ordered]
    out_dir = Path(LOG_FOLDER).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"simulation_{stamp}.xlsx"
    final_actions = pd.DataFrame([
        {
            "action": item["action"],
            "pn": item["record"].get("pn", ""),
            "patient_name": item["record"].get("patient_name", ""),
            "email": item["record"].get("email", ""),
            "appt_date": item["record"].get("appt_date", ""),
            "appt_time": item["record"].get("appt_time", ""),
            "original_appt_date": item["record"].get("original_appt_date", ""),
            "original_appt_time": item["record"].get("original_appt_time", ""),
            "location": item["record"].get("location", ""),
            "appt_type": item["record"].get("appt_type", ""),
            "email_source": "placeholder" if item["email_missing"] else "mailchimp",
        }
        for item in sorted(candidates, key=_action_sort_tuple)
        if item["row_index"] in kept_indexes
    ])
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Simulation", index=False)
        final_actions.to_excel(w, sheet_name="Final Actions", index=False)
    print(f"Wrote simulation log: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
