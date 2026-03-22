#!/usr/bin/env python3
"""
Generate dummy Action file + Actual file + Mailchimp-style export.

Default (**minimal**): one row per core scenario so you can verify each path in isolation.
Use **--full** for the older large matrix (many edge cases).

Run from project root:
  python scripts/create_dummy_data.py
  python scripts/create_dummy_data.py --full

Outputs (project folder by default):
  - dummy_action_report.xlsx, dummy_actual_report.xlsx, dummy_mailchimp.xlsx
  - dummy_cancel_followup.xlsx (minimal only; see docstring below)
"""
from __future__ import annotations

import argparse
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

TEST_EMAIL = "purudani.2015@gmail.com"
BASE = Path(__file__).resolve().parent.parent

# Column names must match config.py (Scheduler Activity Report)
C = {
    "PN": "PN",
    "Patient Name": "Patient Name",
    "Type": "Type",
    "Location": "Location",
    "Provider": "Provider",
    "Date": "Date",
    "Time": "Time",
    "Action": "Action",
    "Reason": "Reason",
    "Comment": "Comment",
    "Reschedule Into": "Reschedule Into",
    "Action Date": "Action Date",
    "Action Time": "Action Time",
    "User ID": "User ID",
    "User Name": "User Name",
    "Has newer": "Has newer",
    "Date Year\\Month": "Date Year\\Month",
    "Appointment ID": "Appointment ID",
}


def _md(d: datetime) -> str:
    return d.strftime("%m/%d/%Y")


def _ym(d: datetime) -> str:
    return d.strftime("%y-%m")


def _row(
    pn: int | None,
    name: str,
    typ: str,
    loc: str,
    prov: str,
    dt: datetime,
    tm: str,
    action: str,
    reason: str = "",
    comment: str = "",
    reschedule_into: str = "",
    action_dt: str = "",
    action_tm: str = "",
    user_id: str = "ADTC",
    user_name: str = "tcruz",
    has_newer: str = "",
    appt_id: str = "",
) -> dict:
    return {
        C["PN"]: float("nan") if pn is None else pn,
        C["Patient Name"]: name,
        C["Type"]: typ,
        C["Location"]: loc,
        C["Provider"]: prov,
        C["Date"]: _md(dt),
        C["Time"]: tm,
        C["Action"]: action,
        C["Reason"]: reason,
        C["Comment"]: comment,
        C["Reschedule Into"]: reschedule_into,
        C["Action Date"]: action_dt or _md(datetime.now()),
        C["Action Time"]: action_tm or "05:36p",
        C["User ID"]: user_id,
        C["User Name"]: user_name,
        C["Has newer"]: has_newer,
        C["Date Year\\Month"]: _ym(dt),
        C["Appointment ID"]: appt_id,
    }


def write_full(out_dir: Path) -> None:
    """Large matrix: many PNs and edge cases (legacy)."""
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    t1 = today + timedelta(days=3)
    t2 = today + timedelta(days=4)
    t3 = today + timedelta(days=5)
    t4 = today + timedelta(days=6)
    t5 = today + timedelta(days=7)
    t6 = today + timedelta(days=8)
    t7 = today + timedelta(days=9)
    t8 = today + timedelta(days=10)
    t9 = today + timedelta(days=11)
    t10 = today + timedelta(days=12)

    P = {
        "1": 100000101,
        "2": 100000102,
        "3": 100000103,
        "4": 100000104,
        "5": 100000105,
        "6": 100000106,
        "7": 100000107,
        "8": 100000108,
        "9": 100000109,
        "10": 100000110,
        "11": 100000111,
        "12": 100000112,
        "13": 100000113,
        "15": 100000115,
        "16": 100000116,
        "17": 100000117,
    }

    action_rows = [
        _row(P["1"], "O'Connor, Steven", "30DN", "LIB", "PTNB", t1, "10:00a", "CREATE", appt_id="APT-001"),
        _row(P["2"], "Lister, Balthazar", "MT50", "LIBN", "MTJM", t2, "11:30a", "CREATE", appt_id="APT-002"),
        _row(
            P["3"],
            "TestPatient, Alpha",
            "PTDN",
            "LIB",
            "PTNS",
            t3,
            "12:00p",
            "RESCHEDULE",
            reschedule_into="Time: 12:00p -> 10:00a",
            appt_id="APT-003",
        ),
        _row(
            P["4"],
            "TestPatient, Beta",
            "LPTDN",
            "LIB",
            "PTTS",
            t4,
            "06:00p",
            "RESCHEDULE",
            has_newer="Yes",
            appt_id="APT-004",
        ),
        _row(P["5"], "TestPatient, Gamma", "UNCPT", "LIBN", "MTSH", t5, "09:00a", "RESCHEDULE", has_newer="Yes", appt_id="APT-005"),
        _row(P["6"], "TestPatient, Delta", "IENEW", "LIB", "PTAM2", t6, "08:00a", "RESCHEDULE", has_newer="Yes", appt_id="APT-006"),
        _row(P["7"], "TestPatient, Epsilon", "30DN", "LIBJ", "PTRP", t7, "04:30p", "CANCEL w. remove", reason="CONF", appt_id="APT-007"),
        _row(P["8"], "TestPatient, Zeta", "MT50", "LIBJ", "PTTS", t8, "02:00p", "DELETE", reason="NOREA", appt_id="APT-008"),
        _row(P["9"], "TestPatient, Eta", "PTDN", "LIB", "PTNB", t9, "01:00p", "CREATE", has_newer="Yes", appt_id="APT-009"),
        _row(P["10"], "TestPatient, Theta", "30DN", "LIBN", "PTNS", t10, "03:00p", "DELETE", has_newer="Yes", appt_id="APT-010"),
        _row(P["11"], "TestPatient, Iota", "LPTDN", "LIB", "MTJM", t1, "05:00p", "CANCEL w. remove", has_newer="Yes", appt_id="APT-011"),
        _row(P["12"], "TestPatient, Kappa", "UNCPT", "LIBJ", "PTTS", t2, "04:00p", "CANCEL w. remove", has_newer="Yes", appt_id="APT-012"),
        _row(P["13"], "TestPatient, Lambda", "IEATH", "LIBJ", "PTEC", t5, "01:30p", "CREATE", appt_id="APT-013"),
        _row(None, "", "30DN", "LIB", "PTNB", t6, "10:00a", "CREATE"),
        _row(P["15"], "TestPatient, Mu", "PTDN", "LIB", "PTNB", t7, "10:00a", "EDIT"),
        _row(P["16"], "TestPatient, Nu", "30DN", "LIB", "PTNB", today, "02:00p", "CREATE", appt_id="APT-016"),
        _row(P["17"], "TestPatient, Xi", "MT50", "LIBN", "MTSH", t8, "09:00a", "CREATE", appt_id="APT-017"),
        _row(
            P["17"],
            "TestPatient, Xi",
            "MT50",
            "LIBN",
            "MTSH",
            t8,
            "09:00a",
            "RESCHEDULE",
            reschedule_into=f"Date: {_md(t8)} -> {_md(t9)} Time: 09:00a -> 03:00p",
            appt_id="APT-017",
        ),
        _row(
            P["3"],
            "TestPatient, Alpha",
            "PTDN",
            "LIB",
            "PTNS",
            t3,
            "10:00a",
            "RESCHEDULE",
            reschedule_into=f"Date: {t3.year}-{t3.month:02d}-{t3.day:02d} -> {t4.year}-{t4.month:02d}-{t4.day:02d} Time: 10:00a -> 02:30p",
            appt_id="APT-003B",
        ),
        _row(P["7"], "TestPatient, Epsilon", "30DN", "LIBJ", "PTRP", t7, "04:30p", "CREATE", appt_id="APT-007"),
        _row(P["8"], "TestPatient, Zeta", "MT50", "LIBJ", "PTTS", t8, "02:00p", "CREATE", appt_id="APT-008"),
    ]

    action_df = pd.DataFrame(action_rows)

    actual_rows = [
        {
            C["PN"]: P["4"],
            C["Patient Name"]: "TestPatient, Beta",
            C["Type"]: "LPTDN",
            C["Location"]: "LIBN",
            C["Provider"]: "PTTS",
            C["Date"]: _md(t5),
            C["Time"]: "02:30p",
            C["Action"]: "",
            C["Appointment ID"]: "APT-004",
        },
        {
            C["PN"]: P["6"],
            C["Patient Name"]: "TestPatient, Delta",
            C["Type"]: "IENEW",
            C["Location"]: "LIB",
            C["Provider"]: "PTAM2",
            C["Date"]: _md(today),
            C["Time"]: "12:00p",
            C["Action"]: "",
            C["Appointment ID"]: "APT-006",
        },
        {
            C["PN"]: P["12"],
            C["Patient Name"]: "TestPatient, Kappa",
            C["Type"]: "UNCPT",
            C["Location"]: "LIBJ",
            C["Provider"]: "PTTS",
            C["Date"]: _md(t2),
            C["Time"]: "04:00p",
            C["Action"]: "",
            C["Appointment ID"]: "APT-012",
        },
    ]
    actual_df = pd.DataFrame(actual_rows)

    mc_pns = list(P.values())
    mc_rows = []
    for i, pn in enumerate(mc_pns):
        last = f"TestLast{i+1:02d}"
        first = f"TestFirst{i+1:02d}"
        mc_rows.append(
            {
                "PN": pn,
                "Prog": "AC001, PT",
                "Last": last,
                "First": first,
                "DOB": "09/08/1954",
                "Birthday": "09/08",
                "Sex": "F",
                "Email": TEST_EMAIL,
                "SMS Phone Num": 12013334444,
                "City": "Jersey City",
                "Zip": "07302",
                "Marketing P": "Returning P",
                "Tot. Act. Vis": 50,
                "Pend. Visits": 9,
                "Total Visits": 59,
                "Last Actual": "03/02/2026",
                "First Visit D": "10/01/2025",
                "Last Schedul": _md(today + timedelta(days=20)),
                "Payor FC": "MC",
                "Key Providi": "PTEC",
                "Case Type": "Other, Low",
                "Case Reason": "Pain",
                "Case Status": "Active",
                "Payor End I": "12/31/2026",
                "Payments": 1079.08,
                "Payor Name": "MEDICARE",
                "Payor Group": "",
                "Payor SSID": "1WD3T209",
                "Pay Per Act": 61.58,
            }
        )

    mailchimp_df = pd.DataFrame(mc_rows)

    action_path = out_dir / "dummy_action_report.xlsx"
    with pd.ExcelWriter(action_path, engine="openpyxl") as w:
        action_df.to_excel(w, sheet_name="Action", startrow=2, index=False)

    actual_path = out_dir / "dummy_actual_report.xlsx"
    with pd.ExcelWriter(actual_path, engine="openpyxl") as w:
        actual_df.to_excel(w, sheet_name="Actual", startrow=2, index=False)

    mailchimp_path = out_dir / "dummy_mailchimp.xlsx"
    mailchimp_df.to_excel(mailchimp_path, sheet_name="Sheet1", index=False, engine="openpyxl")

    print(f"Wrote {action_path}")
    print(f"Wrote {actual_path}")
    print(f"Wrote {mailchimp_path}")
    print("(full matrix) See TESTING.md for scenario notes.")


def write_minimal(out_dir: Path) -> None:
    """
    Exactly three Action rows (unique Appointment IDs — no dedupe collisions):

    | Appointment ID | Scenario |
    |----------------|----------|
    | APT-M-CREATE   | CREATE — new invite, SEQUENCE 0 |
    | APT-M-RESCHED  | RESCHEDULE — **Reschedule Into** only (time change) |
    | APT-M-ACTUAL   | RESCHEDULE + **Has newer** — slot from **Actual** sheet |

    **Cancel/Delete** cannot run in the same file as CREATE for the same ID (last-wins
    dedupe). After **run 1** with `dummy_action_report.xlsx`, the store has UID for
    **APT-M-CREATE**. Run **run 2** with `ACTION_REPORT_PATH` pointing at
    **dummy_cancel_followup.xlsx** (one DELETE row for that ID) to verify **cancel.ics**.

    Optional: set `EVENT_ID_STORE_PATH` to a test-only JSON path so production store is untouched.
    """
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    t_create = today + timedelta(days=3)
    t_resched = today + timedelta(days=4)
    # Has newer: action row = old slot; Actual = new future slot
    t_act_action = today + timedelta(days=5)
    t_act_actual = today + timedelta(days=7)

    P_CREATE = 100000201
    P_RESCHED = 100000202
    P_ACTUAL = 100000203

    action_rows = [
        _row(
            P_CREATE,
            "Smoke, CreateOnly",
            "30DN",
            "LIB",
            "PTNB",
            t_create,
            "10:00a",
            "CREATE",
            appt_id="APT-M-CREATE",
        ),
        _row(
            P_RESCHED,
            "Smoke, RescheduleInto",
            "PTDN",
            "LIB",
            "PTNS",
            t_resched,
            "12:00p",
            "RESCHEDULE",
            reschedule_into="Time: 12:00p -> 10:00a",
            appt_id="APT-M-RESCHED",
        ),
        _row(
            P_ACTUAL,
            "Smoke, HasNewerActual",
            "LPTDN",
            "LIB",
            "PTTS",
            t_act_action,
            "06:00p",
            "RESCHEDULE",
            has_newer="Yes",
            appt_id="APT-M-ACTUAL",
        ),
    ]
    action_df = pd.DataFrame(action_rows)

    actual_rows = [
        {
            C["PN"]: P_ACTUAL,
            C["Patient Name"]: "Smoke, HasNewerActual",
            C["Type"]: "LPTDN",
            C["Location"]: "LIBN",
            C["Provider"]: "PTTS",
            C["Date"]: _md(t_act_actual),
            C["Time"]: "02:30p",
            C["Action"]: "",
            C["Appointment ID"]: "APT-M-ACTUAL",
        },
    ]
    actual_df = pd.DataFrame(actual_rows)

    mc_rows = []
    for pn, last, first in (
        (P_CREATE, "CreateOnly", "Smoke"),
        (P_RESCHED, "RescheduleInto", "Smoke"),
        (P_ACTUAL, "HasNewerActual", "Smoke"),
    ):
        mc_rows.append(
            {
                "PN": pn,
                "Prog": "AC001, PT",
                "Last": last,
                "First": first,
                "DOB": "01/15/1980",
                "Email": TEST_EMAIL,
                "City": "Jersey City",
                "Zip": "07302",
            }
        )
    mailchimp_df = pd.DataFrame(mc_rows)

    action_path = out_dir / "dummy_action_report.xlsx"
    with pd.ExcelWriter(action_path, engine="openpyxl") as w:
        action_df.to_excel(w, sheet_name="Action", startrow=2, index=False)

    actual_path = out_dir / "dummy_actual_report.xlsx"
    with pd.ExcelWriter(actual_path, engine="openpyxl") as w:
        actual_df.to_excel(w, sheet_name="Actual", startrow=2, index=False)

    mailchimp_path = out_dir / "dummy_mailchimp.xlsx"
    mailchimp_df.to_excel(mailchimp_path, sheet_name="Sheet1", index=False, engine="openpyxl")

    # Second-run cancel test: same PN / date / time / ID as CREATE row
    cancel_rows = [
        _row(
            P_CREATE,
            "Smoke, CreateOnly",
            "30DN",
            "LIB",
            "PTNB",
            t_create,
            "10:00a",
            "DELETE",
            reason="TEST",
            appt_id="APT-M-CREATE",
        ),
    ]
    cancel_df = pd.DataFrame(cancel_rows)
    cancel_path = out_dir / "dummy_cancel_followup.xlsx"
    with pd.ExcelWriter(cancel_path, engine="openpyxl") as w:
        cancel_df.to_excel(w, sheet_name="Action", startrow=2, index=False)
    # Minimal Actual for second run (job still loads Actual; one row is fine)
    with pd.ExcelWriter(out_dir / "dummy_actual_report_cancel_only.xlsx", engine="openpyxl") as w:
        pd.DataFrame(
            [
                {
                    C["PN"]: P_CREATE,
                    C["Patient Name"]: "Smoke, CreateOnly",
                    C["Type"]: "30DN",
                    C["Location"]: "LIB",
                    C["Provider"]: "PTNB",
                    C["Date"]: _md(t_create),
                    C["Time"]: "10:00a",
                    C["Action"]: "",
                    C["Appointment ID"]: "APT-M-CREATE",
                }
            ]
        ).to_excel(w, sheet_name="Actual", startrow=2, index=False)

    print(f"Wrote {action_path}")
    print(f"Wrote {actual_path}")
    print(f"Wrote {mailchimp_path}")
    print(f"Wrote {cancel_path}  (use for run 2 cancel/delete test)")
    print(f"Wrote {out_dir / 'dummy_actual_report_cancel_only.xlsx'}  (optional Actual for run 2)")
    print()
    print("Minimal scenarios (expect 3 emails on run 1):")
    print("  APT-M-CREATE   -> create (invite.ics)")
    print("  APT-M-RESCHED  -> reschedule via Reschedule Into (invite.ics, SEQUENCE)")
    print("  APT-M-ACTUAL   -> reschedule using Actual sheet (invite.ics)")
    print()
    print("Run 2 (after run 1 populated event_id_store for APT-M-CREATE):")
    print(f"  ACTION_REPORT_PATH={cancel_path}")
    print(f"  ACTUAL_REPORT_PATH={out_dir / 'dummy_actual_report_cancel_only.xlsx'}  (or keep dummy_actual_report.xlsx)")
    print("Expect 1 email with cancel.ics; invite_sent_log.xlsx gains a row with action=delete.")
    print(f"All test invites -> {TEST_EMAIL}")


def main() -> None:
    ap = argparse.ArgumentParser(description="Generate dummy Excel files for daily-patient-reminder.")
    ap.add_argument(
        "--full",
        action="store_true",
        help="Generate the large matrix (legacy) instead of minimal smoke scenarios.",
    )
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=BASE,
        help="Output directory (default: project root).",
    )
    args = ap.parse_args()
    out_dir = args.out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.full:
        write_full(out_dir)
    else:
        write_minimal(out_dir)

    print()
    print("Set ACTION_REPORT_PATH / ACTUAL_REPORT_PATH / MAILCHIMP_EXPORT_PATH in .env.")
    print("Optional audit: INVITE_LOG_PATH (default invite_sent_log.xlsx in project folder).")


if __name__ == "__main__":
    main()
