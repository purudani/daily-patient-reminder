#!/usr/bin/env python3
"""
Dry-run simulator for the daily reminder job.

It does NOT send emails. It evaluates the same logic used by run_daily.py and
writes an Excel workbook showing row-level decisions and the final actions that
would be sent.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from config import COL_PATIENT_NAME, SIMULATION_FOLDER
from excel_reader import evaluate_daily_actions, load_action_df


def generate_simulation_workbook(
    evaluation: dict[str, list[dict[str, object]]] | None = None,
) -> Path:
    """Write the simulation workbook and return its path."""
    evaluation = evaluation or evaluate_daily_actions()
    decisions = evaluation["decisions"]
    final_actions = evaluation["actions"]

    df = pd.DataFrame(decisions)
    if not df.empty and "patient_name" in df.columns:
        action_df = load_action_df()
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

    final_actions_df = pd.DataFrame(
        [
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
            }
            for item in final_actions
        ]
    )

    out_dir = Path(SIMULATION_FOLDER).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"simulation_{stamp}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Simulation", index=False)
        final_actions_df.to_excel(writer, sheet_name="Final Actions", index=False)
    return out_path


def main() -> int:
    out_path = generate_simulation_workbook()
    print(f"Wrote simulation log: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
