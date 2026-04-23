#!/usr/bin/env python3
"""
Daily 8PM job: read action/actual/Mailchimp from Excel, send one patient email per action:
HTML + invite.ics (or cancel.ics) via Graph Mail.Send.

Run from cron / Task Scheduler:
  - Mac/Linux (cron, 8 PM):  0 20 * * * /usr/bin/env python3 /path/to/run_daily.py
  - Windows (Task Scheduler): create daily task at 8:00 PM, action: python run_daily.py

Data flow:
  - Action column + Reschedule/Location columns (and Actual sheet when "Has Newer Action")
  - Location codes (LIB, LIBN, LIBJ) → full address in mail via config.LOCATION_MAP
  - Patient email from Mailchimp by PN
"""
from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

# Load .env before config is imported (optional; keeps secrets out of repo)
try:
    from dotenv import load_dotenv
    _env_file = Path(__file__).resolve().parent / ".env"
    if _env_file.exists():
        load_dotenv(_env_file)
except ImportError:
    pass

from config import (
    GRAPH_CLIENT_ID,
    GRAPH_CLIENT_SECRET,
    GRAPH_MAILBOX_USER,
    GRAPH_TENANT_ID,
    LOG_FOLDER,
    LOG_INVITES_AND_CHANGES,
    LOCATION_MAP,
)
from excel_reader import evaluate_daily_actions
from calendar_actions import do_cancel, do_create, do_delete, do_reschedule
from graph_auth import get_access_token
from simulate_daily import generate_simulation_workbook


def setup_logging() -> logging.Logger:
    """Log to a file in LOG_FOLDER and to stderr."""
    log_dir = Path(LOG_FOLDER).expanduser()
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"reminder_{datetime.now().strftime('%Y%m%d')}.log"
    fmt = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    date_fmt = "%Y-%m-%d %H:%M:%S"

    root = logging.getLogger()
    root.setLevel(logging.INFO)
    for h in list(root.handlers):
        root.removeHandler(h)

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.INFO)
    fh.setFormatter(logging.Formatter(fmt, datefmt=date_fmt))
    root.addHandler(fh)

    ch = logging.StreamHandler(sys.stderr)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter(fmt, datefmt=date_fmt))
    root.addHandler(ch)

    logger = logging.getLogger("run_daily")
    logger.info("Log file: %s", log_file)
    return logger


def main() -> int:
    logger = setup_logging()
    if not (GRAPH_CLIENT_ID and str(GRAPH_CLIENT_ID).strip()):
        logger.error(
            "GRAPH_CLIENT_ID is not set. Set it in .env or environment. See README."
        )
        return 1
    secret = (GRAPH_CLIENT_SECRET or "").strip()
    if secret:
        if not (GRAPH_TENANT_ID or "").strip():
            logger.error(
                "GRAPH_TENANT_ID is required when GRAPH_CLIENT_SECRET is set (app-only auth)."
            )
            return 1
        if not (GRAPH_MAILBOX_USER or "").strip():
            logger.error(
                "GRAPH_MAILBOX_USER is required for app-only auth (mailbox that sends mail, e.g. deepak@libertyptnj.com). "
                "Mail API uses /users/{GRAPH_MAILBOX_USER}/sendMail — not /me."
            )
            return 1
        logger.info(
            "Graph mode: app-only (no sign-in); mailbox=%s",
            (GRAPH_MAILBOX_USER or "").strip(),
        )
    else:
        logger.info(
            "Graph mode: delegated (sign-in when prompted). "
            "Set GRAPH_MAILBOX_USER to use /users/{email}/... instead of /me."
        )
    try:
        evaluation = evaluate_daily_actions()
        simulation_path = generate_simulation_workbook(evaluation)
        logger.info("Simulation workbook: %s", simulation_path)

        actions = evaluation["actions"]
        logger.info("Actions to process: %d", len(actions))
        if not actions:
            return 0

        token = get_access_token()
        created = rescheduled = cancelled = errors = 0

        for item in actions:
            action = item["action"]
            record = item["record"]
            pn = record.get("pn", "")
            email = record.get("email", "")
            try:
                if action == "create":
                    if do_create(token, record, LOCATION_MAP):
                        created += 1
                        if LOG_INVITES_AND_CHANGES:
                            logger.info(
                                "Create sent: PN=%s %s at %s -> %s",
                                pn,
                                record.get("appt_date"),
                                record.get("appt_time"),
                                email,
                            )
                elif action == "reschedule":
                    if do_reschedule(token, record, LOCATION_MAP):
                        rescheduled += 1
                        if LOG_INVITES_AND_CHANGES:
                            logger.info(
                                "Reschedule sent: PN=%s %s at %s -> %s",
                                pn,
                                record.get("appt_date"),
                                record.get("appt_time"),
                                email,
                            )
                elif action in ("cancel", "delete"):
                    # Cancel and Delete are the same: send cancellation to patient and remove from store
                    cancelled_ok = (
                        do_cancel(token, record)
                        if action == "cancel"
                        else do_delete(token, record)
                    )
                    if cancelled_ok:
                        cancelled += 1
                        if LOG_INVITES_AND_CHANGES:
                            logger.info("Cancel/Delete sent: PN=%s", pn)
                    elif LOG_INVITES_AND_CHANGES:
                        logger.info(
                            "Cancel/Delete skipped: PN=%s (no stored invite UID; nothing to cancel)",
                            pn,
                        )
                else:
                    logger.warning("Unknown action '%s' for PN=%s; skipping", action, pn)
            except Exception as e:
                errors += 1
                logger.exception("Failed %s for PN=%s: %s", action, pn, e)

        logger.info("Done: created=%d rescheduled=%d cancelled=%d errors=%d", created, rescheduled, cancelled, errors)
        return 0 if errors == 0 else 1
    except FileNotFoundError as e:
        logger.error("File not found: %s", e)
        return 1
    except Exception as e:
        logger.exception("Job failed: %s", e)
        return 1


if __name__ == "__main__":
    sys.exit(main())
