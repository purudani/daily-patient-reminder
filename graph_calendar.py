"""
Microsoft Graph calendar API: create, update, cancel, delete events.

Optional / not used by the daily patient reminder (patient flow is sendMail + .ics + web links).
Uses /users/{GRAPH_MAILBOX_USER}/events when GRAPH_MAILBOX_USER is set, else /me/events.
"""
from __future__ import annotations

import logging
from datetime import datetime
from typing import Any
from urllib.parse import quote

import requests

from config import GRAPH_MAILBOX_USER

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _calendar_root() -> str:
    """Base path for calendar: /users/{upn} or /me."""
    u = (GRAPH_MAILBOX_USER or "").strip()
    if u:
        return f"{GRAPH_BASE}/users/{quote(u, safe=':@')}"
    return f"{GRAPH_BASE}/me"


def _headers(access_token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }


def _format_datetime(dt: datetime, timezone: str) -> dict[str, str]:
    """Graph dateTimeTimeZone: { "dateTime": "YYYY-MM-DDTHH:mm:ss", "timeZone": "..." }."""
    return {
        "dateTime": dt.strftime("%Y-%m-%dT%H:%M:%S"),
        "timeZone": timezone,
    }


def build_event_payload(
    *,
    subject: str,
    start: datetime,
    end: datetime,
    location_display_name: str,
    attendee_email: str,
    attendee_name: str | None = None,
    body_content: str | None = None,
    timezone: str = "Eastern Standard Time",
    reminder_minutes_before_start: int = 48 * 60,
    response_requested: bool = True,
) -> dict[str, Any]:
    """Build the JSON body for create/update event."""
    payload: dict[str, Any] = {
        "subject": subject,
        "start": _format_datetime(start, timezone),
        "end": _format_datetime(end, timezone),
        "location": {"displayName": location_display_name},
        "attendees": [
            {
                "emailAddress": {
                    "address": attendee_email,
                    "name": attendee_name or attendee_email,
                },
                "type": "required",
            }
        ],
        "responseRequested": response_requested,
        "isReminderOn": True,
        "reminderMinutesBeforeStart": reminder_minutes_before_start,
    }
    if body_content:
        payload["body"] = {"contentType": "HTML", "content": body_content}
    return payload


def create_event(
    access_token: str,
    subject: str,
    start: datetime,
    end: datetime,
    location_display_name: str,
    attendee_email: str,
    attendee_name: str | None = None,
    body_content: str | None = None,
    timezone: str = "Eastern Standard Time",
    reminder_minutes_before_start: int = 48 * 60,
) -> dict[str, str]:
    """
    POST .../events on the configured mailbox. Returns event_id and iCalUId for storage.
    """
    payload = build_event_payload(
        subject=subject,
        start=start,
        end=end,
        location_display_name=location_display_name,
        attendee_email=attendee_email,
        attendee_name=attendee_name,
        body_content=body_content,
        timezone=timezone,
        reminder_minutes_before_start=reminder_minutes_before_start,
    )
    url = f"{_calendar_root()}/events"
    resp = requests.post(url, json=payload, headers=_headers(access_token), timeout=30)
    resp.raise_for_status()
    data = resp.json()
    event_id = data.get("id") or ""
    i_cal_uid = data.get("iCalUId") or ""
    logger.info("Created event id=%s", event_id)
    return {"event_id": event_id, "i_cal_uid": i_cal_uid}


def update_event(
    access_token: str,
    event_id: str,
    *,
    subject: str | None = None,
    start: datetime | None = None,
    end: datetime | None = None,
    location_display_name: str | None = None,
    attendee_email: str | None = None,
    attendee_name: str | None = None,
    body_content: str | None = None,
    timezone: str = "Eastern Standard Time",
    reminder_minutes_before_start: int | None = None,
) -> None:
    """PATCH .../events/{id} on the configured mailbox."""
    payload: dict[str, Any] = {}
    if subject is not None:
        payload["subject"] = subject
    if start is not None:
        payload["start"] = _format_datetime(start, timezone)
    if end is not None:
        payload["end"] = _format_datetime(end, timezone)
    if location_display_name is not None:
        payload["location"] = {"displayName": location_display_name}
    if attendee_email is not None:
        payload["attendees"] = [
            {
                "emailAddress": {
                    "address": attendee_email,
                    "name": attendee_name or attendee_email,
                },
                "type": "required",
            }
        ]
    if body_content is not None:
        payload["body"] = {"contentType": "HTML", "content": body_content}
    if reminder_minutes_before_start is not None:
        payload["isReminderOn"] = True
        payload["reminderMinutesBeforeStart"] = reminder_minutes_before_start

    if not payload:
        logger.warning("update_event called with no changes for id=%s", event_id)
        return

    url = f"{_calendar_root()}/events/{event_id}"
    resp = requests.patch(url, json=payload, headers=_headers(access_token), timeout=30)
    resp.raise_for_status()
    logger.info("Updated event id=%s", event_id)


def cancel_event(access_token: str, event_id: str, comment: str | None = None) -> None:
    """POST .../events/{id}/cancel on the configured mailbox."""
    url = f"{_calendar_root()}/events/{event_id}/cancel"
    body = {} if not comment else {"Comment": comment}
    resp = requests.post(url, json=body, headers=_headers(access_token), timeout=30)
    resp.raise_for_status()
    logger.info("Cancelled event id=%s", event_id)


def delete_event(access_token: str, event_id: str) -> None:
    """DELETE .../events/{id} on the configured mailbox."""
    url = f"{_calendar_root()}/events/{event_id}"
    resp = requests.delete(url, headers=_headers(access_token), timeout=30)
    resp.raise_for_status()
    logger.info("Deleted event id=%s", event_id)
