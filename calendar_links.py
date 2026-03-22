"""
Universal “add to calendar” links (same email for every recipient).

- Google Calendar: official TEMPLATE URL (works in browser / Gmail).
- Outlook on the web: Microsoft calendar compose deeplink (M365 / many work accounts).

Apple Calendar and others typically use the attached invite.ics file.
"""
from __future__ import annotations

from datetime import datetime, timezone
from urllib.parse import quote

from ics_calendar import local_naive_to_utc


def _utc(start_naive: datetime, end_naive: datetime, tz_name: str) -> tuple[datetime, datetime]:
    s = local_naive_to_utc(start_naive, tz_name)
    e = local_naive_to_utc(end_naive, tz_name)
    if s.tzinfo is None:
        s = s.replace(tzinfo=timezone.utc)
    else:
        s = s.astimezone(timezone.utc)
    if e.tzinfo is None:
        e = e.replace(tzinfo=timezone.utc)
    else:
        e = e.astimezone(timezone.utc)
    return s, e


def google_calendar_template_url(
    *,
    title: str,
    start_naive: datetime,
    end_naive: datetime,
    tz_name: str,
    details: str,
    location: str,
) -> str:
    """https://calendar.google.com/calendar/render?action=TEMPLATE&..."""
    s_utc, e_utc = _utc(start_naive, end_naive, tz_name)
    dates = f"{s_utc.strftime('%Y%m%dT%H%M%SZ')}/{e_utc.strftime('%Y%m%dT%H%M%SZ')}"
    base = "https://calendar.google.com/calendar/render"
    return (
        f"{base}?action=TEMPLATE"
        f"&text={quote(title)}"
        f"&dates={dates}"
        f"&details={quote(details)}"
        f"&location={quote(location)}"
    )


def outlook_web_compose_url(
    *,
    title: str,
    start_naive: datetime,
    end_naive: datetime,
    tz_name: str,
    body: str,
    location: str,
    base_url: str = "https://outlook.office.com",
) -> str:
    """
    Outlook calendar compose deeplink (Office 365).

    For personal Outlook.com accounts, set base_url to https://outlook.live.com
    via config OUTLOOK_CALENDAR_WEB_BASE.
    """
    s_utc, e_utc = _utc(start_naive, end_naive, tz_name)
    # Outlook accepts ISO-like startdt/enddt; UTC with Z is widely accepted.
    start_iso = s_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_iso = e_utc.strftime("%Y-%m-%dT%H:%M:%SZ")
    root = base_url.rstrip("/")
    path = "/calendar/0/deeplink/compose"
    return (
        f"{root}{path}"
        f"?subject={quote(title)}"
        f"&startdt={quote(start_iso)}"
        f"&enddt={quote(end_iso)}"
        f"&location={quote(location)}"
        f"&body={quote(body)}"
    )
