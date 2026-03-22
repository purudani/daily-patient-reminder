"""
Build iCalendar (.ics) payloads for email attachments.

Uses METHOD:REQUEST for new/updated appointments and METHOD:CANCEL for removals.
Stable UID + incrementing SEQUENCE lets clients update the same calendar entry.
"""
from __future__ import annotations

import hashlib
from datetime import datetime, timezone
from typing import Final
from zoneinfo import ZoneInfo

# Map Windows / Graph-style names to IANA (for naive datetimes from the sheet)
_TIMEZONE_ALIASES: Final[dict[str, str]] = {
    "Eastern Standard Time": "America/New_York",
    "US Eastern Standard Time": "America/New_York",
    "America/New_York": "America/New_York",
}


def resolve_iana_tz(graph_or_iana_name: str) -> str:
    """Return IANA zone id for config TIMEZONE value."""
    key = (graph_or_iana_name or "").strip()
    return _TIMEZONE_ALIASES.get(key, key or "America/New_York")


def stable_ical_uid(appointment_key: str, domain: str = "libertyptnj.com") -> str:
    """Deterministic UID so the same appointment always maps to the same VEVENT."""
    digest = hashlib.sha256(appointment_key.encode("utf-8")).hexdigest()[:40]
    return f"liberty-appt-{digest}@{domain}"


def _cn_param_value(name: str) -> str:
    """Format CN= parameter (quote if needed)."""
    if not name:
        return '""'
    if any(c in name for c in ";,\\\"&"):
        esc = name.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{esc}"'
    return name


def _escape_ics_text(s: str) -> str:
    """Escape for TEXT values in ICS (RFC 5545)."""
    return (
        s.replace("\\", "\\\\")
        .replace(";", r"\;")
        .replace(",", r"\,")
        .replace("\r\n", "\n")
        .replace("\n", r"\n")
    )


def _fold_line(line: str, max_len: int = 75) -> list[str]:
    """Fold long content lines (octet-oriented; ASCII OK for our content)."""
    if len(line) <= max_len:
        return [line]
    out: list[str] = []
    pos = 0
    first = True
    while pos < len(line):
        chunk_len = max_len if first else max_len - 1  # continuation prefixed with space
        chunk = line[pos : pos + chunk_len]
        if first:
            out.append(chunk)
            first = False
        else:
            out.append(" " + chunk)
        pos += chunk_len
    return out


def _format_utc(dt: datetime) -> str:
    """UTC form: YYYYMMDDTHHMMSSZ."""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    else:
        dt = dt.astimezone(timezone.utc)
    return dt.strftime("%Y%m%dT%H%M%SZ")


def local_naive_to_utc(naive: datetime, tz_name: str) -> datetime:
    """Interpret naive datetime as local in tz_name, return UTC aware."""
    iana = resolve_iana_tz(tz_name)
    local = naive.replace(tzinfo=ZoneInfo(iana))
    return local.astimezone(timezone.utc)


def build_ics_calendar(
    *,
    method: str,
    uid: str,
    sequence: int,
    dtstart_utc: datetime,
    dtend_utc: datetime,
    summary: str,
    description_plain: str,
    location: str,
    organizer_email: str,
    organizer_cn: str,
    attendee_email: str,
    attendee_cn: str,
    reminder_minutes_before: int = 48 * 60,
    reminder_minutes_before_list: list[int] | None = None,
    status: str | None = None,
) -> bytes:
    """
    Full VCALENDAR with one VEVENT. method is REQUEST or CANCEL.
    If status is None and method is CANCEL, STATUS:CANCELLED is added.
    """
    method = method.upper()
    lines: list[str] = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Liberty PT//Appointment Reminder//EN",
        "CALSCALE:GREGORIAN",
        f"METHOD:{method}",
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"SEQUENCE:{sequence}",
        f"DTSTAMP:{_format_utc(datetime.now(timezone.utc))}",
        f"DTSTART:{_format_utc(dtstart_utc)}",
        f"DTEND:{_format_utc(dtend_utc)}",
        f"SUMMARY:{_escape_ics_text(summary)}",
    ]
    if description_plain.strip():
        lines.append(f"DESCRIPTION:{_escape_ics_text(description_plain.strip())}")
    if location.strip():
        lines.append(f"LOCATION:{_escape_ics_text(location.strip())}")
    org_cn = _cn_param_value(organizer_cn)
    att_cn = _cn_param_value(attendee_cn or attendee_email)
    lines.append(f"ORGANIZER;CN={org_cn}:mailto:{organizer_email}")
    lines.append(f"ATTENDEE;CN={att_cn};RSVP=TRUE:mailto:{attendee_email}")
    if method == "CANCEL" or (status and status.upper() == "CANCELLED"):
        lines.append("STATUS:CANCELLED")
    if method == "REQUEST":
        alarm_values: list[int] = []
        if reminder_minutes_before_list:
            for v in reminder_minutes_before_list:
                try:
                    n = int(v)
                except Exception:
                    continue
                if n > 0 and n not in alarm_values:
                    alarm_values.append(n)
        else:
            try:
                n = int(reminder_minutes_before)
                if n > 0:
                    alarm_values.append(n)
            except Exception:
                pass

        for mins in alarm_values:
            lines.extend(
                [
                    "BEGIN:VALARM",
                    f"TRIGGER:-PT{mins}M",
                    "ACTION:DISPLAY",
                    "DESCRIPTION:Appointment reminder",
                    "END:VALARM",
                ]
            )
    lines.extend(["END:VEVENT", "END:VCALENDAR"])

    folded: list[str] = []
    for raw in lines:
        for part in _fold_line(raw):
            folded.append(part)
    text = "\r\n".join(folded) + "\r\n"
    return text.encode("utf-8")
