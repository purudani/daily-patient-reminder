"""
Patient notifications: one email per action, works across Gmail, Outlook, Apple, etc.

Strategy (industry-standard, no per-vendor branching):
- HTML confirmation (branded body).
- Attached invite.ics / cancel.ics with stable UID + SEQUENCE (RFC 5545 updates).
- Body points patients to the .ics attachment for their calendar app.

Uses Microsoft Graph sendMail only (Mail.Send). No Graph Calendar API for patients.
"""
from __future__ import annotations

import html
import logging
from datetime import datetime, timedelta
from typing import Any

from event_id_store import (
    get_all_keys,
    get_invite_state,
    remove_event,
    resolve_store_key_for_cancel,
    set_invite_state,
)
from graph_mail import send_mail_with_ics
from graph_user import resolve_organizer_email
from ics_calendar import build_ics_calendar, local_naive_to_utc, stable_ical_uid

logger = logging.getLogger(__name__)


def _human_readable_datetime(dt: datetime) -> str:
    date_part = dt.strftime("%B ") + str(dt.day) + dt.strftime(", %Y")
    t = dt.strftime("%I:%M %p").lstrip("0")
    if t.startswith(":"):
        t = "12" + t
    return f"{date_part} at {t}"


def _first_name_from_record(record: dict[str, Any]) -> str:
    fn = (record.get("first_name") or "").strip()
    if fn:
        return fn
    name = (record.get("patient_name") or "").strip()
    if not name or "@" in name:
        return "there"
    if "," in name:
        after = name.split(",", 1)[1].strip()
        return (after.split() or ["there"])[0]
    return (name.split() or ["there"])[0]


def _parse_record_datetime(date_str: Any, time_str: Any) -> datetime | None:
    if date_str is None or time_str is None:
        return None
    if isinstance(date_str, datetime):
        return date_str
    ds = str(date_str).strip()
    ts = str(time_str).strip()
    if not ds or not ts:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M"):
        try:
            return datetime.strptime(f"{ds} {ts}".strip(), fmt)
        except ValueError:
            continue
    return None


def _subject_and_preview(action_kind: str, display_when: str) -> tuple[str, str]:
    k = (action_kind or "create").lower()
    org = "Liberty PT & Wellness"
    if k == "reschedule":
        return (
            f"Appointment Changed - {display_when}",
            f"Appointment Rescheduled to {display_when} at {org}",
        )
    if k == "cancel":
        return (
            f"Appointment Canceled - {display_when}",
            f"Appointment Canceled {display_when} at {org}",
        )
    return (
        f"New Appointment - {display_when}",
        f"New Appointment scheduled {display_when} at {org}",
    )


def _build_confirmation_html(
    *,
    record: dict[str, Any],
    start_dt: datetime,
    loc_address: str,
    preview_text: str,
    logo_url: str,
    action_kind: str,
    previous_display_when: str | None,
    calendar_button_label: str,
    ics_filename: str,
    ics_cid: str,
    phone_display: str,
    phone_href: str,
) -> str:
    display_when = _human_readable_datetime(start_dt)
    fn = html.escape(_first_name_from_record(record))

    if logo_url:
        safe_logo = html.escape(logo_url, quote=True)
        logo_block = (
            f'<img src="{safe_logo}" alt="Liberty PT &amp; Wellness" '
            'style="max-height:56px;max-width:280px;display:block;margin:0 auto 8px auto;" />'
        )
    else:
        logo_block = (
            '<div class="lpw-brand" style="font-size:22px;font-weight:700;color:#ffffff;letter-spacing:-0.02em;text-shadow:0 1px 2px rgba(0,0,0,0.15);">'
            "Liberty PT &amp; Wellness</div>"
            '<div style="font-size:12px;color:rgba(255,255,255,0.9);margin-top:6px;">Physical Therapy</div>'
        )

    safe_preview = html.escape(preview_text)
    preheader = (
        '<div style="display:none;font-size:1px;color:#fefefe;line-height:1px;max-height:0;max-width:0;'
        f'opacity:0;overflow:hidden;">{safe_preview}</div>'
    )

    action_kind = (action_kind or "create").lower()
    if action_kind == "reschedule":
        intro_line = (
            f"Dear {fn},<br/><br/>"
            "Your appointment at <strong>Liberty PT &amp; Wellness</strong> has been updated."
        )
    else:
        intro_line = (
            f"Dear {fn},<br/><br/>"
            "This email is to confirm your appointment at <strong>Liberty PT &amp; Wellness</strong>."
        )

    if previous_display_when:
        time_block = f"""
                <p style="margin:0 0 8px 0;font-size:15px;color:#5a7a8c;text-transform:uppercase;letter-spacing:0.06em;">Appointment time</p>
                <p class="lpw-when" style="margin:0 0 10px 0;font-size:22px;line-height:1.35;">
                  <span style="text-decoration:line-through;color:#7a8a96;">{html.escape(previous_display_when)}</span><br/>
                  <span style="font-weight:700;color:#1a4d6e;">{html.escape(display_when)}</span>
                </p>"""
    else:
        time_block = f"""
                <p style="margin:0 0 8px 0;font-size:15px;color:#5a7a8c;text-transform:uppercase;letter-spacing:0.06em;">Appointment time</p>
                <p class="lpw-when" style="margin:0;font-size:22px;font-weight:700;color:#1a4d6e;line-height:1.35;">{html.escape(display_when)}</p>"""

    cid_href = html.escape(f"cid:{ics_cid}", quote=True)
    safe_phone_d = html.escape(phone_display)
    safe_fn = html.escape(ics_filename)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style type="text/css">
@media only screen and (max-width: 600px) {{
  .lpw-body {{ font-size: 19px !important; line-height: 1.65 !important; }}
  .lpw-when {{ font-size: 24px !important; }}
  .lpw-btn {{ font-size: 19px !important; padding: 18px 28px !important; }}
  .lpw-brand {{ font-size: 24px !important; }}
}}
</style>
</head>
<body style="margin:0;padding:0;background:#f4f7f9;">
{preheader}
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f4f7f9;font-family:Segoe UI,Roboto,Helvetica,Arial,sans-serif;">
  <tr><td align="center" style="padding:20px 12px;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(26,77,110,0.08);">
      <tr>
        <td style="background:linear-gradient(135deg,#1a4d6e 0%,#2d6a8f 100%);padding:28px 24px;text-align:center;color:#ffffff;">
          {logo_block}
        </td>
      </tr>
      <tr>
        <td class="lpw-body" style="padding:28px 24px 24px 24px;color:#222222;font-size:18px;line-height:1.65;">
          <p style="margin:0 0 18px 0;">{intro_line}</p>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#f0f6fa;border-radius:8px;margin:20px 0;">
            <tr>
              <td style="padding:22px 22px;">
                {time_block}
                <p style="margin:16px 0 0 0;font-size:15px;color:#5a7a8c;text-transform:uppercase;letter-spacing:0.06em;">Location</p>
                <p style="margin:6px 0 0 0;font-size:19px;font-weight:600;color:#333;">{html.escape(str(loc_address))}</p>
              </td>
            </tr>
          </table>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:24px 0 8px 0;">
            <tr>
              <td align="center">
                <a class="lpw-btn" href="{cid_href}" style="display:inline-block;background:#1a4d6e;color:#ffffff;text-decoration:none;padding:16px 32px;border-radius:10px;font-weight:700;font-size:18px;">{html.escape(calendar_button_label)}</a>
              </td>
            </tr>
          </table>
          <p style="margin:16px 0 20px 0;font-size:16px;color:#444;">Open the attached <strong>{safe_fn}</strong> file to add this appointment to your calendar (many phones: tap the attachment at the bottom of this email).</p>
          <p style="margin:0 0 20px 0;">If you can&rsquo;t make it, please let us know <strong>at least 48 hours in advance</strong> so we can offer that time to another patient on our waiting list. You can reply to this email or call us at <a href="{html.escape(phone_href, quote=True)}" style="color:#1a4d6e;font-weight:600;">{safe_phone_d}</a>.</p>
        </td>
      </tr>
      <tr>
        <td style="padding:20px 24px;background:#f8fafb;border-top:1px solid #e8eef2;text-align:center;font-size:14px;color:#8899a6;">
          <strong>Liberty PT &amp; Wellness</strong>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
</body>
</html>"""


def _build_cancel_html(
    *,
    record: dict[str, Any],
    display_when: str,
    preview_text: str,
    phone_display: str,
    phone_href: str,
    ics_cid: str,
) -> str:
    fn = html.escape(_first_name_from_record(record))
    safe_preview = html.escape(preview_text)
    preheader = (
        '<div style="display:none;font-size:1px;color:#fefefe;line-height:1px;max-height:0;max-width:0;'
        f'opacity:0;overflow:hidden;">{safe_preview}</div>'
    )
    cid_href = html.escape(f"cid:{ics_cid}", quote=True)
    safe_phone_d = html.escape(phone_display)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<style type="text/css">
@media only screen and (max-width: 600px) {{
  .lpw-body {{ font-size: 19px !important; line-height: 1.65 !important; }}
  .lpw-btn {{ font-size: 19px !important; padding: 18px 28px !important; }}
}}
</style>
</head>
<body style="margin:0;padding:0;background:#f4f7f9;">
{preheader}
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f4f7f9;font-family:Segoe UI,Roboto,Helvetica,Arial,sans-serif;">
  <tr><td align="center" style="padding:20px 12px;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(26,77,110,0.08);">
      <tr>
        <td style="background:linear-gradient(135deg,#1a4d6e 0%,#2d6a8f 100%);padding:28px 24px;text-align:center;color:#ffffff;">
          <div style="font-size:22px;font-weight:700;">Liberty PT &amp; Wellness</div>
          <div style="font-size:12px;color:rgba(255,255,255,0.9);margin-top:6px;">Physical Therapy</div>
        </td>
      </tr>
      <tr>
        <td class="lpw-body" style="padding:28px 24px;color:#222222;font-size:18px;line-height:1.65;">
          <p style="margin:0 0 16px 0;">Dear {fn},</p>
          <p style="margin:0 0 16px 0;">Your appointment at <strong>Liberty PT &amp; Wellness</strong> scheduled for <strong>{html.escape(display_when)}</strong> has been <strong>cancelled</strong>.</p>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:8px 0 16px 0;">
            <tr>
              <td align="center">
                <a class="lpw-btn" href="{cid_href}" style="display:inline-block;background:#1a4d6e;color:#ffffff;text-decoration:none;padding:16px 32px;border-radius:10px;font-weight:700;font-size:18px;">Remove from Calendar</a>
              </td>
            </tr>
          </table>
          <p style="margin:0 0 16px 0;">We attached <strong>cancel.ics</strong> so your calendar can remove this event if it was already added (tap the attachment on your phone).</p>
          <p style="margin:0;">If you have questions, reply to this email or call <a href="{html.escape(phone_href, quote=True)}" style="color:#1a4d6e;font-weight:600;">{safe_phone_d}</a>.</p>
        </td>
      </tr>
      <tr>
        <td style="padding:20px 24px;background:#f8fafb;border-top:1px solid #e8eef2;text-align:center;font-size:14px;color:#8899a6;">
          <strong>Liberty PT &amp; Wellness</strong>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
</body>
</html>"""


def _slot_from_store_key(key: str) -> tuple[str, str] | None:
    """
    Recover (appt_date, appt_time) from keys like ``12345_2025-04-01_09:00:00``.
    Used when canceling against legacy store rows that lack last_appt_* metadata.
    """
    try:
        parts = str(key).rsplit("_", 2)
        if len(parts) != 3:
            return None
        _pn, d, t = parts
        if len(d) == 10 and d[4] == "-" and d[7] == "-" and len(t) >= 5:
            return d, t
    except Exception:
        pass
    return None


def _appointment_key(record: dict[str, Any]) -> str:
    """Stable key for ICS UID store: PN + date + time (reschedule uses original_* from Action row)."""
    pn = record.get("pn") or record.get("PN") or record.get("PatientNumber") or ""
    date_part = record.get("original_appt_date") or record.get("appt_date") or record.get("ApptDate") or record.get("Date") or ""
    time_part = record.get("original_appt_time") or record.get("appt_time") or record.get("ApptTime") or record.get("Time") or ""
    return f"{pn}_{date_part}_{time_part}".replace(" ", "_")


def _build_common_event_params(
    record: dict[str, Any],
    location_map: dict[str, str],
    *,
    action_kind: str,
    duration_minutes_override: int | None = None,
) -> dict[str, Any]:
    from config import (
        APPOINTMENT_TYPE_DURATION_MINUTES,
        DEFAULT_DURATION_MINUTES,
        EMAIL_LOGO_URL,
        ICS_INVITE_CONTENT_ID,
        MT30_DURATION_MINUTES,
        REMINDER_MINUTES_BEFORE,
        TIMEZONE,
        phone_for_location_code,
    )

    date_str = record.get("appt_date") or record.get("ApptDate") or record.get("Date") or ""
    time_str = record.get("appt_time") or record.get("ApptTime") or record.get("Time") or ""
    loc_code = (record.get("location") or record.get("Location") or "LIB").strip().upper()
    loc_address = record.get("location_address") or location_map.get(loc_code) or loc_code
    appt_type = (record.get("appt_type") or record.get("ApptType") or record.get("Type") or "").strip().upper()
    if duration_minutes_override is not None:
        duration = int(duration_minutes_override)
    else:
        duration = APPOINTMENT_TYPE_DURATION_MINUTES.get(
            appt_type,
            MT30_DURATION_MINUTES if appt_type == "MT30" else DEFAULT_DURATION_MINUTES,
        )

    if isinstance(date_str, datetime):
        start_dt = date_str
    else:
        parsed = _parse_record_datetime(date_str, time_str)
        start_dt = parsed if parsed is not None else datetime.now()

    end_dt = start_dt + timedelta(minutes=duration)
    display_when = _human_readable_datetime(start_dt)
    subject, preview = _subject_and_preview(action_kind, display_when)

    ak = (action_kind or "create").lower()
    previous_display_when: str | None = None
    if ak == "reschedule":
        prev_dt = _parse_record_datetime(
            record.get("original_appt_date"),
            record.get("original_appt_time"),
        )
        if prev_dt is not None:
            previous_display_when = _human_readable_datetime(prev_dt)

    phone_display, phone_href = phone_for_location_code(loc_code)

    if ak == "cancel":
        body: str = "<html><body></body></html>"
    else:
        cal_label = "Update Calendar" if ak == "reschedule" else "Add to Calendar"
        body = _build_confirmation_html(
            record=record,
            start_dt=start_dt,
            loc_address=loc_address,
            preview_text=preview,
            logo_url=EMAIL_LOGO_URL,
            action_kind=action_kind,
            previous_display_when=previous_display_when,
            calendar_button_label=cal_label,
            ics_filename="invite.ics",
            ics_cid=ICS_INVITE_CONTENT_ID,
            phone_display=phone_display,
            phone_href=phone_href,
        )

    reminder_minutes = (
        [int(v) for v in REMINDER_MINUTES_BEFORE if int(v) > 0]
        if REMINDER_MINUTES_BEFORE
        else [48 * 60]
    )
    reminder_min = reminder_minutes[0]

    return {
        "subject": subject,
        "start": start_dt,
        "end": end_dt,
        "location_display_name": loc_address,
        "location_code": loc_code,
        "attendee_email": record.get("email") or record.get("Email") or "",
        "attendee_name": record.get("patient_name") or record.get("PatientName") or None,
        "body_content": body,
        "timezone": TIMEZONE,
        "reminder_minutes_before_start": reminder_min,
        "reminder_minutes_before_list": reminder_minutes,
        "display_when": display_when,
        "preview_text": preview,
    }


def _send_ics_mail(
    access_token: str,
    *,
    params: dict[str, Any],
    ical_uid: str,
    sequence: int,
    method: str,
    ics_filename: str,
    subject_override: str | None = None,
    html_override: str | None = None,
    invite_log_record: dict[str, Any] | None = None,
    invite_log_key: str | None = None,
    invite_log_action: str | None = None,
) -> None:
    from config import CALENDAR_TITLE, ORGANIZER_NAME

    to_addr = (params.get("attendee_email") or "").strip()
    if not to_addr:
        raise ValueError("No attendee email")

    org_email = resolve_organizer_email(access_token)
    dt_start_utc = local_naive_to_utc(params["start"], params["timezone"])
    dt_end_utc = local_naive_to_utc(params["end"], params["timezone"])

    if method.upper() == "CANCEL":
        desc_plain = "This appointment has been cancelled."
    else:
        desc_plain = (
            f"Liberty PT & Wellness appointment.\n{params['display_when']}\n"
            f"{params['location_display_name']}\n"
            f"Open the attached invite.ics to add this to your calendar."
        )

    ics_bytes = build_ics_calendar(
        method=method.upper(),
        uid=ical_uid,
        sequence=sequence,
        dtstart_utc=dt_start_utc,
        dtend_utc=dt_end_utc,
        summary=CALENDAR_TITLE,
        description_plain=desc_plain,
        location=params["location_display_name"],
        organizer_email=org_email,
        organizer_cn=ORGANIZER_NAME,
        attendee_email=to_addr,
        attendee_cn=params.get("attendee_name") or to_addr,
        reminder_minutes_before=params.get("reminder_minutes_before_start", 48 * 60),
        reminder_minutes_before_list=params.get("reminder_minutes_before_list"),
        status="CANCELLED" if method.upper() == "CANCEL" else None,
    )

    subject_final = subject_override or params["subject"]

    from config import ICS_CANCEL_CONTENT_ID, ICS_INVITE_CONTENT_ID

    ics_cid = ICS_CANCEL_CONTENT_ID if method.upper() == "CANCEL" else ICS_INVITE_CONTENT_ID

    send_mail_with_ics(
        access_token,
        to_address=to_addr,
        to_name=params.get("attendee_name"),
        subject=subject_final,
        html_body=html_override or params["body_content"],
        ics_bytes=ics_bytes,
        ics_filename=ics_filename,
        calendar_method=method.upper(),
        ics_content_id=ics_cid,
    )

    if invite_log_record is not None and invite_log_key and invite_log_action:
        try:
            from config import INVITE_LOG_PATH

            if (INVITE_LOG_PATH or "").strip():
                from invite_log import append_invite_row

                duration_minutes = int(
                    (params["end"] - params["start"]).total_seconds() // 60
                )
                append_invite_row(
                    from_email=org_email,
                    to_email=to_addr,
                    to_name=params.get("attendee_name"),
                    action=invite_log_action,
                    subject=subject_final,
                    record=invite_log_record,
                    appointment_key=invite_log_key,
                    ical_uid=ical_uid,
                    ics_sequence=sequence,
                    ics_method=method,
                    ics_attachment=ics_filename,
                    duration_minutes=duration_minutes,
                )
        except Exception:
            logger.exception(
                "Invite log append failed for key=%s; mail was still sent.",
                invite_log_key,
            )


def _uid_seq_for_reschedule(key: str) -> tuple[str, int]:
    state = get_invite_state(key)
    if state:
        return state["ical_uid"], state["sequence"] + 1
    return stable_ical_uid(key), 0


def do_create(access_token: str, record: dict[str, Any], location_map: dict[str, str]) -> bool:
    key = _appointment_key(record)
    params = _build_common_event_params(record, location_map, action_kind="create")
    if not params["attendee_email"]:
        logger.warning("Skipping create: no attendee email for record %s", record)
        return False

    state = get_invite_state(key)
    if state:
        logger.warning(
            "Create: invite state already exists for key=%s; sending ICS update (SEQUENCE+1)",
            key,
        )
        uid, seq = state["ical_uid"], state["sequence"] + 1
    else:
        uid, seq = stable_ical_uid(key), 0

    _send_ics_mail(
        access_token,
        params=params,
        ical_uid=uid,
        sequence=seq,
        method="REQUEST",
        ics_filename="invite.ics",
        invite_log_record=record,
        invite_log_key=key,
        invite_log_action="create",
    )
    dur_min = int((params["end"] - params["start"]).total_seconds() // 60)
    set_invite_state(
        key,
        uid,
        seq,
        last_appt_date=str(record.get("appt_date") or ""),
        last_appt_time=str(record.get("appt_time") or ""),
        last_duration_minutes=dur_min,
    )
    return True


def do_reschedule(access_token: str, record: dict[str, Any], location_map: dict[str, str]) -> bool:
    key = _appointment_key(record)
    params = _build_common_event_params(record, location_map, action_kind="reschedule")
    if not params["attendee_email"]:
        logger.warning("Skipping reschedule: no attendee email for record %s", record)
        return False

    uid, seq = _uid_seq_for_reschedule(key)
    _send_ics_mail(
        access_token,
        params=params,
        ical_uid=uid,
        sequence=seq,
        method="REQUEST",
        ics_filename="invite.ics",
        invite_log_record=record,
        invite_log_key=key,
        invite_log_action="reschedule",
    )
    dur_min = int((params["end"] - params["start"]).total_seconds() // 60)
    set_invite_state(
        key,
        uid,
        seq,
        last_appt_date=str(record.get("appt_date") or ""),
        last_appt_time=str(record.get("appt_time") or ""),
        last_duration_minutes=dur_min,
    )
    return True


def do_cancel(
    access_token: str,
    record: dict[str, Any],
    comment: str | None = None,
    *,
    log_action: str = "cancel",
) -> bool:
    _ = comment
    row_key = _appointment_key(record)
    pn = str(record.get("pn") or record.get("PN") or "").strip()
    key = resolve_store_key_for_cancel(pn, row_key)
    state = get_invite_state(key) if key else None
    if not key or not state:
        same_pn_keys = [k for k in get_all_keys() if pn and str(k).startswith(f"{pn}_")]
        logger.warning(
            "Cancel: no stored invite UID for row key=%s. Keys in store for PN=%s: %s. "
            "After a reschedule, the store uses the **original** Date/Time; put that on the cancel "
            "row, or keep only one open visit per PN. See event_id_store.json.",
            row_key,
            pn or "?",
            same_pn_keys if same_pn_keys else "none",
        )
        return False
    if key != row_key:
        logger.info(
            "Cancel: using store key=%s (Action row keyed as %s — e.g. current vs original slot)",
            key,
            row_key,
        )

    from config import ICS_CANCEL_CONTENT_ID, LOCATION_MAP, phone_for_location_code

    merged = dict(record)
    if state:
        ld = state.get("last_appt_date")
        lt = state.get("last_appt_time")
        if ld and lt:
            merged["appt_date"] = ld
            merged["appt_time"] = lt
        elif key:
            slot = _slot_from_store_key(key)
            if slot:
                merged["appt_date"], merged["appt_time"] = slot
    dur_ov = state.get("last_duration_minutes") if state else None
    if not isinstance(dur_ov, int):
        dur_ov = None

    params = _build_common_event_params(
        merged,
        LOCATION_MAP,
        action_kind="cancel",
        duration_minutes_override=dur_ov,
    )
    if not params["attendee_email"]:
        logger.warning("Cancel: no attendee email for key=%s; removing store only", key)
        remove_event(key)
        return False

    uid = state["ical_uid"]
    seq = state["sequence"] + 1
    loc_code = (merged.get("location") or record.get("location") or "LIB").strip().upper()
    phone_display, phone_href = phone_for_location_code(loc_code)
    html_body = _build_cancel_html(
        record=merged,
        display_when=params["display_when"],
        preview_text=params["preview_text"],
        phone_display=phone_display,
        phone_href=phone_href,
        ics_cid=ICS_CANCEL_CONTENT_ID,
    )

    _send_ics_mail(
        access_token,
        params=params,
        ical_uid=uid,
        sequence=seq,
        method="CANCEL",
        ics_filename="cancel.ics",
        html_override=html_body,
        invite_log_record=record,
        invite_log_key=key,
        invite_log_action=log_action,
    )
    remove_event(key)
    return True


def do_delete(access_token: str, record: dict[str, Any]) -> bool:
    return do_cancel(access_token, record, comment=None, log_action="delete")
