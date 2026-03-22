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
from datetime import datetime
from typing import Any

from event_id_store import get_invite_state, remove_event, set_invite_state
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


def _build_confirmation_html(
    *,
    start_dt: datetime,
    loc_address: str,
    duration_minutes: int,
    website_url: str,
    phone: str,
    preview_text: str,
    logo_url: str,
    action_kind: str,
) -> str:
    display_when = _human_readable_datetime(start_dt)
    duration_note = (
        "This visit is scheduled for 30 minutes."
        if duration_minutes == 30
        else "This visit is scheduled for 60 minutes."
    )

    if logo_url:
        safe_logo = html.escape(logo_url, quote=True)
        logo_block = (
            f'<img src="{safe_logo}" alt="Liberty PT &amp; Wellness" '
            'style="max-height:56px;max-width:280px;display:block;margin:0 auto 8px auto;" />'
        )
    else:
        logo_block = (
            '<div style="font-size:22px;font-weight:700;color:#ffffff;letter-spacing:-0.02em;text-shadow:0 1px 2px rgba(0,0,0,0.15);">'
            "Liberty PT &amp; Wellness</div>"
            '<div style="font-size:12px;color:rgba(255,255,255,0.9);margin-top:6px;">Physical Therapy</div>'
        )

    safe_preview = html.escape(preview_text)
    preheader = (
        f'<div style="display:none;font-size:1px;color:#fefefe;line-height:1px;max-height:0;max-width:0;opacity:0;overflow:hidden;">'
        f"{safe_preview}</div>"
    )

    action_kind = (action_kind or "create").lower()
    if action_kind == "reschedule":
        intro_line = "Your appointment at <strong>Liberty PT &amp; Wellness</strong> has been updated."
    else:
        intro_line = "This email is to confirm your appointment at <strong>Liberty PT &amp; Wellness</strong>."

    return f"""{preheader}
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f4f7f9;font-family:Segoe UI,Roboto,Helvetica,Arial,sans-serif;">
  <tr><td align="center" style="padding:24px 16px;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(26,77,110,0.08);">
      <tr>
        <td style="background:linear-gradient(135deg,#1a4d6e 0%,#2d6a8f 100%);padding:28px 32px;text-align:center;color:#ffffff;">
          {logo_block}
        </td>
      </tr>
      <tr>
        <td style="padding:32px 36px 24px 36px;color:#333333;font-size:16px;line-height:1.6;">
          <p style="margin:0 0 16px 0;">{intro_line}</p>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#f0f6fa;border-radius:8px;margin:20px 0;">
            <tr>
              <td style="padding:20px 24px;">
                <p style="margin:0 0 8px 0;font-size:14px;color:#5a7a8c;text-transform:uppercase;letter-spacing:0.06em;">Appointment time</p>
                <p style="margin:0;font-size:20px;font-weight:600;color:#1a4d6e;">{display_when}</p>
                <p style="margin:12px 0 0 0;font-size:14px;color:#5a7a8c;text-transform:uppercase;letter-spacing:0.06em;">Location</p>
                <p style="margin:4px 0 0 0;font-size:17px;font-weight:600;color:#333;">{html.escape(str(loc_address))}</p>
              </td>
            </tr>
          </table>
          <p style="margin:0 0 20px 0;font-size:14px;color:#555;">{duration_note}</p>
          <p style="margin:0 0 20px 0;font-size:14px;color:#555;">Please use the attached <strong>invite.ics</strong> file to add this appointment to your calendar.</p>
          <p style="margin:0 0 20px 0;">If you can&rsquo;t make it, please let us know <strong>at least 48 hours in advance</strong> so we can offer that time to another patient on our waiting list. You can reply to this email or call us directly at <a href="tel:+12013661115" style="color:#1a4d6e;font-weight:600;">{phone}</a>.</p>
          <p style="margin:0;text-align:center;padding-top:8px;">
            <a href="{html.escape(website_url, quote=True)}" style="display:inline-block;background:#1a4d6e;color:#ffffff;text-decoration:none;padding:12px 28px;border-radius:8px;font-weight:600;font-size:15px;">{html.escape(website_url.replace("https://", "").replace("http://", "").rstrip("/"))}</a>
          </p>
        </td>
      </tr>
      <tr>
        <td style="padding:20px 36px;background:#f8fafb;border-top:1px solid #e8eef2;text-align:center;font-size:12px;color:#8899a6;">
          <strong>Liberty PT &amp; Wellness</strong><br/>
          <a href="{html.escape(website_url, quote=True)}" style="color:#5a7a8c;">{html.escape(website_url)}</a>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
"""


def _build_cancel_html(*, display_when: str, website_url: str, phone: str) -> str:
    return f"""<html><body style="font-family:Segoe UI,Roboto,Helvetica,Arial,sans-serif;font-size:16px;line-height:1.6;color:#333;">
<p>Your appointment at <strong>Liberty PT &amp; Wellness</strong> scheduled for <strong>{html.escape(display_when)}</strong> has been <strong>cancelled</strong>.</p>
<p>We attached <strong>cancel.ics</strong> so your calendar can remove or update the event if it was already added.</p>
<p>If you have questions, reply to this email or call <a href="tel:+12013661115">{html.escape(phone)}</a>.</p>
<p><a href="{html.escape(website_url, quote=True)}">{html.escape(website_url)}</a></p>
</body></html>"""


def _appointment_key(record: dict[str, Any]) -> str:
    key = record.get("appointment_id") or record.get("AppointmentID")
    if key and str(key).strip():
        return str(key).strip()
    pn = record.get("pn") or record.get("PN") or record.get("PatientNumber") or ""
    date_part = record.get("original_appt_date") or record.get("appt_date") or record.get("ApptDate") or record.get("Date") or ""
    time_part = record.get("original_appt_time") or record.get("appt_time") or record.get("ApptTime") or record.get("Time") or ""
    return f"{pn}_{date_part}_{time_part}".replace(" ", "_")


def _build_common_event_params(
    record: dict[str, Any], location_map: dict[str, str], *, action_kind: str
) -> dict[str, Any]:
    from config import (
        APPOINTMENT_TYPE_DURATION_MINUTES,
        CONFIRMATION_PHONE,
        DEFAULT_DURATION_MINUTES,
        EMAIL_LOGO_URL,
        EMAIL_PREVIEW_TEXT,
        MT30_DURATION_MINUTES,
        REMINDER_MINUTES_BEFORE,
        TIMEZONE,
        WEBSITE_URL,
    )

    date_str = record.get("appt_date") or record.get("ApptDate") or record.get("Date") or ""
    time_str = record.get("appt_time") or record.get("ApptTime") or record.get("Time") or ""
    loc_code = (record.get("location") or record.get("Location") or "LIB").strip().upper()
    loc_address = record.get("location_address") or location_map.get(loc_code) or loc_code
    appt_type = (record.get("appt_type") or record.get("ApptType") or record.get("Type") or "").strip().upper()
    duration = APPOINTMENT_TYPE_DURATION_MINUTES.get(
        appt_type,
        MT30_DURATION_MINUTES if appt_type == "MT30" else DEFAULT_DURATION_MINUTES,
    )

    try:
        if isinstance(date_str, datetime):
            start_dt = date_str
        else:
            start_dt = datetime.strptime(f"{date_str} {time_str}".strip(), "%Y-%m-%d %H:%M:%S")
    except Exception:
        try:
            start_dt = datetime.strptime(f"{date_str} {time_str}".strip(), "%m/%d/%Y %H:%M:%S")
        except Exception:
            try:
                start_dt = datetime.strptime(f"{date_str} {time_str}".strip(), "%m/%d/%Y %H:%M")
            except Exception:
                start_dt = datetime.now()

    from datetime import timedelta

    end_dt = start_dt + timedelta(minutes=duration)
    display_when = _human_readable_datetime(start_dt)
    if action_kind.lower() == "reschedule":
        subject = f"Appointment Updated - {display_when}"
    else:
        subject = f"Appointment Confirmation - {display_when}"

    body = _build_confirmation_html(
        start_dt=start_dt,
        loc_address=loc_address,
        duration_minutes=duration,
        website_url=WEBSITE_URL.rstrip("/") + "/",
        phone=CONFIRMATION_PHONE,
        preview_text=EMAIL_PREVIEW_TEXT,
        logo_url=EMAIL_LOGO_URL,
        action_kind=action_kind,
    )
    reminder_minutes = [int(v) for v in REMINDER_MINUTES_BEFORE if int(v) > 0] if REMINDER_MINUTES_BEFORE else [48 * 60]
    reminder_min = reminder_minutes[0]

    return {
        "subject": subject,
        "start": start_dt,
        "end": end_dt,
        "location_display_name": loc_address,
        "attendee_email": record.get("email") or record.get("Email") or "",
        "attendee_name": record.get("patient_name") or record.get("PatientName") or None,
        "body_content": body,
        "timezone": TIMEZONE,
        "reminder_minutes_before_start": reminder_min,
        "reminder_minutes_before_list": reminder_minutes,
        "display_when": display_when,
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

    send_mail_with_ics(
        access_token,
        to_address=to_addr,
        to_name=params.get("attendee_name"),
        subject=subject_final,
        html_body=html_override or params["body_content"],
        ics_bytes=ics_bytes,
        ics_filename=ics_filename,
        calendar_method=method.upper(),
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
    set_invite_state(key, uid, seq)
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
    set_invite_state(key, uid, seq)
    return True


def do_cancel(
    access_token: str,
    record: dict[str, Any],
    comment: str | None = None,
    *,
    log_action: str = "cancel",
) -> bool:
    _ = comment
    key = _appointment_key(record)
    state = get_invite_state(key)
    if not state:
        logger.warning("Cancel: no stored invite UID for key=%s; skipping", key)
        return False

    from config import CONFIRMATION_PHONE, LOCATION_MAP, WEBSITE_URL

    params = _build_common_event_params(record, LOCATION_MAP, action_kind="cancel")
    if not params["attendee_email"]:
        logger.warning("Cancel: no attendee email for key=%s; removing store only", key)
        remove_event(key)
        return False

    uid = state["ical_uid"]
    seq = state["sequence"] + 1
    subj = f"Appointment cancelled - {params['display_when']}"
    html_body = _build_cancel_html(
        display_when=params["display_when"],
        website_url=WEBSITE_URL.rstrip("/") + "/",
        phone=CONFIRMATION_PHONE,
    )

    _send_ics_mail(
        access_token,
        params=params,
        ical_uid=uid,
        sequence=seq,
        method="CANCEL",
        ics_filename="cancel.ics",
        subject_override=subj,
        html_override=html_body,
        invite_log_record=record,
        invite_log_key=key,
        invite_log_action=log_action,
    )
    remove_event(key)
    return True


def do_delete(access_token: str, record: dict[str, Any]) -> bool:
    return do_cancel(access_token, record, comment=None, log_action="delete")
