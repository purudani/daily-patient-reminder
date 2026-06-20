"""
Send email via Microsoft Graph with a text/calendar (.ics) attachment.

Uses /users/{GRAPH_MAILBOX_USER}/sendMail when GRAPH_MAILBOX_USER is set, else /me/sendMail.
"""
from __future__ import annotations

import base64
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
import logging
import time
from typing import Any
from urllib.parse import quote
from uuid import uuid4

import requests

from config import (
    EMAIL_REPLY_TO,
    GRAPH_MAILBOX_USER,
    GRAPH_MAIL_MAX_ATTEMPTS,
    GRAPH_MAIL_RETRY_BACKOFF_SECONDS,
    GRAPH_MAIL_TIMEOUT_SECONDS,
)

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TRANSIENT_SENDMAIL_STATUS_CODES = {429, 500, 502, 503, 504}
TRANSIENT_SENDMAIL_EXCEPTIONS = (
    requests.ConnectionError,
    requests.Timeout,
)


def _user_root() -> str:
    u = (GRAPH_MAILBOX_USER or "").strip()
    if u:
        return f"{GRAPH_BASE}/users/{quote(u, safe=':@')}"
    return f"{GRAPH_BASE}/me"


def _headers(access_token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }


def _retry_after_seconds(value: str | None) -> float | None:
    raw = (value or "").strip()
    if not raw:
        return None
    try:
        return max(0.0, float(raw))
    except ValueError:
        pass
    try:
        retry_at = parsedate_to_datetime(raw)
    except (TypeError, ValueError):
        return None
    if retry_at.tzinfo is None:
        retry_at = retry_at.replace(tzinfo=timezone.utc)
    return max(0.0, (retry_at - datetime.now(timezone.utc)).total_seconds())


def _retry_delay_seconds(resp: requests.Response | None, attempt: int) -> float:
    if resp is not None:
        retry_after = _retry_after_seconds(resp.headers.get("Retry-After"))
        if retry_after is not None:
            return retry_after
    return GRAPH_MAIL_RETRY_BACKOFF_SECONDS * (2 ** max(0, attempt - 1))


def _post_send_mail(
    access_token: str,
    payload: dict[str, Any],
    *,
    description: str,
) -> None:
    url = f"{_user_root()}/sendMail"
    max_attempts = max(1, int(GRAPH_MAIL_MAX_ATTEMPTS))
    request_id = str(uuid4())

    last_error: BaseException | None = None
    for attempt in range(1, max_attempts + 1):
        headers = _headers(access_token)
        headers["client-request-id"] = request_id
        headers["return-client-request-id"] = "true"
        try:
            resp = requests.post(
                url,
                json=payload,
                headers=headers,
                timeout=GRAPH_MAIL_TIMEOUT_SECONDS,
            )
            if resp.status_code not in TRANSIENT_SENDMAIL_STATUS_CODES:
                resp.raise_for_status()
                if attempt > 1:
                    logger.info(
                        "Graph sendMail succeeded for %s on attempt %d/%d",
                        description,
                        attempt,
                        max_attempts,
                    )
                return

            try:
                resp.raise_for_status()
            except requests.HTTPError as exc:
                last_error = exc
            else:
                return

            if attempt >= max_attempts:
                break
            delay = _retry_delay_seconds(resp, attempt)
            logger.warning(
                "Graph sendMail got HTTP %s for %s on attempt %d/%d; "
                "retrying in %.1f seconds (client_request_id=%s)",
                resp.status_code,
                description,
                attempt,
                max_attempts,
                delay,
                request_id,
            )
            time.sleep(delay)
        except TRANSIENT_SENDMAIL_EXCEPTIONS as exc:
            last_error = exc
            if attempt >= max_attempts:
                break
            delay = _retry_delay_seconds(None, attempt)
            logger.warning(
                "Graph sendMail had a transient network error for %s on attempt %d/%d: %s; "
                "retrying in %.1f seconds (client_request_id=%s)",
                description,
                attempt,
                max_attempts,
                exc,
                delay,
                request_id,
            )
            time.sleep(delay)

    logger.error(
        "Graph sendMail failed for %s after %d attempt(s) (client_request_id=%s)",
        description,
        max_attempts,
        request_id,
    )
    if last_error is not None:
        raise last_error
    raise RuntimeError(f"Graph sendMail failed for {description}")


def _email_recipient(address: str, name: str | None = None) -> dict[str, Any]:
    return {
        "emailAddress": {
            "address": address.strip(),
            "name": (name or address).strip(),
        }
    }


def _configured_reply_to(
    reply_to_address: str | None,
    reply_to_name: str | None = None,
) -> list[dict[str, Any]]:
    address = (reply_to_address if reply_to_address is not None else EMAIL_REPLY_TO).strip()
    if not address:
        return []
    return [_email_recipient(address, reply_to_name or address)]


def send_mail_with_ics(
    access_token: str,
    *,
    to_address: str,
    to_name: str | None,
    subject: str,
    html_body: str,
    ics_bytes: bytes,
    ics_filename: str = "invite.ics",
    calendar_method: str = "PUBLISH",
    save_to_sent_items: bool = True,
    ics_content_id: str | None = None,
    reply_to_address: str | None = None,
    reply_to_name: str | None = None,
) -> None:
    """
    Send one message with HTML body and a calendar attachment.

    calendar_method should be PUBLISH (appointment item), REQUEST, or CANCEL.
    """
    to_address = to_address.strip()
    if not to_address:
        raise ValueError("to_address is required")

    method_upper = calendar_method.upper()
    content_type = f"text/calendar; method={method_upper}; charset=UTF-8"

    attachment: dict[str, Any] = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": ics_filename,
        "contentType": content_type,
        "contentBytes": base64.b64encode(ics_bytes).decode("ascii"),
        "isInline": False,
    }
    cid = (ics_content_id or "").strip()
    if cid:
        attachment["contentId"] = cid

    recipient = _email_recipient(to_address, to_name or to_address)

    message: dict[str, Any] = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": html_body,
        },
        "toRecipients": [recipient],
        "attachments": [attachment],
    }
    reply_to = _configured_reply_to(reply_to_address, reply_to_name)
    if reply_to:
        message["replyTo"] = reply_to

    payload = {
        "message": message,
        "saveToSentItems": save_to_sent_items,
    }

    _post_send_mail(
        access_token,
        payload,
        description=f"{to_address} subject={subject[:60]}",
    )
    logger.info(
        "Sent mail with %s to %s (subject=%s)",
        ics_filename,
        to_address,
        subject[:60] + ("..." if len(subject) > 60 else ""),
    )


def send_html_email(
    access_token: str,
    *,
    to_address: str,
    to_name: str | None = None,
    subject: str,
    html_body: str,
    save_to_sent_items: bool = True,
    reply_to_address: str | None = None,
    reply_to_name: str | None = None,
) -> None:
    """Send one plain HTML message via Microsoft Graph."""
    to_address = to_address.strip()
    if not to_address:
        raise ValueError("to_address is required")

    recipient = _email_recipient(to_address, to_name or to_address)

    message: dict[str, Any] = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": html_body,
        },
        "toRecipients": [recipient],
    }
    reply_to = _configured_reply_to(reply_to_address, reply_to_name)
    if reply_to:
        message["replyTo"] = reply_to

    payload = {
        "message": message,
        "saveToSentItems": save_to_sent_items,
    }

    _post_send_mail(
        access_token,
        payload,
        description=f"{to_address} subject={subject[:60]}",
    )
    logger.info(
        "Sent HTML mail to %s (subject=%s)",
        to_address,
        subject[:60] + ("..." if len(subject) > 60 else ""),
    )
