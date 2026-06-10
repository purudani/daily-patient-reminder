"""
Send email via Microsoft Graph with a text/calendar (.ics) attachment.

Uses /users/{GRAPH_MAILBOX_USER}/sendMail when GRAPH_MAILBOX_USER is set, else /me/sendMail.
"""
from __future__ import annotations

import base64
import logging
from typing import Any
from urllib.parse import quote

import requests

from config import EMAIL_REPLY_TO, GRAPH_MAILBOX_USER

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


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

    url = f"{_user_root()}/sendMail"
    resp = requests.post(url, json=payload, headers=_headers(access_token), timeout=60)
    resp.raise_for_status()
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

    url = f"{_user_root()}/sendMail"
    resp = requests.post(url, json=payload, headers=_headers(access_token), timeout=60)
    resp.raise_for_status()
    logger.info(
        "Sent HTML mail to %s (subject=%s)",
        to_address,
        subject[:60] + ("..." if len(subject) > 60 else ""),
    )
