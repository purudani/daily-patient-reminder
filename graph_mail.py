"""
Send email via Microsoft Graph with an text/calendar (.ics) attachment.

Uses /users/{GRAPH_MAILBOX_USER}/sendMail when GRAPH_MAILBOX_USER is set, else /me/sendMail.
"""
from __future__ import annotations

import base64
import logging
from typing import Any
from urllib.parse import quote

import requests

from config import GRAPH_MAILBOX_USER

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


def send_mail_with_ics(
    access_token: str,
    *,
    to_address: str,
    to_name: str | None,
    subject: str,
    html_body: str,
    ics_bytes: bytes,
    ics_filename: str = "invite.ics",
    calendar_method: str = "REQUEST",
    save_to_sent_items: bool = True,
) -> None:
    """
    Send one message with HTML body and a calendar attachment.

    calendar_method should be REQUEST (invite/update) or CANCEL.
    """
    to_address = to_address.strip()
    if not to_address:
        raise ValueError("to_address is required")

    method_upper = calendar_method.upper()
    content_type = f"text/calendar; charset=UTF-8; method={method_upper}"

    attachment: dict[str, Any] = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": ics_filename,
        "contentType": content_type,
        "contentBytes": base64.b64encode(ics_bytes).decode("ascii"),
        "isInline": False,
    }

    recipient: dict[str, Any] = {
        "emailAddress": {
            "address": to_address,
            "name": (to_name or to_address).strip(),
        }
    }

    message: dict[str, Any] = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": html_body,
        },
        "toRecipients": [recipient],
        "attachments": [attachment],
    }

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
