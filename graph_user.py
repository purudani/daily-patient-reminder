"""Resolve the sending user's mailbox address (delegated /me) when GRAPH_MAILBOX_USER is unset."""
from __future__ import annotations

import logging

import requests

logger = logging.getLogger(__name__)

GRAPH_ME = "https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName"


def resolve_organizer_email(access_token: str) -> str:
    """
    Return GRAPH_MAILBOX_USER if set; otherwise GET /me (delegated only).

    App-only flows must set GRAPH_MAILBOX_USER (no /me).
    """
    from config import GRAPH_CLIENT_SECRET, GRAPH_MAILBOX_USER

    u = (GRAPH_MAILBOX_USER or "").strip()
    if u:
        return u
    if (GRAPH_CLIENT_SECRET or "").strip():
        raise RuntimeError(
            "GRAPH_MAILBOX_USER is required for app-only mode (used as organizer in .ics and sendMail path)."
        )
    resp = requests.get(
        GRAPH_ME,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    mail = (data.get("mail") or data.get("userPrincipalName") or "").strip()
    if not mail:
        raise RuntimeError(
            "Could not read mail from Graph /me; set GRAPH_MAILBOX_USER in .env."
        )
    logger.info("Using organizer email from /me: %s", mail)
    return mail
