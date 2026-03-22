"""
Microsoft Graph authentication.

Two modes (see config.py):
- App-only (client credentials): GRAPH_CLIENT_SECRET + GRAPH_TENANT_ID + GRAPH_MAILBOX_USER
  → no sign-in; token is for the app; mail uses /users/{mailbox}/sendMail
- Delegated (public client): GRAPH_CLIENT_SECRET empty → device code / interactive; /me or /users/{GRAPH_MAILBOX_USER}
"""
from __future__ import annotations

import logging
from pathlib import Path

logger = logging.getLogger(__name__)


def _client_secret_set() -> bool:
    from config import GRAPH_CLIENT_SECRET
    return bool((GRAPH_CLIENT_SECRET or "").strip())


def get_access_token() -> str:
    """Return a valid Graph access token (app-only or delegated)."""
    try:
        import msal
    except ImportError:
        raise ImportError("Install msal: pip install msal")

    from config import (
        GRAPH_APP_ONLY_SCOPE,
        GRAPH_CLIENT_ID,
        GRAPH_CLIENT_SECRET,
        GRAPH_SCOPES,
        GRAPH_TENANT_ID,
    )

    if not (GRAPH_CLIENT_ID or "").strip():
        raise RuntimeError("GRAPH_CLIENT_ID is not set.")

    if _client_secret_set():
        tenant = (GRAPH_TENANT_ID or "").strip()
        if not tenant:
            raise RuntimeError(
                "GRAPH_TENANT_ID is required when GRAPH_CLIENT_SECRET is set (app-only auth)."
            )
        authority = f"https://login.microsoftonline.com/{tenant}"
        app = msal.ConfidentialClientApplication(
            GRAPH_CLIENT_ID.strip(),
            authority=authority,
            client_credential=GRAPH_CLIENT_SECRET.strip(),
        )
        result = app.acquire_token_for_client(scopes=GRAPH_APP_ONLY_SCOPE)
        if "access_token" not in result:
            raise RuntimeError(
                "App-only token failed: "
                + result.get("error_description", result.get("error", "unknown"))
            )
        logger.info("Using Microsoft Graph app-only (client credentials) token")
        return result["access_token"]

    # Delegated: public client, device code or browser
    tenant = GRAPH_TENANT_ID.strip() if GRAPH_TENANT_ID else "common"
    authority = f"https://login.microsoftonline.com/{tenant}"
    cache_file = Path(__file__).resolve().parent / ".graph_token_cache.json"

    app = msal.PublicClientApplication(
        GRAPH_CLIENT_ID.strip(),
        authority=authority,
    )

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
        if result:
            logger.info("Using cached delegated Graph token")
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
    if "message" in flow:
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)
    else:
        result = app.acquire_token_interactive(scopes=GRAPH_SCOPES)

    if "access_token" not in result:
        raise RuntimeError(
            "Delegated token failed: "
            + result.get("error_description", result.get("error", "unknown"))
        )
    logger.info("Using delegated Graph token (sign-in)")
    return result["access_token"]
