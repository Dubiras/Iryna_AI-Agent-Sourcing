# © 2026 Iryna Subbotina. All Rights Reserved. Proprietary — not open source. See LICENSE/NOTICE. Unauthorized copying, redistribution, or claim of authorship is prohibited.
"""Shared Microsoft Graph credentials for Excel/Outlook integration."""
from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any, Optional

import httpx
import msal

log = logging.getLogger(__name__)

SECRETS_DIR = Path(os.environ.get("MS_SECRETS_DIR", "/app/secrets"))
TOKEN_CACHE_PATH = SECRETS_DIR / "ms-token-cache.bin"

MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID", "")
MS_TENANT_ID = os.environ.get("MS_TENANT_ID", "")
AUTHORITY = f"https://login.microsoftonline.com/{MS_TENANT_ID}"

# Reserved scopes (offline_access, openid, profile) are added by MSAL automatically
# and must not be listed explicitly here.
SCOPES = ["Files.ReadWrite", "Calendars.ReadWrite", "Mail.Read"]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

_app: Optional[msal.PublicClientApplication] = None


def _load_app() -> msal.PublicClientApplication:
    global _app
    if _app is not None:
        return _app
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.is_file():
        cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))
    _app = msal.PublicClientApplication(
        MS_CLIENT_ID, authority=AUTHORITY, token_cache=cache
    )
    return _app


def _persist_cache(app: msal.PublicClientApplication) -> None:
    cache = app.token_cache
    if not cache.has_state_changed:
        return
    try:
        TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")
        TOKEN_CACHE_PATH.chmod(0o600)
    except OSError:
        log.debug("Cannot persist refreshed token cache (read-only FS) — using in-memory cache")


def get_access_token() -> str:
    if not MS_CLIENT_ID or not MS_TENANT_ID:
        raise RuntimeError(
            "MS_CLIENT_ID / MS_TENANT_ID env vars are not set. "
            "Register an Azure AD app and set them in .env."
        )
    app = _load_app()
    accounts = app.get_accounts()
    if not accounts:
        raise RuntimeError(
            f"No Microsoft account signed in (no token cache at {TOKEN_CACHE_PATH}). "
            "Run ms_auth_setup.py to authorize."
        )
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        raise RuntimeError(
            "Microsoft token refresh failed or requires re-consent. "
            "Re-run ms_auth_setup.py to re-authorize."
        )
    _persist_cache(app)
    return result["access_token"]


def request(method: str, path: str, **kwargs: Any):
    """Make an authenticated Microsoft Graph API call and return parsed JSON (or {} for empty responses)."""
    token = get_access_token()
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"
    url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
    resp = httpx.request(method, url, headers=headers, timeout=30, **kwargs)
    resp.raise_for_status()
    if resp.status_code == 204 or not resp.content:
        return {}
    return resp.json()


def invalidate_cache() -> None:
    """Force the next call to reload the app/token cache from disk."""
    global _app
    _app = None
