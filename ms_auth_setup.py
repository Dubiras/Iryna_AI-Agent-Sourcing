#!/usr/bin/env python3
# © 2026 Iryna Subbotina. All Rights Reserved. Proprietary — not open source. See LICENSE/NOTICE. Unauthorized copying, redistribution, or claim of authorship is prohibited.
"""Authorize Scout against Microsoft Graph via device-code login.

Unlike setup_gmail_auth.py, this needs no local browser or port-forwarding —
it can be run directly on the server over SSH. Run it once, then restart mcp-scout.
"""
import re
import sys
from pathlib import Path

try:
    import msal
except ImportError:
    print("Installing required package...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "msal"])
    import msal

SCOPES = ["Files.ReadWrite", "Calendars.ReadWrite", "Mail.Read"]

ENV_PATH = Path(".env")
TOKEN_OUT = Path("secrets/ms-token-cache.bin")


def _read_env_var(name: str) -> str:
    value = ""
    if ENV_PATH.is_file():
        for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
            m = re.match(rf"^{name}=(.*)$", line.strip())
            if m:
                value = m.group(1).strip()
    if not value:
        value = input(f"{name}: ").strip()
    return value


client_id = _read_env_var("MS_CLIENT_ID")
tenant_id = _read_env_var("MS_TENANT_ID")

if not client_id or not tenant_id:
    print("❌ MS_CLIENT_ID / MS_TENANT_ID required (set in .env or enter above).")
    sys.exit(1)

TOKEN_OUT.parent.mkdir(parents=True, exist_ok=True)

cache = msal.SerializableTokenCache()
app = msal.PublicClientApplication(
    client_id,
    authority=f"https://login.microsoftonline.com/{tenant_id}",
    token_cache=cache,
)

flow = app.initiate_device_flow(scopes=SCOPES)
if "user_code" not in flow:
    print(f"❌ Failed to start device flow: {flow.get('error_description', flow)}")
    sys.exit(1)

print(flow["message"])
print()
print("Waiting for you to complete sign-in in a browser (any device)...")

result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    print(f"❌ Авторизація не вдалась: {result.get('error_description', result)}")
    sys.exit(1)

TOKEN_OUT.write_text(cache.serialize(), encoding="utf-8")
TOKEN_OUT.chmod(0o600)
print(f"✅ Токен збережено: {TOKEN_OUT}")
print()
print("Перезапусти mcp-scout:")
print("  docker compose restart mcp-scout")
