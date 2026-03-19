"""Acquire a Microsoft Graph delegated token via MSAL device-code flow.

First run: opens browser for device-code login.
Subsequent runs: silently refreshes cached token.

Token is printed to stdout (for piping to start-mcp.cmd).
All prompts/errors go to stderr.

Requires: pip install msal
"""

import json
import os
import sys

import msal

# Entra app "Graph-MCP-User" — public client, delegated permissions
# Create this app in Azure Portal > App registrations > New registration
#   - Name: Graph-MCP-User
#   - Supported account types: Single tenant
#   - Redirect URI: (none needed for device-code)
#   - Under Authentication > Advanced > Allow public client flows: Yes
#   - Under API permissions > Add: Microsoft Graph delegated permissions
CLIENT_ID = os.getenv("GRAPH_MCP_CLIENT_ID", "02fa0ea1-4b30-4bd9-9c4a-483f97d63b21")
TENANT_ID = os.getenv("AZURE_TENANT_ID", "d54794b1-f598-4c0f-a276-6039a39774ac")

# Delegated Graph permissions — user only gets what their Entra role permits
SCOPES = [
    "User.Read",
    "User.ReadBasic.All",
    "Mail.ReadWrite",
    "Mail.Send",
    "Files.ReadWrite",
    "Calendars.ReadWrite",
    "Sites.Read.All",
    "Tasks.ReadWrite",
    "Chat.ReadWrite",
    "OnlineMeetings.Read",
]

CACHE_DIR = os.path.join(os.path.expanduser("~"), ".connector-graph")
CACHE_FILE = os.path.join(CACHE_DIR, "token_cache.bin")


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE) as f:
            cache.deserialize(f.read())
    return cache


def _save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        os.makedirs(CACHE_DIR, exist_ok=True)
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def main():
    if not CLIENT_ID:
        print(
            "ERROR: GRAPH_MCP_CLIENT_ID not set.\n"
            "Create the Entra app 'Graph-MCP-User' (public client) and set this env var to its client ID.\n"
            "See connector-graph/CLAUDE.md for setup instructions.",
            file=sys.stderr,
        )
        sys.exit(1)

    cache = _load_cache()
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )

    # Try silent token acquisition first (cached refresh token)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            print(result["access_token"])  # stdout — piped to start-mcp.cmd
            return

    # Fall back to device code flow (interactive, first time or expired refresh token)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print(
            f"ERROR: Could not initiate device flow: {json.dumps(flow, indent=2)}",
            file=sys.stderr,
        )
        sys.exit(1)

    # Show device code prompt on stderr (stdout reserved for token)
    print(flow["message"], file=sys.stderr)

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        _save_cache(cache)
        print(result["access_token"])  # stdout only
    else:
        print(
            f"ERROR: {result.get('error_description', json.dumps(result))}",
            file=sys.stderr,
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
