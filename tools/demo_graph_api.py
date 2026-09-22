#!/usr/bin/env python3
"""
Demo script: Read-only access to Outlook emails via Microsoft Graph API.

This script authenticates the user interactively and retrieves the 5 most
recent emails (subject and sender only). No emails are modified, moved,
or deleted — this is a read-only demonstration.

Prerequisites:
    pip install msal requests

Usage:
    1. Update CLIENT_ID and TENANT_ID below with values from your
       Azure AD app registration.
    2. Run: python3 demo_graph_api.py
    3. A browser window will open for Microsoft login on first run.
       Subsequent runs use the cached token silently.
"""

import json
import os
import sys

# This standalone demo lives outside the shipped workflow; resolve msal/requests
# from the sibling src/lib directory.
_SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir, "src")
for _p in (os.path.join(_SRC, "lib"), _SRC):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import msal
import requests

# =============================================================
# Configuration — replace these with your app registration values
# =============================================================
CLIENT_ID = "YOUR_CLIENT_ID_HERE"
TENANT_ID = "YOUR_TENANT_ID_HERE"

SCOPES = ["Mail.Read"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"

# Token cache file (stored locally, keeps you signed in between runs)
TOKEN_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token_cache.json")


def get_token_cache():
    """Load or create a persistent MSAL token cache."""
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache


def save_token_cache(cache):
    """Save the token cache to disk if it has changed."""
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def authenticate():
    """Authenticate and return an access token."""
    cache = get_token_cache()

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )

    # Try silent authentication first (uses cached refresh token)
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    # Fall back to interactive login
    if not result:
        print("Opening browser for Microsoft login...")
        result = app.acquire_token_interactive(SCOPES)

    save_token_cache(cache)

    if "access_token" not in result:
        print("Authentication failed.")
        print(json.dumps(result, indent=2))
        sys.exit(1)

    return result["access_token"]


def get_recent_emails(token, count=5):
    """Fetch the most recent emails (subject and sender only)."""
    response = requests.get(
        f"{GRAPH_ENDPOINT}/me/messages",
        headers={"Authorization": f"Bearer {token}"},
        params={
            "$top": count,
            "$select": "subject,from,receivedDateTime,isRead",
            "$orderby": "receivedDateTime desc",
        },
    )

    if response.status_code != 200:
        print(f"Error {response.status_code}: {response.text}")
        sys.exit(1)

    return response.json().get("value", [])


def main():
    if CLIENT_ID == "YOUR_CLIENT_ID_HERE":
        print("Please update CLIENT_ID and TENANT_ID in this script")
        print("with values from your Azure AD app registration.")
        sys.exit(1)

    print("Authenticating...")
    token = authenticate()
    print("Authenticated successfully.\n")

    print("Fetching 5 most recent emails...\n")
    emails = get_recent_emails(token)

    if not emails:
        print("No emails found.")
        return

    print(f"{'#':<4} {'Read':<6} {'From':<35} {'Subject'}")
    print("-" * 90)

    for i, msg in enumerate(emails, 1):
        sender = msg.get("from", {}).get("emailAddress", {})
        sender_str = sender.get("name", sender.get("address", "Unknown"))
        subject = msg.get("subject", "(no subject)")
        read = "Yes" if msg.get("isRead") else "No"
        # Truncate long fields for display
        sender_str = sender_str[:33] + ".." if len(sender_str) > 35 else sender_str
        subject = subject[:45] + ".." if len(subject) > 47 else subject
        print(f"{i:<4} {read:<6} {sender_str:<35} {subject}")

    print(f"\nDone. Retrieved {len(emails)} emails (read-only, no modifications made).")


if __name__ == "__main__":
    main()
