#!/usr/bin/env python3
"""
Test script: Verify that the app registration credentials work.

This script authenticates using the Microsoft Graph API and performs
two simple read-only operations:
  1. Fetches your profile info (name, email)
  2. Fetches the 5 most recent emails (subject and sender only)

No emails are modified, moved, or deleted.

Prerequisites:
    - Virtual environment with msal and requests installed
    - config.json with client_id, tenant_id, and client_secret

Usage:
    ./venv/bin/python test_connection.py
"""

import json
import os
import sys
import threading
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# This CLI utility lives outside the shipped workflow; resolve its deps
# (bundled libs + config.json + token cache) from the sibling src/ directory.
_SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir, "src")
for _p in (os.path.join(_SRC, "lib"), _SRC):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import msal
import requests

SCRIPT_DIR = _SRC
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")
TOKEN_CACHE_FILE = os.path.join(SCRIPT_DIR, "token_cache.json")
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"


def load_config():
    if not os.path.exists(CONFIG_FILE):
        print(f"Config file not found: {CONFIG_FILE}")
        sys.exit(1)
    with open(CONFIG_FILE) as f:
        return json.load(f)


def get_token_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE) as f:
            cache.deserialize(f.read())
    return cache


def save_token_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


REDIRECT_PORT = 5050
REDIRECT_URI = f"http://localhost:{REDIRECT_PORT}"


class AuthCallbackHandler(BaseHTTPRequestHandler):
    """Handles the OAuth redirect callback on localhost."""

    auth_response = None

    def do_GET(self):
        parsed = urlparse(self.path)
        AuthCallbackHandler.auth_response = parse_qs(parsed.query)
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        self.wfile.write(b"<html><body><h3>Authentication complete. You can close this tab.</h3></body></html>")

    def log_message(self, format, *args):
        pass  # suppress server logs


def authenticate(config):
    cache = get_token_cache()
    authority = f"https://login.microsoftonline.com/{config['tenant_id']}"
    scopes = config.get("scopes", ["Mail.Read"])

    app = msal.ConfidentialClientApplication(
        config["client_id"],
        authority=authority,
        client_credential=config["client_secret"],
        token_cache=cache,
    )

    # Try silent authentication first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            print("Authenticated silently (cached token).")
            save_token_cache(cache)
            return result["access_token"]

    # Interactive login via auth code flow with local redirect server
    flow = app.initiate_auth_code_flow(scopes, redirect_uri=REDIRECT_URI)
    if "auth_uri" not in flow:
        print("Failed to initiate auth code flow.")
        print(json.dumps(flow, indent=2))
        sys.exit(1)

    # Start a local server to capture the redirect
    server = HTTPServer(("localhost", REDIRECT_PORT), AuthCallbackHandler)

    print("Opening browser for Microsoft login...")
    webbrowser.open(flow["auth_uri"])

    # Wait for the callback
    server.handle_request()
    server.server_close()

    # Build the full response URL from the captured query params
    auth_response = AuthCallbackHandler.auth_response
    if not auth_response:
        print("No response received from browser.")
        sys.exit(1)

    # Flatten the query params (parse_qs returns lists)
    auth_response_flat = {k: v[0] for k, v in auth_response.items()}

    result = app.acquire_token_by_auth_code_flow(flow, auth_response_flat)

    save_token_cache(cache)

    if "access_token" not in result:
        print("\nAuthentication FAILED.")
        print(json.dumps(result, indent=2))
        sys.exit(1)

    print("Authenticated successfully.")
    return result["access_token"]


def test_profile(token):
    """Test 1: Fetch the signed-in user's profile."""
    print("\n--- Test 1: User Profile ---")
    response = requests.get(
        f"{GRAPH_ENDPOINT}/me",
        headers={"Authorization": f"Bearer {token}"},
        params={"$select": "displayName,mail,jobTitle"},
    )
    if response.status_code != 200:
        print(f"FAILED (HTTP {response.status_code}): {response.text}")
        return False

    profile = response.json()
    print(f"  Name:  {profile.get('displayName', 'N/A')}")
    print(f"  Email: {profile.get('mail', 'N/A')}")
    print(f"  Title: {profile.get('jobTitle', 'N/A')}")
    return True


def test_recent_emails(token, count=5):
    """Test 2: Fetch the most recent emails."""
    print(f"\n--- Test 2: {count} Most Recent Emails ---")
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
        print(f"FAILED (HTTP {response.status_code}): {response.text}")
        return False

    emails = response.json().get("value", [])
    if not emails:
        print("  No emails found (mailbox may be empty).")
        return True

    print(f"  {'#':<3} {'Read':<6} {'From':<35} {'Subject'}")
    print(f"  {'-'*80}")
    for i, msg in enumerate(emails, 1):
        sender = msg.get("from", {}).get("emailAddress", {})
        sender_str = sender.get("name", sender.get("address", "Unknown"))
        subject = msg.get("subject", "(no subject)")
        read = "Yes" if msg.get("isRead") else "No"
        sender_str = (sender_str[:33] + "..") if len(sender_str) > 35 else sender_str
        subject = (subject[:42] + "..") if len(subject) > 44 else subject
        print(f"  {i:<3} {read:<6} {sender_str:<35} {subject}")

    return True


def main():
    config = load_config()
    print("Configuration loaded.")
    print(f"  Client ID: {config['client_id'][:8]}...")
    print(f"  Tenant ID: {config['tenant_id'][:8]}...")
    print()

    token = authenticate(config)

    passed = 0
    total = 2

    if test_profile(token):
        passed += 1
    if test_recent_emails(token):
        passed += 1

    print(f"\n{'='*50}")
    print(f"Results: {passed}/{total} tests passed.")
    if passed == total:
        print("Everything is working! You're ready to build the full workflow.")
    else:
        print("Some tests failed. Check the error messages above.")


if __name__ == "__main__":
    main()
