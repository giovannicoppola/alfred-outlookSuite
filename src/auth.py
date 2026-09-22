"""Shared authentication for Microsoft Graph API scripts."""

import json
import os
import sys
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

import msal

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")
TOKEN_CACHE_FILE = os.path.join(SCRIPT_DIR, "token_cache.json")
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"

REDIRECT_PORT = 5050
REDIRECT_URI = f"http://localhost:{REDIRECT_PORT}"


class MissingConfigError(Exception):
    """Raised when config.json is absent, so callers can show a setup hint."""

    def __init__(self, path=CONFIG_FILE):
        self.path = path
        super().__init__(f"Config file not found: {path}")


def setup_required_items_json():
    """Alfred Script Filter JSON prompting the user to configure sign-in."""
    return json.dumps({"items": [{
        "title": "Microsoft sign-in not configured",
        "subtitle": "Create config.json (client_id + tenant_id) — see SETUP.md",
        "valid": False,
        "icon": {"path": "icon.png"},
    }]})


class _AuthCallbackHandler(BaseHTTPRequestHandler):
    auth_response = None

    def do_GET(self):
        parsed = urlparse(self.path)
        _AuthCallbackHandler.auth_response = parse_qs(parsed.query)
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        self.wfile.write(b"<html><body><h3>Authentication complete. You can close this tab.</h3></body></html>")

    def log_message(self, format, *args):
        pass


def _load_config():
    if not os.path.exists(CONFIG_FILE):
        raise MissingConfigError(CONFIG_FILE)
    with open(CONFIG_FILE) as f:
        return json.load(f)


def _get_token_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE) as f:
            cache.deserialize(f.read())
    return cache


def _save_token_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def get_token():
    """Authenticate and return an access token. Uses cached token if available."""
    config = _load_config()
    cache = _get_token_cache()
    authority = f"https://login.microsoftonline.com/{config['tenant_id']}"
    scopes = config.get("scopes", ["Mail.ReadWrite", "Calendars.Read"])

    # Two supported app-registration models, chosen by config.json:
    #  - client_secret present  -> confidential client (org apps provisioned with
    #    a secret; no "Allow public client flows" needed on the registration).
    #  - client_secret absent    -> public client + PKCE (no secret to manage;
    #    the registration must have "Allow public client flows" enabled).
    # The auth-code + silent flow below is identical for both.
    secret = config.get("client_secret")
    if secret:
        app = msal.ConfidentialClientApplication(
            config["client_id"],
            authority=authority,
            client_credential=secret,
            token_cache=cache,
        )
    else:
        app = msal.PublicClientApplication(
            config["client_id"],
            authority=authority,
            token_cache=cache,
        )

    # Try silent authentication first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            _save_token_cache(cache)
            return result["access_token"]

    # Interactive login via auth code flow
    flow = app.initiate_auth_code_flow(scopes, redirect_uri=REDIRECT_URI)
    if "auth_uri" not in flow:
        print("Failed to initiate auth code flow.", file=sys.stderr)
        print(json.dumps(flow, indent=2), file=sys.stderr)
        sys.exit(1)

    server = HTTPServer(("localhost", REDIRECT_PORT), _AuthCallbackHandler)
    print("Opening browser for Microsoft login...", file=sys.stderr)
    webbrowser.open(flow["auth_uri"])
    server.handle_request()
    server.server_close()

    auth_response = _AuthCallbackHandler.auth_response
    if not auth_response:
        print("No response received from browser.", file=sys.stderr)
        sys.exit(1)

    auth_response_flat = {k: v[0] for k, v in auth_response.items()}
    result = app.acquire_token_by_auth_code_flow(flow, auth_response_flat)
    _save_token_cache(cache)

    if "access_token" not in result:
        print("Authentication failed.", file=sys.stderr)
        print(json.dumps(result, indent=2), file=sys.stderr)
        sys.exit(1)

    return result["access_token"]
