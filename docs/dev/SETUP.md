# Outlook Graph API — Setup Guide

This guide walks through everything needed to run the Outlook Alfred workflows with the Microsoft Graph API (required for the new Outlook for Mac, which dropped AppleScript support).

---

## Prerequisites

- macOS with [Homebrew](https://brew.sh) installed
- [Alfred](https://www.alfredapp.com) with Powerpack
- Python 3.9+ — macOS's built-in `/usr/bin/python3` works (bundled deps are pure-Python)
- Microsoft Outlook for Mac (new or legacy)
- macOS Accessibility permissions for Alfred (System Settings > Privacy & Security > Accessibility)

---

## Step 1: Request an Entra App Registration

You need your IT team to register an application in Microsoft Entra (Azure AD). Send them the following requirements:

### App Registration Settings

| Setting | Value |
|---------|-------|
| Application name | Outlook Keyboard Shortcuts (or your choice) |
| Supported account types | Single tenant (your organization only) |
| Platform | Mobile and desktop applications |
| Redirect URI | `http://localhost:5050` |
| Allow public client flows | Yes (no client secret — PKCE) |

### API Permissions (Delegated)

| Permission | Type | Purpose |
|------------|------|---------|
| `User.Read` | Delegated | Sign-in and basic profile |
| `Mail.ReadWrite` | Delegated | Read, move, delete, flag emails |
| `Calendars.Read` | Delegated | Read calendar events |

All permissions must be **delegated** (not application). Admin consent may be required depending on your tenant policy.

### What IT Will Give You

You will receive two values:

1. **Client ID** (Application ID) — a UUID like `aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee`
2. **Tenant ID** (Directory ID) — a UUID like `11111111-2222-3333-4444-555555555555`

There is no client secret — this is a public client app that authenticates with PKCE, so nothing needs to be stored securely and nothing expires.

---

## Step 2: Clone the Repository

```bash
git clone https://github.com/giovannicoppola/alfred-outlookSuite.git
cd alfred-outlookSuite/src
```

Or if you already have the repo, navigate to the `src` directory.

---

## Step 3: Create the Python Virtual Environment

```bash
cd src
python3 -m venv venv
source venv/bin/activate
pip install msal requests
deactivate
```

This creates a `venv/` folder with its own Python and the two required packages:
- **msal** — Microsoft Authentication Library (handles OAuth2 login + token caching)
- **requests** — HTTP client for Graph API calls

---

## Step 4: Create the Config File

Create a file called `config.json` in the `src` directory:

```json
{
    "client_id": "YOUR_CLIENT_ID_HERE",
    "tenant_id": "YOUR_TENANT_ID_HERE",
    "scopes": ["Mail.ReadWrite", "Calendars.Read"]
}
```

Replace the two placeholder values with the credentials from Step 1.

**Important:** This file is gitignored and must never be committed to version control.

---

## Step 5: First-Run Authentication

Run any script to trigger the initial login:

```bash
./venv/bin/python3 inbox_count.py
```

This will:
1. Open your default browser to the Microsoft login page
2. You sign in with your organization credentials (including MFA if required)
3. The browser redirects to `http://localhost:5050` — a temporary local server captures the auth code
4. You see "Authentication complete. You can close this tab." in the browser
5. The script prints the inbox count (e.g. `Inbox: 1175 emails (42 unread)`)

A `token_cache.json` file is created locally. Subsequent runs will use the cached token silently (no browser popup). The token refreshes automatically until the refresh token expires (typically weeks/months), at which point the browser login will repeat once.

---

## Step 6: Grant Accessibility Permissions

The email detection scripts use the macOS Accessibility API (via JXA/osascript) to read which email or calendar event is selected in Outlook. This requires:

1. Open **System Settings > Privacy & Security > Accessibility**
2. Ensure **Alfred** is listed and enabled
3. If using Terminal for testing, also enable **Terminal** (or **iTerm2**)

Without this, the scripts that detect the selected email/event will fail silently.

---

## Step 7: Configure Alfred Workflow

In each Alfred workflow "Run Script" action, use the following pattern:

```bash
OUTLOOK_SHORTCUTS_DIR="$HOME/path/to/alfred-outlookSuite/src"
"$OUTLOOK_SHORTCUTS_DIR/venv/bin/python3" "$OUTLOOK_SHORTCUTS_DIR/SCRIPT_NAME.py"
```

Replace `SCRIPT_NAME.py` with the appropriate script. See `ALFRED_COMMANDS.md` for the full list of commands and their arguments.

**Tip:** Adjust `OUTLOOK_SHORTCUTS_DIR` to match your installation path. Always quote paths that contain spaces or special characters.

---

## Workflow Configuration Variables

If you want to make the workflow configurable for other users, expose these as Alfred workflow environment variables:

| Variable | Default | Description |
|----------|---------|-------------|
| `OUTLOOK_SHORTCUTS_DIR` | `$HOME/path/to/alfred-outlookSuite/src` | Path to the scripts directory |
| `MINUTES_AHEAD` | `5` or `10` | How far ahead to look for upcoming meetings |
| `SAVE_DIR` | `~/Documents/SavedEmails` | Where .eml files are saved |

---

## File Reference

| File | Committed | Purpose |
|------|-----------|---------|
| `auth.py` | Yes | Shared authentication (OAuth2 + token caching) |
| `config.json` | **No** (gitignored) | Your client ID and tenant ID (no secret — public client) |
| `token_cache.json` | **No** (gitignored) | Cached OAuth tokens (auto-generated) |
| `venv/` | **No** (gitignored) | Python virtual environment (auto-generated) |
| `get_selected.py` | Yes | Detect selected email via Accessibility API + Graph API |
| `snooze.py` | Yes | Move selected email/conversation to Snoozed folder |
| `save_eml.py` | Yes | Save selected email as .eml file |
| `inbox_count.py` | Yes | Get inbox count (verbose: total + unread) |
| `get_inbox_count.py` | Yes | Get inbox count (just the number, with new/old Outlook auto-detection) |
| `get_meeting_details.py` | Yes | Get upcoming meetings in delimited format |
| `get_current_meeting_note.py` | Yes | Get current meeting as markdown note |
| `get_selected_event.py` | Yes | Get selected calendar event as markdown note |

---

## Troubleshooting

### "Config file not found"
Create `config.json` per Step 4.

### Browser opens but authentication fails
- Verify the redirect URI in your Entra app registration is exactly `http://localhost:5050`
- Check that port 5050 is not in use by another application
- Ensure admin consent was granted for the API permissions

### "No email selected in Outlook"
- Make sure Outlook is the frontmost app with an email selected
- Check Accessibility permissions (Step 6)
- The Accessibility API reads the UI structure of Outlook — if Microsoft changes the layout in an update, the JXA scripts may need adjustment

### Token expired / login keeps prompting
- There's no client secret to expire (public client / PKCE). Delete `token_cache.json` and re-authenticate — the browser sign-in runs once more and caches a fresh token

### Scripts work in Terminal but not from Alfred
- Alfred needs Accessibility permissions (Step 6)
- Alfred's environment may differ from your shell. The scripts use absolute paths to the venv Python to avoid PATH issues
