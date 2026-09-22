# Outlook Email Management via Microsoft Graph API — Workflow Design

## Overview

Replace AppleScript-based Outlook email actions with a Graph API approach, enabling migration from legacy Outlook for Mac to the new version. Actions are triggered via Alfred keyboard shortcuts.

Features:
- **Email management**: archive, delete, flag, move, snooze, mark read/unread
- **Draft creation**: create email drafts with a provided subject line
- **Calendar**: read details about calendar events (selected, today, current hour)

## Core Concept

1. Select an email in Outlook
2. Press an Alfred hotkey
3. Alfred copies the message link from Outlook
4. Extracts the message ID from the link
5. Calls the Graph API to perform the desired action

## Step-by-Step Process

### Step 1: Get the Message ID from Outlook

New Outlook for Mac has a **"Copy Link"** feature (right-click > Copy Link). This produces a URL like:

```
https://outlook.office365.com/mail/id/AAMkADBlYzI1...
```

The message ID is the portion after `/id/`. This can be triggered via:
- A keyboard shortcut assigned to "Copy Link" in Outlook
- AppleScript `System Events` to simulate the right-click menu action
- The Shortcuts app on macOS to automate the UI interaction

### Step 2: Extract the Message ID

Parse the clipboard content to extract the Graph API message ID:

```python
import urllib.parse

link = "https://outlook.office365.com/mail/id/AAMkADBlYzI1..."
message_id = urllib.parse.unquote(link.split("/id/")[-1])
```

### Step 3: Authenticate with Graph API

Use OAuth2 with a cached token. On first run, the user logs in via browser. Subsequent calls use a refresh token silently.

```python
from msal import PublicClientApplication

app = PublicClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}")
scopes = ["Mail.ReadWrite", "Calendars.Read"]

# Try silent token acquisition first
accounts = app.get_accounts()
if accounts:
    result = app.acquire_token_silent(scopes, account=accounts[0])
else:
    # Interactive login (first time or expired)
    result = app.acquire_token_interactive(scopes)

access_token = result["access_token"]
```

### Step 4: Perform Actions via Graph API

With the access token and message ID, call the appropriate endpoint.

#### Archive (move to Archive folder)
```python
requests.post(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/move",
    headers={"Authorization": f"Bearer {access_token}"},
    json={"destinationId": "archive"}
)
```

#### Delete
```python
requests.delete(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}",
    headers={"Authorization": f"Bearer {access_token}"}
)
```

#### Move to a specific folder
```python
# First, get folder IDs
response = requests.get(
    "https://graph.microsoft.com/v1.0/me/mailFolders",
    headers={"Authorization": f"Bearer {access_token}"}
)
folders = response.json()["value"]

# Then move
requests.post(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/move",
    headers={"Authorization": f"Bearer {access_token}"},
    json={"destinationId": folder_id}
)
```

#### Mark as read/unread
```python
requests.patch(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}",
    headers={"Authorization": f"Bearer {access_token}"},
    json={"isRead": True}  # or False to mark unread
)
```

#### Flag / unflag
```python
requests.patch(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}",
    headers={"Authorization": f"Bearer {access_token}"},
    json={"flag": {"flagStatus": "flagged"}}  # or "notFlagged"
)
```

#### Snooze (move to snooze folder + SQLite tracking)

Snoozing works by moving the email out of the inbox and recording when to bring it back.

**Step 1: Move email to a dedicated snooze folder**
```python
response = requests.post(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/move",
    headers={"Authorization": f"Bearer {access_token}"},
    json={"destinationId": snooze_folder_id}
)
# IMPORTANT: moving returns a new message ID
new_message_id = response.json()["id"]
```

**Step 2: Record the snooze in SQLite**
```python
import sqlite3
from datetime import datetime

conn = sqlite3.connect("snooze.db")
conn.execute("""
    CREATE TABLE IF NOT EXISTS snoozed (
        message_id TEXT PRIMARY KEY,
        subject TEXT,
        snooze_until TEXT
    )
""")
conn.execute(
    "INSERT INTO snoozed (message_id, subject, snooze_until) VALUES (?, ?, ?)",
    (new_message_id, subject, "2026-03-06T09:00:00"),
)
conn.commit()
```

**Step 3: Background job to resurface snoozed emails**

A cron job or macOS launchd agent runs periodically (e.g., every 5 minutes) and moves due emails back to the inbox:

```python
now = datetime.now().isoformat()
due = conn.execute(
    "SELECT message_id FROM snoozed WHERE snooze_until <= ?", (now,)
).fetchall()

for (msg_id,) in due:
    requests.post(
        f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}/move",
        headers={"Authorization": f"Bearer {access_token}"},
        json={"destinationId": "inbox"}
    )
    conn.execute("DELETE FROM snoozed WHERE message_id = ?", (msg_id,))

conn.commit()
```

#### Get message details (subject, sender, etc.)
```python
response = requests.get(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}",
    headers={"Authorization": f"Bearer {access_token}"},
    params={"$select": "subject,from,receivedDateTime,isRead"}
)
message = response.json()
```

#### Create a draft email with a subject line
```python
requests.post(
    "https://graph.microsoft.com/v1.0/me/messages",
    headers={"Authorization": f"Bearer {access_token}"},
    json={
        "subject": "Your subject line here",
        "body": {"contentType": "text", "content": ""},
        "toRecipients": []  # left empty — user fills in manually
    }
)
```
This creates a draft in the Drafts folder. The user can then open it in Outlook to add recipients and body text.

## Calendar Events

Calendar actions don't require a selected email — they query the calendar directly.

### Get today's events
```python
from datetime import datetime, timezone

today_start = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0).isoformat()
today_end = datetime.now(timezone.utc).replace(hour=23, minute=59, second=59).isoformat()

response = requests.get(
    "https://graph.microsoft.com/v1.0/me/calendarView",
    headers={
        "Authorization": f"Bearer {access_token}",
        "Prefer": 'outlook.timezone="Eastern Standard Time"',
    },
    params={
        "startDateTime": today_start,
        "endDateTime": today_end,
        "$select": "subject,start,end,location,organizer,isAllDay",
        "$orderby": "start/dateTime",
    },
)
events = response.json().get("value", [])
```

### Get events happening right now (this hour)
```python
from datetime import datetime, timedelta, timezone

now = datetime.now(timezone.utc)
hour_end = now + timedelta(hours=1)

response = requests.get(
    "https://graph.microsoft.com/v1.0/me/calendarView",
    headers={
        "Authorization": f"Bearer {access_token}",
        "Prefer": 'outlook.timezone="Eastern Standard Time"',
    },
    params={
        "startDateTime": now.isoformat(),
        "endDateTime": hour_end.isoformat(),
        "$select": "subject,start,end,location,organizer,onlineMeeting",
        "$orderby": "start/dateTime",
    },
)
events = response.json().get("value", [])
```

### Get a specific event by ID
Similar to emails, if Outlook provides a calendar event link/ID:
```python
response = requests.get(
    f"https://graph.microsoft.com/v1.0/me/events/{event_id}",
    headers={"Authorization": f"Bearer {access_token}"},
    params={"$select": "subject,start,end,location,organizer,attendees,body,onlineMeeting"},
)
event = response.json()
```

### Useful event fields
| Field | Description |
|-------|-------------|
| `subject` | Event title |
| `start.dateTime` / `end.dateTime` | Start and end times |
| `location.displayName` | Meeting room or location |
| `organizer.emailAddress` | Who organized the meeting |
| `attendees` | List of attendees with response status |
| `onlineMeeting.joinUrl` | Teams meeting link (if applicable) |
| `body.content` | Event body/notes |
| `isAllDay` | Whether it's an all-day event |

## Save Email as .eml File

Save the complete email as an `.eml` file that can be opened directly in Outlook (or any email client). This is the raw MIME content — includes headers, body, inline images, and all attachments in a single file.

```python
import os
import re

# Get subject for the filename
msg_meta = requests.get(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}",
    headers={"Authorization": f"Bearer {access_token}"},
    params={"$select": "subject,receivedDateTime"},
).json()

# Download raw MIME content
response = requests.get(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/$value",
    headers={"Authorization": f"Bearer {access_token}"},
)

safe_subject = re.sub(r'[^\w\s-]', '', msg_meta['subject'])[:80].strip()
filename = f"{msg_meta['receivedDateTime'][:10]}_{safe_subject}.eml"

save_dir = os.path.expanduser("~/Documents/SavedEmails")
os.makedirs(save_dir, exist_ok=True)

with open(os.path.join(save_dir, filename), "wb") as f:
    f.write(response.content)
```

## Save Email as Text

Save the full email (metadata + body) to a local text file. Useful for plain-text archiving or further processing.

```python
import os
import re

response = requests.get(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}",
    headers={
        "Authorization": f"Bearer {access_token}",
        "Prefer": 'outlook.body-content-type="text"',
    },
    params={"$select": "subject,from,toRecipients,ccRecipients,receivedDateTime,body"},
)
msg = response.json()

# Build a readable text file
sender = msg["from"]["emailAddress"]
to_list = ", ".join(r["emailAddress"]["address"] for r in msg.get("toRecipients", []))
cc_list = ", ".join(r["emailAddress"]["address"] for r in msg.get("ccRecipients", []))

content = f"""Subject: {msg['subject']}
From: {sender['name']} <{sender['address']}>
To: {to_list}
Cc: {cc_list}
Date: {msg['receivedDateTime']}

{msg['body']['content']}
"""

# Sanitize subject for filename
safe_subject = re.sub(r'[^\w\s-]', '', msg['subject'])[:80].strip()
filename = f"{msg['receivedDateTime'][:10]}_{safe_subject}.txt"

save_dir = os.path.expanduser("~/Documents/SavedEmails")
os.makedirs(save_dir, exist_ok=True)

with open(os.path.join(save_dir, filename), "w") as f:
    f.write(content)
```

## Save Attachments

Download all attachments from an email to a local folder.

```python
import base64

response = requests.get(
    f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments",
    headers={"Authorization": f"Bearer {access_token}"},
)
attachments = response.json().get("value", [])

save_dir = os.path.expanduser("~/Documents/SavedEmails/attachments")
os.makedirs(save_dir, exist_ok=True)

for att in attachments:
    if att.get("@odata.type") == "#microsoft.graph.fileAttachment":
        file_content = base64.b64decode(att["contentBytes"])
        filepath = os.path.join(save_dir, att["name"])
        with open(filepath, "wb") as f:
            f.write(file_content)
```

You can combine both — save the email text and its attachments in one action, or offer them as separate hotkeys.

## Alfred Workflow Design

Each action would be a separate Alfred hotkey or keyword:

| Hotkey | Action |
|--------|--------|
| Ctrl+Shift+A | Archive selected email |
| Ctrl+Shift+D | Delete selected email |
| Ctrl+Shift+R | Mark as read |
| Ctrl+Shift+U | Mark as unread |
| Ctrl+Shift+F | Toggle flag |
| Ctrl+Shift+S | Snooze (prompt for duration) |
| Ctrl+Shift+M | Move to folder (show folder list) |
| Ctrl+Shift+N | Create draft with subject (prompt for subject) |
| Ctrl+Shift+T | Show today's calendar events |
| Ctrl+Shift+H | Show events happening this hour |
| Ctrl+Shift+E | Save email as .eml (openable in Outlook) |
| Ctrl+Shift+W | Save email as plain text |
| Ctrl+Shift+X | Save email attachments to folder |

Each hotkey triggers a script that:
1. Simulates "Copy Link" in Outlook (via `osascript` or keyboard shortcut)
2. Reads the clipboard
3. Extracts the message ID
4. Calls a Python script with the action and message ID as arguments

## Dependencies

- **Python 3**
- **msal** — Microsoft Authentication Library (`pip install msal`)
- **requests** — HTTP client (`pip install requests`)
- **Alfred Powerpack** — for hotkey workflows

## Configuration

A config file (e.g., `~/.outlook-shortcuts/config.json`) stores:

```json
{
    "client_id": "your-app-client-id",
    "tenant_id": "your-tenant-id",
    "scopes": ["Mail.ReadWrite", "Calendars.Read"],
    "token_cache": "~/.outlook-shortcuts/token_cache.json"
}
```

## Migration Map: Current AppleScript Workflow to Graph API

Your current `alfred-outlook` workflow uses a hybrid approach: Python scripts that query Outlook's local SQLite database for searching/listing, and AppleScript for actions (move, snooze, unsnooze). Here's how each piece maps:

| Current component | What it does | Graph API replacement |
|---|---|---|
| `main.py` + `compileSQL()` | Searches emails by querying `Outlook.sqlite` directly with SQL filters (from, to, subject, folder, etc.) | `GET /me/messages` with `$filter` and `$search` query parameters |
| `main.py` + `handle()` | Reads email metadata (subject, sender, date, read flag, attachments, preview) from SQLite | `GET /me/messages` with `$select` for the fields you need |
| `snoozer.py` | AppleScript moves selected emails to "Snoozed" folder, Python stores message IDs + snooze date in JSON | `POST /me/messages/{id}/move` to snooze folder + SQLite/JSON for tracking (same approach, new ID after move) |
| `unSnoozer.py` | Reads JSON for due emails, converts IDs, AppleScript moves them back to Inbox, logs to Beeminder | `POST /me/messages/{id}/move` back to inbox + same Beeminder curl call |
| `thread.py` | Queries SQLite for all messages with same `Message_ThreadTopic` | `GET /me/messages?$filter=conversationId eq '{id}'` (Graph uses conversation IDs for threads) |
| `outlook-rebuild.py` | Rebuilds local folder/account/contact caches from SQLite | `GET /me/mailFolders`, `GET /me/contacts` — or may not be needed since Graph handles this server-side |
| `countSnoozed.py` / `countSnoozedDates.py` | Reads snooze JSON to report counts by date | No change needed — same JSON/SQLite logic works |
| `consts.py` `OUTLOOK_DB_FILE` | Path to local Outlook SQLite database | No longer needed — all data comes from Graph API |
| `consts.py` `OUTLOOK_MSG_FOLDER` | Path to local message data files | No longer needed — email body/attachments come from Graph API |
| `util.py` `fetchRecordID` / `fetchEmailID` | Converts between internal DB IDs and message IDs | No longer needed — Graph API uses a single message ID |

### Key differences to be aware of

1. **No local SQLite database**: Your current search is extremely fast because it queries a local database. Graph API searches go over the network. For simple queries this is fine (~200-500ms), but complex multi-filter queries may feel slower.

2. **Message IDs change on move**: In your current system, you use `fetchRecordID` to get a stable message ID for the snoozer. With Graph API, the message ID changes after a `move` operation — you must capture the new ID from the response.

3. **No direct "get selected message"**: Your AppleScript `set msgSet to selection` gets the currently selected message directly. With Graph API, you need the "Copy Link" workaround to identify which message the user is looking at.

4. **Folder IDs**: Your current code looks up folders by name in SQLite. Graph API uses folder IDs, but also supports well-known folder names like `"inbox"`, `"archive"`, `"drafts"`. Custom folders need a one-time ID lookup via `GET /me/mailFolders`.

5. **Beeminder integration**: The `unSnoozer.py` Beeminder posting (curl to Beeminder API) works independently of Outlook — this transfers directly with no changes.

## Limitations and Considerations

- **Network required**: Graph API calls go to Microsoft's servers, so an internet connection is needed (unlike AppleScript which worked locally).
- **Latency**: Slightly slower than AppleScript — expect ~200-500ms per action.
- **Getting the selected message**: This is the main friction point. The "Copy Link" step adds complexity compared to AppleScript's direct access to the selected message.
- **Token expiry**: Tokens expire after ~1 hour, but MSAL handles refresh automatically if the cache is set up correctly.
- **Rate limits**: Microsoft Graph has rate limits, but for personal use this won't be an issue.
