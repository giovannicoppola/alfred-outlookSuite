# Alfred Workflow Commands for Outlook Graph API

> **Layout note (v0.10.0):** there is **no venv** — every script runs under system
> `/usr/bin/python3` and imports `msal`/`requests` from the bundled `src/lib/`.
> The scripts split into two locations:
> - **`src/`** — scripts wired into the Alfred workflow (`main.py`, `snoozer.py`,
>   `saveEmail.py`, `captureSelected.py`, `draftEmail.py`, `get_selected_event.py`,
>   `get_selected.py`, `rsvp.py`, `auth.py`, …).
> - **`tools/`** — standalone CLI utilities **not** bundled into the workflow, kept
>   for use by other workflows (`get_inbox_count.py`, `get_meeting_details.py`,
>   `get_current_meeting_note.py`, `inbox_count.py`, `list_folders.py`,
>   `save_eml.py`, plus the `demo_graph_api.py` / `test_connection.py` demos). These
>   add `../src` and `../src/lib` to `sys.path` so they resolve `auth.py` + libs.
>
> In the examples below, `venv/bin/python3` should read as `/usr/bin/python3`, and
> the `tools/` utilities live under `alfred-outlook/tools/` (not `src/`).

---

## Mail Operations

### 1. Get Inbox Count

Returns just the number (e.g. `1175`) to stdout. Auto-detects new vs old Outlook.

```bash
OUTLOOK_SHORTCUTS_DIR="$HOME/path/to/alfred-outlook/src"
"$OUTLOOK_SHORTCUTS_DIR/venv/bin/python3" "$OUTLOOK_SHORTCUTS_DIR/get_inbox_count.py"
```

**Output:** `1175`

### 2. Get Selected Email Info

Reads the currently selected email in Outlook (via accessibility API) and prints its details.

```bash
OUTLOOK_SHORTCUTS_DIR="$HOME/path/to/alfred-outlook/src"
"$OUTLOOK_SHORTCUTS_DIR/venv/bin/python3" "$OUTLOOK_SHORTCUTS_DIR/get_selected.py"
```

**Output:**
```
Subject:  Re: Weekly update
Sender:   john.smith@company.com
Time:     Today at 2:47 PM

Found message:
  ID:         AAMkADJi...
  Subject:    Re: Weekly update
  From:       John Smith <john.smith@company.com>
  Received:   2026-03-11T18:47:08Z
  Read:       True
  Message-ID: <abc123@company.com>
```

### 3. Snooze Selected Email/Conversation

Moves the selected email (or entire conversation thread) to the "Snoozed" folder.

```bash
OUTLOOK_SHORTCUTS_DIR="$HOME/path/to/alfred-outlook/src"
"$OUTLOOK_SHORTCUTS_DIR/venv/bin/python3" "$OUTLOOK_SHORTCUTS_DIR/snooze.py"
```

**Output:**
```
Moving to Snoozed:
  Subject: GenomeWeb Daily
  From:    GenomeWeb <news@genomeweb.com>
Done!
```

Or for a conversation:
```
Moving conversation (3 messages) to Snoozed:
  Subject: Weekly update
  Moved: John Smith (2026-03-11T14:47)
  Moved: Jane Doe (2026-03-11T15:02)
  Moved: John Smith (2026-03-11T16:30)
Done! Moved 3/3 messages.
```

### 4. Save Selected Email as .eml (Two-Step Flow)

Saves the selected email as `.eml` to a user-chosen folder and archives it. This is a two-step process in the `alfred-outlook` workflow:

**Step 1: Capture selected email** (runs at keyword/hotkey trigger, BEFORE the folder picker)

```bash
# In the alfred-outlook workflow src/ directory
python3 captureSelected.py
```

This caches the selected email's Graph API message ID to `selected_email_cache.json`.

**Step 2: Save and archive** (runs after user picks destination folder)

```bash
# In the alfred-outlook workflow src/ directory
python3 saveEmail.py "/path/to/destination/folder"
```

Reads the cached message ID, downloads the raw MIME content, saves as `.eml`, and archives the message.

**Output (stdout):** `/path/to/destination/folder/2026-03-11_Important document_AAMkADJi....eml`

**Alfred wiring:**
- Keyword `ols` → Run Script `python3 captureSelected.py` → File Filter → Run Script `python3 saveEmail.py "{query}"`
- Hotkey → **(2-second delay)** → Run Script `python3 captureSelected.py` → File Filter → Run Script `python3 saveEmail.py "{query}"`

**Note:** Hotkey triggers require a 2-second delay before `captureSelected.py`. Without it, the accessibility API may not detect the selected email due to a focus-switch race condition. Keyword triggers need no delay.

### 5. Create Email Draft

Creates a new email draft with the given subject line. Auto-detects new vs old Outlook.

```bash
# In the alfred-outlook workflow src/ directory
python3 draftEmail.py "Subject line here"
```

**Legacy Outlook:** Creates the draft via AppleScript and opens it in Outlook.

**New Outlook:** Creates the draft via Graph API (`POST /me/messages`). Prints the `webLink` URL to stdout — chain with `open "{query}"` to open the draft in the browser.

### 6. Open Email from Search Results

**Legacy Outlook:** Search results return a local file path as `arg`. Open directly:

```bash
open "{query}"
```

**New Outlook:** Search results return a `webLink` URL as `arg`. Opens in the default browser (Outlook Web):

```bash
open "{query}"
```

The same `open` command works for both — it handles file paths and URLs. There is no way to programmatically open a specific message in the new Outlook desktop app.

---

## Calendar Operations

### 7. Get Current/Upcoming Meeting Details

Returns events happening within the next N minutes, formatted with `|||---|||` and `[][][]` delimiters (for fetchMeetingDetails workflow parsing).

```bash
OUTLOOK_SHORTCUTS_DIR="$HOME/path/to/alfred-outlook/src"
"$OUTLOOK_SHORTCUTS_DIR/venv/bin/python3" "$OUTLOOK_SHORTCUTS_DIR/get_meeting_details.py" 10
```

**Argument:** minutes ahead (default: `10`)

**Output:**
```
Wednesday, March 11, 2026 – 2:00:00 PM-3:00:00 PM|||---|||Team Standup|||---|||John Smith, Jane Doe|||---|||[][][]
```

### 8. Get Current Meeting Note (Markdown)

Returns the current/upcoming meeting as a markdown note with body text. Used by the "Create Meeting Notes from Obsidian" path.

```bash
OUTLOOK_SHORTCUTS_DIR="$HOME/path/to/alfred-outlook/src"
"$OUTLOOK_SHORTCUTS_DIR/venv/bin/python3" "$OUTLOOK_SHORTCUTS_DIR/get_current_meeting_note.py" 5
```

**Argument:** minutes ahead (default: `5`)

**Output:**
```markdown
# Team Standup
Wednesday, March 11, 2026 2:00:00 PM
	John Smith, Jane Doe
## Agenda
Discuss progress on Q1 goals...

```

### 9. Browse Calendar Events + Get Event Note (Markdown)

Lists events for the currently visible date range in Outlook's calendar, then outputs a markdown meeting note for the selected event. Works from any Outlook view (calendar shows that view's date range, other views default to today).

**Step 1 — Script Filter** (list events):
```bash
# In the alfred-outlook workflow src/ directory
python3 get_selected_event.py "{query}"
```

Reads the visible date range from Outlook's calendar toolbar (e.g. "March 9 - March 13, 2026"), fetches events via Graph API, and outputs Alfred Script Filter JSON. The user can type to filter by subject, location, or attendee name. Results are cached per date range so filtering is instant (~0.2s).

**Step 2 — Run Script** (get event details):
```bash
python3 get_selected_event.py "{query}"
```

When the user selects an event, `{query}` is the Graph API event ID (starts with `AAMk`). The script fetches the full event and outputs a markdown meeting note.

**Output (Step 2):**
```markdown
# Team Standup
Wednesday, March 11, 2026 2:00:00 PM
	John Smith, Jane Doe
## Agenda
Discuss progress on Q1 goals...
```

**Note:** Replaces the old popover-based approach. The calendar popover dismisses when Outlook loses focus, making it unreliable. The new approach lists all events for the visible date range and lets the user pick.

---

## Backward-Compatible Shell Wrappers

These try the Graph API first, then fall back to legacy AppleScript. Use these if you want to support both new and old Outlook.

### Get Current Meeting Note (with fallback)

```bash
#!/bin/zsh
OUTLOOK_SHORTCUTS_DIR="$HOME/path/to/alfred-outlook/src"
GRAPH_PYTHON="$OUTLOOK_SHORTCUTS_DIR/venv/bin/python3"
GRAPH_SCRIPT="$OUTLOOK_SHORTCUTS_DIR/get_current_meeting_note.py"

meeting_note=""
if [ -f "$GRAPH_PYTHON" ] && [ -f "$GRAPH_SCRIPT" ]; then
    meeting_note=$("$GRAPH_PYTHON" "$GRAPH_SCRIPT" 5 2>/dev/null)
fi

if [ -n "$meeting_note" ]; then
    echo -n "$meeting_note"
    exit 0
fi

# Fallback: legacy AppleScript for old Outlook
osascript -e '
tell application "Microsoft Outlook"
    ...old AppleScript here...
end tell'
```

### Browse Calendar Events (no fallback needed)

The new `get_selected_event.py` (in `alfred-outlook/src/`) uses the Graph API directly and doesn't need a fallback wrapper. It runs as an Alfred Script Filter:

```bash
python3 get_selected_event.py "{query}"
```

See command 9 above for details.

---

## Quick Reference

| Action | Script | Args | Stdout |
|---|---|---|---|
| Inbox count | `get_inbox_count.py` | — | number |
| Selected email info | `get_selected.py` | — | text summary |
| Snooze email/thread | `snooze.py` | — | status text |
| Capture selected email | `captureSelected.py` (in alfred-outlook) | — | — (writes cache file) |
| Save as .eml + archive | `saveEmail.py` (in alfred-outlook) | destination folder | file path |
| Create email draft | `draftEmail.py` (in alfred-outlook) | subject line | webLink URL (new) or — (legacy) |
| Open email | `open "{query}"` | webLink URL or file path | — |
| Meeting details | `get_meeting_details.py` | minutes (default 10) | delimited string |
| Current meeting note | `get_current_meeting_note.py` | minutes (default 5) | markdown |
| Browse calendar events | `get_selected_event.py` (in alfred-outlook) | query (optional filter) | Alfred JSON |
| Get event note | `get_selected_event.py` (in alfred-outlook) | event ID (`AAMk...`) | markdown |

All commands use the same venv Python and share `auth.py` for Graph API authentication (token is cached in `token_cache.json`). First run will open a browser for OAuth login.

---

## Caching

The mail search (`main.py`) caches the 50 most recent messages from the Graph API in `alfred_workflow_cache/graph_messages_cache.json`. This makes filter-only queries (e.g., `has:attach`, `is:unread`, `since:7`, `folder:Inbox`) instant after the first load.

**Configuration:** Set the `CACHE_MINUTES` workflow environment variable (default: `30`). Set to `0` to disable.

Queries that require server-side KQL search (`from:`, `to:`, `cc:`, `subject:`, bare text keywords) always hit the API and are not cached.
