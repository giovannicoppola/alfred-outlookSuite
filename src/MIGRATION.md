# alfred-outlook: New Outlook Migration Guide

This document describes the changes made to support the new Outlook for Mac, which dropped AppleScript and SQLite database access. All scripts now auto-detect which Outlook is running and use the appropriate backend.

---

## Detection Method

All scripts use the same check:

```python
def is_new_outlook():
    try:
        check_output(['osascript', '-e',
            'tell application "Microsoft Outlook" to get first exchange account'],
            stderr=PIPE)
        return False  # Legacy Outlook
    except Exception:
        return True   # New Outlook
```

The new Outlook throws an error on `first exchange account` because it doesn't support AppleScript exchange account objects.

---

## Architecture: Legacy vs New Outlook

| Component | Legacy (AppleScript + SQLite) | New (Microsoft Graph API) |
|---|---|---|
| Mail search | SQL queries on `Outlook.sqlite` | Graph API `$search` (KQL) + `$filter` (OData) |
| Message ID | `Record_RecordID` (local, changes across machines) | `internetMessageId` (RFC Message-ID, stable) |
| Folder lookup | `Record_FolderID` → `Folders` table | `parentFolderId` → `/me/mailFolders` endpoint |
| Open message | File path to `.olk` message file | `webLink` (opens in Outlook web/app) |
| Thread view | `Message_ThreadTopic` column | `conversationId` from Graph API |
| Move message | AppleScript `move aMessage to folder` | Graph API `POST /messages/{id}/move` |
| Inbox count | AppleScript `count messages in folder "Inbox"` | Graph API `/me/mailFolders/inbox` → `totalItemCount` |
| Selected email | AppleScript `selection` | JXA accessibility API + Graph API search |
| Save email | AppleScript `save` + Finder move | Graph API `GET /messages/{id}/$value` (MIME download) |
| Open email | `open` local `.olk` file path | `open` `webLink` URL (opens in browser) |
| Authentication | None (local database) | OAuth2 via MSAL (token cached in `token_cache.json`) |
| Dependencies | None (built-in `sqlite3`) | `msal`, `requests` (bundled in `lib/`) |

---

## Modified Scripts

### `main.py` — Mail Search

**Legacy path:** `compileSQL()` → builds SQL WHERE clause → `handle()` queries SQLite → outputs Alfred JSON with file paths as `arg`.

**New path:** `compileGraphQuery()` → builds KQL `$search` / OData `$filter` → `handleGraph()` queries Graph API → `_format_graph_results()` outputs Alfred JSON with `webLink` as `arg`.

#### Search Syntax Mapping

| User syntax | Legacy (SQL) | New (Graph API) |
|---|---|---|
| `from:name` | `Message_SenderList LIKE '%name%'` | `$search` `"from:name"` |
| `from:me` | Uses `MYSELF` env var | Uses `MYSELF` env var in KQL |
| `to:name` | `Message_RecipientList LIKE '%name%'` | `$search` `"to:name"` |
| `cc:name` | `Message_CCRecipientAddressList LIKE` | `$search` `"cc:name"` |
| `subject:word` | `Message_NormalizedSubject LIKE` | `$search` `"subject:word"` |
| `has:attach` | `Message_HasAttachment = 1` | `$filter` `hasAttachments eq true` |
| `is:read` | `Message_ReadFlag = 1` | `$filter` `isRead eq true` |
| `is:unread` | `Message_ReadFlag = 0` | `$filter` `isRead eq false` |
| `is:important` | `Record_Priority = 1` | `$filter` `importance eq 'high'` |
| `is:unimportant` | `Record_Priority = 5` | `$filter` `importance eq 'low'` |
| `since:N` / `since:Nw` / `since:Nm` | `Message_TimeReceived > epoch` | `$filter` `receivedDateTime ge datetime` |
| `folder:name` | `Record_FolderID = id` | Queries `/me/mailFolders/{id}/messages` |
| `account:name` | `Record_AccountUID = id` | Not supported (Graph API is single-account) |
| `--a` (sort ascending) | `ORDER BY ... ASC` | `$orderby receivedDateTime asc` or client-side sort |
| `-word` (exclude) | `NOT LIKE '%word%'` | Client-side filtering on `subject` + `bodyPreview` |
| bare words | `Message_Preview/Subject LIKE` | `$search` `"word"` |
| `sq:name` (saved query) | Same (UI-only, expands to other operators) | Same |
| Contact autocomplete | Same (`showContacts`) | Same |

#### Graph API Limitations

- **`$search` + `$filter` cannot be combined.** When both are needed, `$search` runs first and `$filter` conditions are applied client-side.
- **`$search` returns results by relevance**, not date. Results are re-sorted by `receivedDateTime` client-side.
- **Excluded folders** (from workflow config `EXCLUDED_FOLDERS`) are filtered client-side for the "no query" listing.
- **Max 50 results** per query (configurable via `$top`).

#### Output Differences

| Field | Legacy | New |
|---|---|---|
| `arg` (main action) | File path to message on disk | `webLink` URL (opens in Outlook web/app) |
| Ctrl modifier (`arg`) | `Message_ThreadTopic` string | `subject` string (used to find `conversationId`) |
| Shift modifier (`arg`) | Preview text from `Message_Preview` | Preview text from `bodyPreview` |

### Thread View (handled in `main.py`)

**Legacy:** SQL query on `Message_ThreadTopic`.
**New:** `main.py` (`handleGraph(is_thread=True)`) searches by subject to find `conversationId`, then queries all messages with that `conversationId`. Sorted ascending by date. (The old standalone `thread.py` was removed — the thread view now lives entirely in `main.py`.)

### `snoozer.py` — Snooze Email

**Legacy:** AppleScript `selection` → moves to "Snoozed" folder → looks up `Message_MessageID` from SQLite → stores in `snoozer.json`.
**New:** JXA accessibility API → Graph API to find message → Graph API `move` to "Snoozed" folder → stores `internetMessageId` in `snoozer.json`.

Both paths store `{RFC_Message_ID: "YYYY-MM-DD"}` in `snoozer.json`. The key is the same value (`Message_MessageID` in SQLite = `internetMessageId` in Graph API), so entries are cross-compatible.

Conversation snooze: detects "Conversation," in the accessibility description, fetches all inbox messages with the same `conversationId`, moves all of them.

### `unSnoozer.py` — Unsnooze Email

**Legacy:** Reads `snoozer.json` → finds entries with date <= today → looks up `Record_RecordID` from SQLite → AppleScript moves from "Snoozed" to "Inbox" → posts to Beeminder if configured.
**New:** Reads `snoozer.json` → finds entries with date <= today → Graph API searches "Snoozed" folder by `internetMessageId` → Graph API `move` to Inbox → logs to file. Beeminder posting **is** supported on the Graph API path via `post_to_beeminder()` (runs when `BEEMINDER` is enabled). Entries are pruned from `snoozer.json` only after a successful move, so a failed move leaves them tracked for the next run.

### `saveEmail.py` — Save Email as .eml + Archive

**Legacy:** Embedded AppleScript reads the `selection`, saves as `.eml` to a staging folder, moves the file to the chosen destination, and archives the message via AppleScript `move`.

**New:** Two-step flow using `captureSelected.py` + `saveEmail.py`:

1. **`captureSelected.py`** runs at keyword/hotkey trigger time (BEFORE the folder picker appears). It uses the JXA accessibility API to detect the selected email, finds it via Graph API, and caches the message ID + metadata to `selected_email_cache.json` in `alfred_workflow_data`.
2. **`saveEmail.py`** runs after the user picks a destination folder. It reads the cached message ID, downloads the raw MIME content via `GET /me/messages/{id}/$value`, saves as `.eml`, and archives via `POST /me/messages/{id}/move`.

This two-step approach is necessary because when Alfred's file picker appears, Outlook's reading pane clears and the accessibility API can no longer detect the selected email.

**Alfred wiring:**
- Keyword `ols` → Run Script `python3 captureSelected.py` → File Filter → Run Script `python3 saveEmail.py "{query}"`
- Hotkey → (2-second delay) → Run Script `python3 captureSelected.py` → File Filter → Run Script `python3 saveEmail.py "{query}"`

**Note:** Hotkey triggers require a **2-second delay** before `captureSelected.py` runs, to allow Outlook to settle after the focus switch. Keyword triggers need no delay.

### `get_selected_event.py` — Browse Calendar Events / Get Event Note

**Legacy:** Not available (calendar events were accessed via AppleScript in other workflows).

**New:** Two-mode script for the Graph API:
1. **List mode** (no args or search query): Activates Outlook to read the visible calendar date range from the toolbar. Fetches events via `GET /me/calendarView`. Outputs Alfred Script Filter JSON. Results are cached per date range — filtering by subject, location, or attendee name is instant (~0.2s).
2. **Detail mode** (argument is a Graph event ID): Fetches a single event by ID via `GET /me/events/{id}`. Outputs a markdown meeting note. (The ID is detected by shape — a long opaque token — rather than a fixed `AAMk` prefix, which varies by tenant.)

Replaces the old popover-based calendar reader, which was unreliable because the calendar popover dismisses when Outlook loses focus.

**Body text cleanup** (also applied in `get_current_meeting_note.py` and `fetchAgenda.py` in alfred-almanac):
- `[​pptx icon] file<URL>` → `[file](URL)` (proper markdown link)
- `[https://...svg] file<URL>` → `[file](URL)` (CDN icon variant)
- `[cid:image@...]` → removed (inline image references)
- Microsoft Teams meeting boilerplate (join links, dial-in, video conferencing) → removed

### `draftEmail.py` — Create Email Draft

**Legacy:** AppleScript `make new outgoing message with properties {subject: ...}` — opens a new draft in the Outlook desktop app.

**New:** `POST /me/messages` with `{subject: ..., isDraft: true}` via Graph API. Returns the `webLink` URL to stdout, which can be opened in the browser to edit the draft.

### Opening Emails

**Legacy:** The `arg` output from search results is a file path to the `.olk` message file on disk. Alfred opens it directly with `open "{query}"`, which launches the message in the Outlook desktop app.

**New:** The `arg` output is a `webLink` URL from the Graph API. There is no way to programmatically open a specific message in the new Outlook desktop app. The action uses `open "{query}"` which opens the `webLink` in the default browser, showing the message in Outlook Web.

### Scripts NOT Modified (no changes needed)

- **`consts.py`** — Configuration constants (env vars, file paths). Works as-is.
- **`util.py`** — Utility functions. SQLite-dependent functions (`fetchRecordID`, `fetchEmailID`, etc.) are only called by the legacy path.
- **`countSnoozed.py`** — Reads `snoozer.json` only. Works as-is.
- **`countSnoozedDates.py`** — Reads `snoozer.json` only. Works as-is.
- **`outlook-rebuild.py`** — Rebuilds SQLite folder/contact caches. Only relevant for legacy path. Harmless if called on new Outlook (will fail silently on missing database).

---

## Dependencies

The Graph API path requires `msal` and `requests`. These are bundled in the `lib/` folder within the workflow:

```
src/lib/
├── msal/
├── requests/
├── certifi/
├── charset_normalizer/
├── idna/
├── jwt/ (PyJWT)
└── urllib3/
```

(`cryptography`, `cffi`, and `pycparser` were removed: they were only needed for the old confidential-client certificate/secret flow, which the public-client + PKCE auth no longer uses. Dropping them also removes the only compiled binaries, so `lib/` is now pure Python and works on any Python 3.9+ and on both Apple Silicon and Intel.)

Installed via:
```bash
pip install --target=src/lib msal requests
```

The scripts add `lib/` to `sys.path` at runtime:
```python
lib_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib")
sys.path.insert(0, lib_path)
```

The Graph API authentication module (`auth.py`) and email detection module (`get_selected.py`) live alongside the other scripts in `src/` and are added to `sys.path` at runtime.

---

## Authentication

The Graph API requires OAuth2 authentication via a Microsoft Entra app registration (a **public client** using PKCE — no client secret). See `SETUP.md` for full setup instructions.

Key files (in `src/`):
- `config.json` — client ID, tenant ID (gitignored; no secret)
- `token_cache.json` — cached OAuth tokens (gitignored, auto-generated)
- `auth.py` — shared authentication module (`msal.PublicClientApplication`)

First run opens a browser for Microsoft login. Subsequent runs use cached tokens silently.

---

## Caching (New Outlook)

The Graph API path includes a message cache to speed up repeated queries. On first run (or when the cache expires), the 50 most recent messages are fetched from the API and stored locally in `alfred_workflow_cache/graph_messages_cache.json` along with the folder map and a timestamp.

### Configuration

- **`CACHE_MINUTES`** — Workflow environment variable. How long the cache is valid, in minutes. Default: `30`. Set to `0` to effectively disable caching.

### What uses the cache

Queries that don't require server-side `$search` (KQL) are resolved entirely from the cache with client-side filtering:

| Query type | Example | Uses cache? |
|---|---|---|
| No query (recent messages) | *(empty)* | Yes |
| Filter: attachments | `has:attach` | Yes |
| Filter: read/unread | `is:read`, `is:unread` | Yes |
| Filter: importance | `is:important`, `is:unimportant` | Yes |
| Filter: date range | `since:7`, `since:2w` | Yes |
| Filter: folder | `folder:Inbox` | Yes |
| Exclusion | `-newsletter` | Yes |
| Sort order | `--a` | Yes |
| Search: sender/recipient | `from:name`, `to:name`, `cc:name` | No (needs API) |
| Search: subject | `subject:word` | No (needs API) |
| Search: free text | `keyword` | No (needs API) |
| Thread view | Ctrl modifier | No (needs API) |

### Cache lifecycle

1. **First query** — Cache miss. Fetches from API, saves to cache.
2. **Subsequent queries within TTL** — Filter-only queries are served from cache instantly. Search queries always hit the API but do not invalidate the cache.
3. **After TTL expires** — Next filter-only or no-query request fetches fresh data from the API and updates the cache.

### Limitations

- The cache stores at most 50 messages (the API's `$top` limit). Filter-only queries on the cache can only return results from these 50 messages. For example, `has:attach` will only find attachments among the 50 most recent messages, not the entire mailbox.
- The cache does not reflect changes made after caching (new messages, read status changes, moves). These appear after the cache expires.

---

## Known Limitations (New Outlook)

1. **`account:` filter** — Not supported. Graph API operates in a single-user context.
2. **Opening messages** — Uses `webLink` which opens in Outlook web or the app, rather than opening a local file directly.
3. **Tentative calendar events** — May not appear in `calendarView` API results (known Graph API behavior for recurring event instances).
4. **Search relevance vs date** — `$search` returns by relevance. Results are re-sorted by date client-side, but the top 50 by relevance may not include all chronologically recent results.
5. **LibreSSL warning** — `urllib3` emits a `NotOpenSSLWarning` on macOS with LibreSSL. This is cosmetic and does not affect functionality.
6. **Beeminder integration** — Supported on the Graph API path via `post_to_beeminder()` in `unSnoozer.py` (posts the inbox count when `BEEMINDER` is enabled). Optional and off by default.
7. **Save email hotkey delay** — When triggered via hotkey, `captureSelected.py` requires a 2-second delay in Alfred before execution. Without the delay, the accessibility API may fail to detect the selected email due to a focus-switch race condition. Keyword triggers do not need this delay.
8. **Save email two-step flow** — The email selection must be captured before Alfred's file picker appears, because the file picker causes Outlook's reading pane to clear. This requires `captureSelected.py` to run as a separate step before the file filter.
