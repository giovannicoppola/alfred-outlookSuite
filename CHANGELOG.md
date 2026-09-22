# Changelog

All notable changes to this workflow are documented here.

## [0.10.0] — New Outlook (Microsoft Graph) support

This release makes the workflow work with the **New (Electron-based) Outlook for Mac**, which dropped AppleScript. Scripts auto-detect your Outlook version and use the Microsoft Graph API on New Outlook, falling back to the legacy AppleScript/SQLite path on classic Outlook.

### Added
- Full New-Outlook support via Microsoft Graph: search, snooze/unsnooze, save, draft, calendar/event notes, and RSVP (via Hammerspoon).
- One-time browser sign-in using OAuth2 + PKCE. No client secret required — register a **public-client** Entra app (see `ADMIN_ACCESS.md` for a plain-English IT request).
- Clear in-Alfred prompt pointing to setup when sign-in isn't configured yet, instead of an empty result.
- Snoozed folder is created automatically if it doesn't exist.
- **Vim-style message-list navigation** (`j`/`k`/`h`/`l`, `gg`/`G`) via a bundled Hammerspoon module (`outlook_keys.lua`), installed by the `olhs` setup. Works on both Legacy and New Outlook.
- **Optional RSVP keyboard shortcuts** (⌘⇧A/T/D/X for accept/tentative/decline/remove, active only while Outlook is frontmost) — enable with `require("outlook_rsvp").bindDefaultHotkeys()` in `init.lua`.

### Changed
- Bundled dependencies are now **pure-Python** — runs on the stock macOS `/usr/bin/python3`, on both Apple Silicon and Intel. Removed `cryptography`/`cffi`/`pycparser` (unused by the public-client flow) and mismatched compiled binaries.
- Calendar events display in your **local timezone** (previously fixed to Eastern).
- All workflow script actions standardized on `/usr/bin/python3`.
- Documentation rewritten for the public-client auth model.

### Fixed
- Snoozed emails can no longer be stranded in the Snoozed folder if an unsnooze move fails — entries are pruned only after a confirmed move, and the store is written atomically.
- Leap-day / year-boundary crash in the snooze-count display.
- Contact autocomplete no longer dead-ends a search when its data source is unavailable.
- Calendar/Graph API errors surface a message instead of appearing as "no events found."
