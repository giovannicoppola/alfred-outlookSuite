# Dependency Analysis: Can We Distribute Without Hammerspoon?

## Context

The current workflow design has two layers of functionality:
1. **Graph API actions** — archive, delete, flag, move, snooze, save emails, create drafts, read calendar
2. **Vim-style navigation + RSVP** — keyboard-driven message list navigation and meeting response shortcuts

The original plan assumed AppleScript could simulate "Copy Link" in Outlook to get the message ID. However, **new Outlook for Mac does not support AppleScript**, which changes the equation.

## Layer 1: Graph API Actions

### Verdict: No Hammerspoon needed

These actions follow the pattern: get message ID from clipboard → call Graph API via Python. The only macOS-specific piece is triggering Outlook's "Copy Link" to get the message ID onto the clipboard.

### Solution without new tools

1. **User assigns a macOS keyboard shortcut** to Outlook's "Copy Link" menu item (one-time setup: System Settings > Keyboard > Keyboard Shortcuts > App Shortcuts).
2. **Alfred hotkey** fires → sends that keystroke to Outlook → waits ~200ms → reads clipboard → passes to Python script.
3. **Python script** extracts message ID, authenticates via MSAL, calls Graph API.

For calendar actions and draft creation, there is no "selected message" step — those call the API directly.

### What users need to install

- Alfred (they already have it)
- Python 3 (ships with macOS or easily available)
- `pip install msal requests`
- An Azure AD app registration (tenant-specific)

No Hammerspoon, no AppleScript, no additional tools.

## Layer 2: Vim-Style Navigation + RSVP

### Verdict: Hammerspoon (or equivalent) is required

These shortcuts perform real-time Accessibility (AX) framework automation:
- Walking the AX tree to find the message list (`AXOutline` in legacy, `AXTable` in new Outlook)
- Sending keystrokes to a specific UI element (not just globally)
- Finding and clicking RSVP buttons by AX role and description
- Detecting legacy vs. new Outlook via AX tree inspection
- Handling modal key sequences like `gg` (two-key vim combos)

### Alternatives evaluated

| Approach | Can it replace Hammerspoon? | Notes |
|---|---|---|
| **`pyobjc`** (Python package) | Partially | Can access AX APIs via `ApplicationServices` framework. Could replicate AX tree walking and element interaction. But global hotkeys and modal key sequences (`gg`, `G`) are much harder — requires building a persistent daemon with `CGEventTap` or `NSEvent.addGlobalMonitorForEvents`. Essentially rebuilding Hammerspoon in Python. |
| **AppleScript / `osascript`** | No | New Outlook does not support AppleScript. `System Events` can send keystrokes but cannot walk AX trees with the precision needed for RSVP button detection or element-specific interaction. |
| **macOS Shortcuts app** | No | Can run scripts and do basic automation but has no AX tree access. |
| **`cliclick`** (CLI tool) | No | Only does mouse/keyboard simulation, no AX awareness. |
| **Swift/ObjC command-line tool** | Yes | Could write a small Swift CLI using the Accessibility framework. But users would need to compile or install that binary — same distribution problem as Hammerspoon. |
| **Alfred built-in** | No | Alfred can send keystrokes and run scripts but cannot interact with the AX tree. |

### Why there is no pure-Python solution

The AX-heavy features require:
- **Global hotkey interception** scoped to a specific app (Outlook)
- **AX tree traversal** to locate UI elements by role and description
- **Element-targeted actions** (pressing a specific button, not just sending a global keystroke)
- **Modal key sequences** (`gg` = two consecutive presses of `g`)

While `pyobjc` can technically access the Accessibility and CoreGraphics frameworks, building a reliable daemon that handles all of the above would be a significant engineering effort and would be harder to maintain than the existing Hammerspoon Lua code.

## Summary

| Feature set | Dependencies | Hammerspoon required? |
|---|---|---|
| Graph API actions (archive, delete, flag, move, snooze, save, calendar, drafts) | Alfred + Python + macOS keyboard shortcut | No |
| Vim-style message list navigation | Hammerspoon (or equivalent AX tool) | Yes |
| RSVP shortcuts (accept/tentative/decline silently) | Hammerspoon (or equivalent AX tool) | Yes |

## Recommendation for distribution

- Distribute the Graph API workflow as a standalone Alfred workflow with Python scripts. Zero extra tools required beyond Alfred and Python.
- Document Hammerspoon as an **optional dependency**: "Install Hammerspoon if you want vim-style navigation and RSVP shortcuts; the core email management works without it."
- Provide the Hammerspoon Lua modules as a separate, optional add-on that users can drop into their `~/.hammerspoon/` directory.
