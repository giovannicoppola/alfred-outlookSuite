"""
Fast snapshot of the currently selected email's identity, for the snooze flow.

Runs (synchronously) while the Alfred snooze date picker is showing — Alfred
holds focus, so the Outlook selection is frozen and correct. snoozer.py later
resolves + moves that cached identity, so the intended email is snoozed even if
the selection drifts after Enter.

Deliberately lightweight: it reads only the selected message-list row via
Hammerspoon (~0.05s) and imports nothing heavy (no requests/msal/auth or the
slow JXA accessibility walk), so it doesn't lag the date picker. For new Outlook
the row description is all get_selected_info() yields anyway — snoozer.py parses
the subject from it and resolves the message via the Graph API.

Outputs nothing to stdout (pass-through for the Alfred chain).
"""

import os
import json
import time
import subprocess

CACHE_FILE = os.path.join(os.getenv('alfred_workflow_data', '/tmp'), 'snooze_selection_cache.json')
HS_BIN = "/Applications/Hammerspoon.app/Contents/Frameworks/hs/hs"


def _hs_get_selection():
    """Selected message-list row description via Hammerspoon. '' if unavailable
    (e.g. legacy Outlook, where snoozer.py uses AppleScript and ignores this)."""
    if not os.path.exists(HS_BIN):
        return ""
    try:
        r = subprocess.run(
            [HS_BIN, "-c", "return require('outlook_rsvp').getSelection()"],
            capture_output=True, text=True, timeout=15,
        )
        return r.stdout.strip() if r.returncode == 0 else ""
    except Exception:
        return ""


def main():
    desc = _hs_get_selection()
    if not desc:
        return  # no selection (or legacy Outlook) — leave any prior cache alone

    cache = {
        "listDesc": desc,
        "subject": "",
        "senderEmail": "",
        "timestamp": "",
        "ts": time.time(),
    }
    os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
    with open(CACHE_FILE, "w") as f:
        json.dump(cache, f)


if __name__ == '__main__':
    main()
