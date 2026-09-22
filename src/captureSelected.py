"""
Capture the currently selected email info and cache it for later use.
Run this BEFORE the folder selector appears (right after keyword trigger).

Outputs nothing to stdout (pass-through for Alfred chain).
Saves email info to a temp file that saveEmail.py reads.
"""

import sys
import os
import json
from subprocess import check_output, PIPE
from util import log

CACHE_FILE = os.path.join(os.getenv('alfred_workflow_data', '/tmp'), 'selected_email_cache.json')


def is_new_outlook():
    try:
        check_output(
            ['osascript', '-e', 'tell application "Microsoft Outlook" to get first exchange account'],
            stderr=PIPE
        ).decode('utf-8').strip()
        return False
    except Exception:
        return True


def capture_graph():
    """Capture selected email info via JXA + Graph API, save to cache."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_path = os.path.join(script_dir, "lib")
    for p in [lib_path, script_dir]:
        if p not in sys.path:
            sys.path.insert(0, p)

    import re
    from auth import get_token, GRAPH_ENDPOINT
    from get_selected import get_selected_info, find_message
    import requests

    info = get_selected_info()
    if not info or (not info.get("subject") and not info.get("listDesc")):
        log("No email selected in Outlook.")
        return False

    subject = info.get("subject", "")
    sender_email = info.get("senderEmail", "")
    timestamp = info.get("timestamp", "")
    list_desc = info.get("listDesc", "")

    if not subject:
        # Fields in the row description are separated by a comma + 2+ spaces
        # (e.g. "..., Subject: Foo,  Meeting message,  ..."). Stop the subject at
        # that delimiter, or at end of string. (Subject commas use a single space.)
        match = re.search(r"Subject:\s*(.+?)(?:,\s{2,}|$)", list_desc)
        if match:
            subject = match.group(1).strip().rstrip(",")

    if not subject:
        log("Could not determine subject.")
        return False

    token = get_token()
    msg = find_message(token, subject, sender_email, timestamp)
    if not msg:
        log("Could not find this message via the Graph API.")
        return False

    # Cache the message ID and metadata
    cache = {
        "id": msg["id"],
        "subject": msg.get("subject", ""),
        "receivedDateTime": msg.get("receivedDateTime", ""),
        "from": msg.get("from", {}),
    }

    os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
    with open(CACHE_FILE, "w") as f:
        json.dump(cache, f)

    log(f"Cached selected email: {cache['subject']}")
    return True


def main():
    if is_new_outlook():
        capture_graph()
    # For legacy Outlook, no caching needed — AppleScript reads selection at save time


if __name__ == '__main__':
    main()
