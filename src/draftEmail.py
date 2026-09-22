"""
Create a new email draft with the given subject.
Supports both legacy (AppleScript) and new (Graph API) Outlook.

Usage: python3 draftEmail.py "Subject line here"

New Outlook: creates the draft via Graph API and saves it silently in the Drafts
folder (it is NOT opened), then posts a "Draft saved" notification. This is a
quick way to jot down an email to finish writing later.
Legacy Outlook: creates the draft via AppleScript.
"""

import sys
import os
from subprocess import check_output, run, PIPE
from util import log


def _notify(title, text):
    """Post a macOS notification. Args are passed via argv to avoid any quoting
    issues with the subject text."""
    try:
        run(['osascript', '-e',
             'on run {t, m}\ndisplay notification m with title t\nend run',
             title, text])
    except Exception:
        pass


def is_new_outlook():
    try:
        check_output(
            ['osascript', '-e', 'tell application "Microsoft Outlook" to get first exchange account'],
            stderr=PIPE
        ).decode('utf-8').strip()
        return False
    except Exception:
        return True


def draft_via_graph(subject):
    """Create a draft email via Graph API and open it in the browser."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_path = os.path.join(script_dir, "lib")
    for p in [lib_path, script_dir]:
        if p not in sys.path:
            sys.path.insert(0, p)

    from auth import get_token, GRAPH_ENDPOINT
    import requests

    token = get_token()
    resp = requests.post(
        f"{GRAPH_ENDPOINT}/me/messages",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"subject": subject, "isDraft": True},
    )

    if resp.status_code != 201:
        log(f"Error creating draft (HTTP {resp.status_code}): {resp.text[:200]}")
        _notify("Outlook draft failed", f"HTTP {resp.status_code}")
        return None

    log(f"Created draft: {subject}")
    # The draft is saved silently in the Drafts folder (not opened) — this is a
    # way to jot down a quick email to write later. Confirm with a notification.
    _notify("Draft saved to Outlook", subject)
    return None


def draft_via_applescript(subject):
    """Create a draft email via AppleScript."""
    scpt = f'''
on run
    tell application "Microsoft Outlook"
        set new_message to make new outgoing message with properties {{subject: "{subject}"}}
    end tell
end run
'''
    check_output(['osascript', '-e', scpt])


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 draftEmail.py \"Subject line\"", file=sys.stderr)
        sys.exit(1)

    subject = sys.argv[1]

    if is_new_outlook():
        web_link = draft_via_graph(subject)
        if web_link:
            print(web_link)
    else:
        draft_via_applescript(subject)


if __name__ == '__main__':
    main()
