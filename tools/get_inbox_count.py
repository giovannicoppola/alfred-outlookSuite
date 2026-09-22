#!/usr/bin/env python3
"""Get Outlook inbox count. Uses Graph API for new Outlook, AppleScript for legacy.

Prints just the count to stdout, suitable for use in shell scripts.
Exit code 0 on success, 1 on failure.
"""

import os
import subprocess
import sys

# This CLI utility lives outside the shipped workflow; resolve its deps
# (auth.py + bundled libs + config.json) from the sibling src/ directory.
_SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir, "src")
for _p in (os.path.join(_SRC, "lib"), _SRC):
    if _p not in sys.path:
        sys.path.insert(0, _p)

def get_count_graph_api():
    """Get inbox count via Microsoft Graph API."""
    try:
        from auth import get_token, GRAPH_ENDPOINT
        import requests
        token = get_token()
        response = requests.get(
            f"{GRAPH_ENDPOINT}/me/mailFolders/inbox",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "totalItemCount"},
        )
        if response.status_code == 200:
            return response.json().get("totalItemCount")
    except Exception:
        pass
    return None


def get_count_applescript():
    """Get inbox count via legacy AppleScript (old Outlook)."""
    result = subprocess.run(
        ["osascript", "-e", """
tell application "Microsoft Outlook"
    set mailAccount to first exchange account
    set InboxCount to (count messages in folder "Inbox" of mailAccount)
end tell"""],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        try:
            return int(result.stdout.strip())
        except ValueError:
            pass
    return None


def is_new_outlook():
    """Check if the new Outlook is running (AppleScript exchange account access fails)."""
    result = subprocess.run(
        ["osascript", "-e",
         'tell application "Microsoft Outlook" to get first exchange account'],
        capture_output=True, text=True
    )
    return result.returncode != 0


def main():
    if is_new_outlook():
        count = get_count_graph_api()
    else:
        count = get_count_applescript()

    if count is not None:
        print(count)
    else:
        print("ERROR: Could not get inbox count", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
