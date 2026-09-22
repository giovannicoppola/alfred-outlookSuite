#!/usr/bin/env python3
"""Save the currently selected email in Outlook as an .eml file."""

import os
import re
import sys

# This CLI utility lives outside the shipped workflow; resolve its deps
# (auth.py, get_selected.py + bundled libs + config.json) from sibling src/.
_SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir, "src")
for _p in (os.path.join(_SRC, "lib"), _SRC):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import requests
from auth import get_token, GRAPH_ENDPOINT
from get_selected import get_selected_info, find_message

SAVE_DIR = os.path.expanduser("~/Documents/SavedEmails")


def save_as_eml(token, message_id, subject, received_date):
    """Download raw MIME content and save as .eml file."""
    response = requests.get(
        f"{GRAPH_ENDPOINT}/me/messages/{message_id}/$value",
        headers={"Authorization": f"Bearer {token}"},
    )
    if response.status_code != 200:
        print(f"Error downloading MIME content (HTTP {response.status_code}): {response.text}", file=sys.stderr)
        return None

    # Build a safe filename from date + subject
    safe_subject = re.sub(r'[^\w\s-]', '', subject)[:80].strip()
    date_prefix = received_date[:10] if received_date else "unknown"
    filename = f"{date_prefix}_{safe_subject}.eml"

    os.makedirs(SAVE_DIR, exist_ok=True)
    filepath = os.path.join(SAVE_DIR, filename)

    with open(filepath, "wb") as f:
        f.write(response.content)

    return filepath


def main():
    info = get_selected_info()
    if not info or (not info.get("subject") and not info.get("listDesc")):
        print("No email selected in Outlook.")
        return

    subject = info.get("subject", "")
    sender_email = info.get("senderEmail", "")
    timestamp = info.get("timestamp", "")
    list_desc = info.get("listDesc", "")

    if not subject:
        match = re.search(r"Subject:\s*(.+?)(?:,\s{4}|\s{4})", list_desc)
        if match:
            subject = match.group(1).strip().rstrip(",")

    if not subject:
        print("Could not determine subject.")
        return

    token = get_token()
    msg = find_message(token, subject, sender_email, timestamp)
    if not msg:
        print("Could not find this message via the Graph API.")
        return

    sender = msg.get("from", {}).get("emailAddress", {})
    print(f"Saving as .eml:")
    print(f"  Subject: {msg['subject']}")
    print(f"  From:    {sender.get('name', '')} <{sender.get('address', '')}>")

    filepath = save_as_eml(token, msg["id"], msg["subject"], msg["receivedDateTime"])
    if filepath:
        print(f"  Saved:   {filepath}")
    else:
        print("  Failed to save.")


if __name__ == "__main__":
    main()
