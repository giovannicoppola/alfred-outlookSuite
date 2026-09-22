#!/usr/bin/env python3
"""Count the number of emails in the Inbox."""

import os
import sys

# This CLI utility lives outside the shipped workflow; resolve its deps
# (auth.py + bundled libs + config.json) from the sibling src/ directory.
_SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir, "src")
for _p in (os.path.join(_SRC, "lib"), _SRC):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import requests
from auth import get_token, GRAPH_ENDPOINT


def main():
    token = get_token()
    response = requests.get(
        f"{GRAPH_ENDPOINT}/me/mailFolders/inbox",
        headers={"Authorization": f"Bearer {token}"},
        params={"$select": "displayName,totalItemCount,unreadItemCount"},
    )
    if response.status_code != 200:
        print(f"Error {response.status_code}: {response.text}")
        return

    folder = response.json()
    total = folder.get("totalItemCount", 0)
    unread = folder.get("unreadItemCount", 0)
    print(f"Inbox: {total} emails ({unread} unread)")


if __name__ == "__main__":
    main()
