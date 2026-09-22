#!/usr/bin/env python3
"""List all mail folders and their IDs."""

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
        f"{GRAPH_ENDPOINT}/me/mailFolders",
        headers={"Authorization": f"Bearer {token}"},
        params={"$top": 100, "$select": "id,displayName,totalItemCount"},
    )
    if response.status_code != 200:
        print(f"Error {response.status_code}: {response.text}")
        return

    folders = response.json().get("value", [])
    print(f"{'Folder':<35} {'Count':>6}  ID")
    print("-" * 100)
    for f in folders:
        print(f"{f['displayName']:<35} {f['totalItemCount']:>6}  {f['id'][:40]}...")


if __name__ == "__main__":
    main()
