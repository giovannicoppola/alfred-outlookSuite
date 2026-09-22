#!/usr/bin/env python3
"""RSVP to the selected Outlook meeting invitation via Hammerspoon.

Two modes:
  python3 rsvp.py                 -> emit the Alfred Script Filter menu
  python3 rsvp.py --run <action>  -> perform the action via Hammerspoon

Actions are "verb:mode" pairs, e.g. "accept:silent". The heavy lifting
(finding the invite in Outlook's reading pane and clicking the right RSVP
button) lives in outlook_rsvp.lua, driven here through the `hs` CLI. This
feature is New-Outlook-only and requires Hammerspoon; see README/SETUP.
"""

import json
import os
import subprocess
import sys

HS = "/Applications/Hammerspoon.app/Contents/Frameworks/hs/hs"

# (arg, title, subtitle) — arg is "verb:mode" consumed by --run below.
ACTIONS = [
    ("accept:silent", "Accept", "Accept without sending a response"),
    ("accept:send", "Accept (notify)", "Accept and send a response"),
    ("tentative:silent", "Tentative", "Tentatively accept without sending a response"),
    ("tentative:send", "Tentative (notify)", "Tentatively accept and send a response"),
    ("decline:silent", "Decline", "Decline without sending a response"),
    ("decline:send", "Decline (notify)", "Decline and send a response"),
    ("remove:silent", "Remove from calendar", "Remove a cancelled / no-response invite"),
]


def emit_menu():
    items = [
        {"title": title, "subtitle": sub, "arg": arg, "icon": {"path": "icon.png"}}
        for arg, title, sub in ACTIONS
    ]
    print(json.dumps({"items": items}))


def run_action(action):
    if not os.path.exists(HS):
        print(
            "Hammerspoon is not installed (or the `hs` CLI is missing).\n"
            "RSVP shortcuts require Hammerspoon — see the workflow README.",
            file=sys.stderr,
        )
        sys.exit(1)

    verb, _, mode = action.partition(":")
    mode = mode or "silent"
    if verb == "remove":
        lua = "require('outlook_rsvp').run('remove')"
    else:
        lua = f"require('outlook_rsvp').run('{verb}', '{mode}')"

    result = subprocess.run([HS, "-c", lua], capture_output=True, text=True)
    if result.returncode != 0:
        sys.stderr.write(result.stderr or "RSVP action failed.\n")
        sys.exit(result.returncode)
    if result.stdout.strip():
        print(result.stdout.strip())


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--run":
        run_action(sys.argv[2])
    else:
        emit_menu()


if __name__ == "__main__":
    main()
