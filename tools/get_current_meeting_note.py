#!/usr/bin/env python3
"""Get the current meeting details formatted as a markdown note.

Output format:
# Meeting Title
Wednesday, March 11, 2026 2:00:00 PM
	Attendee 1, Attendee 2, ...
## Agenda
Body text (stripped of Teams boilerplate)

Used by the "Create Meeting Notes" Alfred workflow.
"""

import os
import re
import sys
from datetime import datetime, timedelta, timezone

# This CLI utility lives outside the shipped workflow; resolve its deps
# (auth.py + bundled libs + config.json) from the sibling src/ directory.
_SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.pardir, "src")
for _p in (os.path.join(_SRC, "lib"), _SRC):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import requests
from auth import get_token, GRAPH_ENDPOINT


def get_current_event(token, minutes_ahead=5):
    now = datetime.now(timezone.utc)
    end = now + timedelta(minutes=int(minutes_ahead))

    response = requests.get(
        f"{GRAPH_ENDPOINT}/me/calendarView",
        headers={
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.timezone="Eastern Standard Time", outlook.body-content-type="text"',
        },
        params={
            "startDateTime": now.isoformat(),
            "endDateTime": end.isoformat(),
            "$select": "subject,start,end,attendees,body",
            "$orderby": "start/dateTime",
            "$top": 10,
        },
    )
    if response.status_code != 200:
        print(f"Error {response.status_code}: {response.text}", file=sys.stderr)
        return []
    return response.json().get("value", [])


def _clean_body(body_text):
    """Clean up event body text: fix links, remove Teams boilerplate."""
    # Convert [​icon] filename<URL> to markdown [filename](URL)
    body_text = re.sub(
        r'\[\u200b\w+ icon\]\s*([^<\n]+)<(https?://[^>]+)>',
        r'[\1](\2)',
        body_text
    )
    # Convert [https://...icon.svg] filename<URL> to markdown [filename](URL)
    body_text = re.sub(
        r'\[https?://[^\]]+\]\s*([^<\n]+)<(https?://[^>]+)>',
        r'[\1](\2)',
        body_text
    )
    # Remove [cid:...] inline image references
    body_text = re.sub(r'\[cid:[^\]]+\]', '', body_text)
    # Remove Teams boilerplate (meeting join info, dial-in, etc.)
    lines = body_text.split("\n")
    filtered = []
    for line in lines:
        if "Microsoft Teams" in line and ("meeting" in line.lower() or "Need help" in line):
            break
        filtered.append(line)
    return "\n".join(filtered).rstrip()


def format_event_note(event):
    """Format a single event as a markdown meeting note."""
    subject = event.get("subject", "(no subject)")

    start_str = event["start"]["dateTime"]
    # Graph returns 7-digit fractional seconds (e.g. ...:00.0000000), which
    # datetime.fromisoformat can't parse — trim the fraction to 6 digits.
    start_str = re.sub(r"(\.\d{6})\d+", r"\1", start_str.replace("Z", "+00:00"))
    dt = datetime.fromisoformat(start_str)
    date_formatted = dt.strftime("%A, %B %-d, %Y %-I:%M:%S %p")

    attendees = event.get("attendees", [])
    name_list = ", ".join(
        a.get("emailAddress", {}).get("name", a.get("emailAddress", {}).get("address", ""))
        for a in attendees
    )

    body_text = event.get("body", {}).get("content", "")
    agenda = _clean_body(body_text)

    note = f"# {subject}\n{date_formatted}\n\t{name_list}\n## Agenda\n{agenda}\n\n"
    return note


def main():
    minutes_ahead = sys.argv[1] if len(sys.argv) > 1 else "5"

    token = get_token()
    events = get_current_event(token, minutes_ahead)

    if not events:
        print("", end="")
        sys.exit(0)

    if len(events) == 1:
        print(format_event_note(events[0]), end="")
    else:
        # Multiple events — print count on stderr, output first one
        print(f"Multiple events ({len(events)}), using first", file=sys.stderr)
        print(format_event_note(events[0]), end="")


if __name__ == "__main__":
    main()
