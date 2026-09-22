#!/usr/bin/env python3
"""Get current/upcoming meeting details via Graph API.

Takes a time interval (minutes) as argument. Returns events happening
between now and now + interval, formatted for the fetchMeetingDetails workflow.

Output format matches the AppleScript version:
  date – startTime-endTime|||---|||subject|||---|||attendees|||---|||[][][]...
"""

import os
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


def get_events(token, minutes_ahead):
    now = datetime.now(timezone.utc)
    end = now + timedelta(minutes=int(minutes_ahead))

    response = requests.get(
        f"{GRAPH_ENDPOINT}/me/calendarView",
        headers={
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.timezone="Eastern Standard Time"',
        },
        params={
            "startDateTime": now.isoformat(),
            "endDateTime": end.isoformat(),
            "$select": "subject,start,end,attendees",
            "$orderby": "start/dateTime",
            "$top": 50,
        },
    )
    if response.status_code != 200:
        print(f"Error {response.status_code}: {response.text}", file=sys.stderr)
        return []
    return response.json().get("value", [])


def format_event_time(dt_str):
    """Parse '2026-03-11T14:00:00.0000000' and return (date_str, time_str)."""
    dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    date_str = dt.strftime("%A, %B %-d, %Y")
    time_str = dt.strftime("%-I:%M:%S %p")
    return date_str, time_str


def main():
    minutes_ahead = sys.argv[1] if len(sys.argv) > 1 else "10"

    token = get_token()
    events = get_events(token, minutes_ahead)

    parts = []
    for ev in events:
        subject = ev.get("subject", "(no subject)")
        start_date, start_time = format_event_time(ev["start"]["dateTime"])
        _, end_time = format_event_time(ev["end"]["dateTime"])

        attendees = ev.get("attendees", [])
        name_list = ", ".join(
            a.get("emailAddress", {}).get("name", a.get("emailAddress", {}).get("address", ""))
            for a in attendees
        )

        parts.append(f"{start_date} – {start_time}-{end_time}|||---|||{subject}|||---|||{name_list}|||---|||")

    # Trailing [][][]  so split produces an empty last element,
    # matching the AppleScript output format
    if parts:
        print("[][][]".join(parts) + "[][][]")
    else:
        print("")


if __name__ == "__main__":
    main()
