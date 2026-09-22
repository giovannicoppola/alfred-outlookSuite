#!/usr/bin/env python3
"""Get a calendar event from the currently visible date range in Outlook.

Two modes:
  1. List mode (no args): Reads the date range from Outlook's calendar UI,
     fetches events via Graph API, outputs Alfred Script Filter JSON.
  2. Detail mode (event ID arg): Fetches that event by ID, outputs a
     markdown meeting note to stdout.
"""

import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime, timedelta, timezone


def _setup_paths():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_path = os.path.join(script_dir, "lib")
    for p in [lib_path, script_dir]:
        if p not in sys.path:
            sys.path.insert(0, p)


_setup_paths()

from auth import get_token, GRAPH_ENDPOINT, MissingConfigError, setup_required_items_json
import requests


# JXA to read the date range from Outlook's calendar toolbar
JXA_GET_DATE_RANGE = """
Application("Microsoft Outlook").activate();
delay(0.3);

var se = Application("System Events");
var outlook = se.processes["Microsoft Outlook"];
var win = outlook.windows[0];
var sg = win.splitterGroups[0];

// Find the static text with a year (date range label)
var dateRange = "";
try {
    var kids = sg.uiElements();
    for (var i = 0; i < kids.length; i++) {
        try {
            if (kids[i].role() === "AXStaticText") {
                var val = kids[i].value();
                if (val && val.match(/\\d{4}/)) {
                    dateRange = val;
                    break;
                }
            }
        } catch(e) {}
    }
} catch(e) {}
dateRange;
"""


def get_calendar_date_range():
    """Read the visible date range from Outlook's calendar toolbar."""
    result = subprocess.run(
        ["osascript", "-l", "JavaScript", "-e", JXA_GET_DATE_RANGE],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        return None
    return result.stdout.strip() or None


def parse_date_range(range_str):
    """Parse date range string into (start_date, end_date).

    Handles formats like:
      "March 9 - March 13, 2026 (Week 11)"    — week view
      "March 13, 2026"                          — day view
      "March 2026"                              — month view
    """
    # Strip week/month annotations
    range_str = re.sub(r'\s*\(Week \d+\)', '', range_str).strip()

    # Range: "March 9 - March 13, 2026"
    range_match = re.match(
        r'(\w+ \d+)\s*-\s*(\w+ \d+),?\s*(\d{4})', range_str
    )
    if range_match:
        start_str, end_str, year = range_match.groups()
        start = datetime.strptime(f"{start_str}, {year}", "%B %d, %Y").date()
        end = datetime.strptime(f"{end_str}, {year}", "%B %d, %Y").date()
        return start, end

    # Same-month range: "March 9 - 13, 2026"
    range_match = re.match(
        r'(\w+) (\d+)\s*-\s*(\d+),?\s*(\d{4})', range_str
    )
    if range_match:
        month, d1, d2, year = range_match.groups()
        start = datetime.strptime(f"{month} {d1}, {year}", "%B %d, %Y").date()
        end = datetime.strptime(f"{month} {d2}, {year}", "%B %d, %Y").date()
        return start, end

    # Single day: "March 13, 2026"
    day_match = re.match(r'(\w+ \d+,?\s*\d{4})', range_str)
    if day_match:
        dt = datetime.strptime(day_match.group(1).replace(',', ''), "%B %d %Y").date()
        return dt, dt

    # Month view: "March 2026"
    month_match = re.match(r'(\w+)\s+(\d{4})', range_str)
    if month_match:
        first = datetime.strptime(f"{month_match.group(1)} 1, {month_match.group(2)}", "%B %d, %Y").date()
        if first.month == 12:
            last = first.replace(year=first.year + 1, month=1, day=1) - timedelta(days=1)
        else:
            last = first.replace(month=first.month + 1, day=1) - timedelta(days=1)
        return first, last

    return None, None


def fetch_events(token, start_date, end_date):
    """Fetch calendar events for the given date range."""
    local_tz = datetime.now().astimezone().tzinfo
    start_dt = datetime.combine(start_date, datetime.min.time()).replace(tzinfo=local_tz)
    end_dt = datetime.combine(end_date, datetime.max.time()).replace(tzinfo=local_tz)

    start_utc = start_dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    end_utc = end_dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    response = requests.get(
        f"{GRAPH_ENDPOINT}/me/calendarView",
        headers={
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.timezone="UTC", outlook.body-content-type="text"',
        },
        params={
            "startDateTime": start_utc,
            "endDateTime": end_utc,
            "$select": "id,subject,start,end,attendees,body,location,organizer,isAllDay",
            "$orderby": "start/dateTime",
            "$top": 50,
        },
    )
    if response.status_code != 200:
        raise RuntimeError(f"Calendar fetch failed (HTTP {response.status_code}): {response.text[:200]}")
    return response.json().get("value", [])


def fetch_event_by_id(token, event_id):
    """Fetch a single event by ID."""
    response = requests.get(
        f"{GRAPH_ENDPOINT}/me/events/{event_id}",
        headers={
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.timezone="UTC", outlook.body-content-type="text"',
        },
        params={
            "$select": "id,subject,start,end,attendees,body,location,organizer,isAllDay",
        },
    )
    if response.status_code != 200:
        raise RuntimeError(f"Event fetch failed (HTTP {response.status_code}): {response.text[:200]}")
    return response.json()


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


def _parse_graph_dt(dt_str):
    """Parse a Graph API datetime string into the machine's local time.

    With Prefer: outlook.timezone="UTC", Graph returns naive UTC strings; we
    attach UTC and convert to local so times display correctly in any region."""
    # Graph API returns 7 fractional digits; Python < 3.11 only handles up to 6
    dt_str = re.sub(r'(\.\d{6})\d+', r'\1', dt_str)
    dt = datetime.fromisoformat(dt_str)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone()


def _normalize_name(name):
    """'Last, First' -> 'First Last'; anything else is returned trimmed.

    Graph often reports display names as 'Last, First'; the Obsidian daily-note
    headers ('# Meeting with [[First Last]]') use the natural order.
    """
    name = (name or "").strip()
    if ", " in name:
        last, first = name.split(", ", 1)
        return f"{first.strip()} {last.strip()}"
    return name


def _display_interval(date_stem):
    """Human 'Xy Xm Xd ' string between date_stem and now ('0d ' for today).

    date_stem is a daily-note stem like '2026-07-27-Mon'; returns '' if no
    leading YYYY-MM-DD can be parsed.
    """
    m = re.match(r"(\d{4}-\d{2}-\d{2})", date_stem or "")
    if not m:
        return ""
    try:
        d = datetime.strptime(m.group(1), "%Y-%m-%d")
    except ValueError:
        return ""
    secs = int(datetime.now().timestamp()) - int(d.timestamp())
    years = int(secs / 60 / 60 / 24 / 30.44 / 12)
    months = int(secs / 60 / 60 / 24 / 30.44 % 12)
    days = int(secs / 60 / 60 / 24 % 30.44)
    out = ""
    if years > 0:
        out += f"{years}y "
    if months > 0:
        out += f"{months}m "
    if days > 0:
        out += f"{days}d "
    if years == 0 and months == 0 and days == 0:
        out = "0d "
    return out


def _find_previous_meeting(attendee_name, vault):
    """Newest daily-note stem containing '# Meeting with [[attendee_name]]'.

    Returns e.g. '2026-07-27-Mon', or None. Daily notes are named
    YYYY-MM-DD-Day.md, so the lexically-greatest basename is the most recent.
    """
    grep = shutil.which("ggrep") or shutil.which("grep")
    if not grep or not attendee_name:
        return None
    needle = f"# Meeting with [[{attendee_name}]]"
    try:
        res = subprocess.run(
            [grep, "-Flr", needle, vault],
            capture_output=True, text=True, timeout=30,
        )
    except (subprocess.SubprocessError, OSError):
        return None
    files = [f for f in res.stdout.splitlines() if f.strip()]
    if not files:
        return None
    newest = max(files, key=lambda p: os.path.basename(p))
    return os.path.basename(newest)[:-3]  # strip '.md'


def _previous_meeting_line(attendees, organizer):
    """Build a 'Previous meeting …' trailer for the first non-self attendee.

    Empty string when MEETING_NOTES_VAULT is unset/invalid or nothing matches,
    so the public workflow (no vault configured) is unaffected. Only one grep
    is run — over the first attendee who isn't MYSELF.
    """
    vault = os.path.expanduser((os.getenv("MEETING_NOTES_VAULT") or "").strip())
    if not vault or not os.path.isdir(vault):
        return ""
    myself = _normalize_name(os.getenv("MYSELF", "")).lower()

    candidates = list(attendees or [])
    if organizer:
        candidates.append(organizer)
    for a in candidates:
        name = _normalize_name(a.get("emailAddress", {}).get("name", ""))
        if not name or name.lower() == myself:
            continue
        stem = _find_previous_meeting(name, vault)
        if stem:
            return (f"\nPrevious meeting ({_display_interval(stem)}ago): "
                    f"[[{stem}#Meeting with {name}]]\n")
        return ""  # first real attendee had no prior meeting; don't scan the rest
    return ""


def format_event_note(event):
    """Format event as a markdown meeting note."""
    subject = event.get("subject", "(no subject)")

    start_str = event["start"]["dateTime"]
    dt = _parse_graph_dt(start_str)
    date_formatted = dt.strftime("%A, %B %-d, %Y %-I:%M:%S %p")

    attendees = event.get("attendees", [])
    name_list = ", ".join(
        a.get("emailAddress", {}).get("name", a.get("emailAddress", {}).get("address", ""))
        for a in attendees
    )

    body_text = event.get("body", {}).get("content", "")
    agenda = _clean_body(body_text)

    note = f"# {subject}\n{date_formatted}\n\t{name_list}\n## Agenda\n{agenda}\n\n"
    return note + _previous_meeting_line(attendees, event.get("organizer"))


def format_alfred_items(events):
    """Format events as Alfred Script Filter JSON."""
    items = []
    for ev in events:
        subject = ev.get("subject", "(no subject)")
        start_str = ev["start"]["dateTime"]
        end_str = ev["end"]["dateTime"]
        dt_start = _parse_graph_dt(start_str)
        dt_end = _parse_graph_dt(end_str)

        is_all_day = ev.get("isAllDay", False)
        if is_all_day:
            time_display = "All day"
        else:
            time_display = f"{dt_start.strftime('%-I:%M %p')} - {dt_end.strftime('%-I:%M %p')}"

        day_display = dt_start.strftime("%a %b %-d")

        attendees = ev.get("attendees", [])
        attendee_count = len(attendees)
        attendee_str = f"{attendee_count} attendee{'s' if attendee_count != 1 else ''}" if attendees else ""

        location = ev.get("location", {}).get("displayName", "")
        sub_parts = [s for s in [time_display, location, attendee_str] if s]
        subtitle = f"{day_display}  |  " + "  |  ".join(sub_parts)

        items.append({
            "uid": ev["id"],
            "title": subject,
            "subtitle": subtitle,
            "arg": ev["id"],
        })

    if not items:
        items.append({
            "title": "No events found",
            "subtitle": "No events in the visible calendar date range",
            "valid": False,
        })

    return json.dumps({"items": items})


CACHE_DIR = os.getenv('alfred_workflow_cache', '/tmp')
EVENTS_CACHE_FILE = os.path.join(CACHE_DIR, 'calendar_events_cache.json')


def _load_events_cache():
    """Load cached events if the cache exists and the date range matches."""
    if not os.path.exists(EVENTS_CACHE_FILE):
        return None, None
    try:
        with open(EVENTS_CACHE_FILE, "r") as f:
            cache = json.load(f)
        return cache.get("date_range"), cache.get("events")
    except Exception:
        return None, None


def _save_events_cache(date_range, events):
    """Save events to cache keyed by date range."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(EVENTS_CACHE_FILE, "w") as f:
        json.dump({"date_range": date_range, "events": events}, f)


def _filter_events(events, query):
    """Filter events by query matching subject, location, or attendee names."""
    q_lower = query.lower()
    return [ev for ev in events
            if q_lower in ev.get("subject", "").lower()
            or q_lower in ev.get("location", {}).get("displayName", "").lower()
            or any(q_lower in a.get("emailAddress", {}).get("name", "").lower()
                   for a in ev.get("attendees", []))]


def _looks_like_event_id(q):
    """True if the query is a Graph event ID rather than a search phrase.

    IDs are long opaque tokens with no spaces; the exact prefix ("AAMk" etc.)
    varies by tenant, so match on shape instead of a fixed prefix."""
    return len(q) > 80 and " " not in q and re.match(r'^[A-Za-z0-9_\-=+/]+$', q) is not None


def main():
    query = sys.argv[1].strip() if len(sys.argv) > 1 else ""

    # Detail mode: argument is a Graph API event ID
    if _looks_like_event_id(query):
        token = get_token()
        event = fetch_event_by_id(token, query)
        if not event:
            sys.exit(1)
        print(format_event_note(event), end="")
        return

    # List mode: read date range from Outlook, list events
    # If there's a query and we have a cache, skip the JXA call entirely
    cached_range, cached_events = _load_events_cache()
    if query and cached_range and cached_events is not None:
        events = cached_events
    else:
        range_str = get_calendar_date_range()
        if not range_str:
            from datetime import date
            range_str = date.today().strftime("%B %-d, %Y")

        if cached_range == range_str and cached_events is not None:
            events = cached_events
        else:
            start_date, end_date = parse_date_range(range_str)
            if not start_date:
                sys.exit(1)
            token = get_token()
            events = fetch_events(token, start_date, end_date)
            _save_events_cache(range_str, events)

    if query:
        events = _filter_events(events, query)

    print(format_alfred_items(events))


if __name__ == "__main__":
    try:
        main()
    except MissingConfigError:
        print(setup_required_items_json())
    except Exception as e:
        print(json.dumps({"items": [{
            "title": "Couldn't reach Outlook / Microsoft Graph",
            "subtitle": str(e)[:120],
            "valid": False,
            "icon": {"path": "icon.png"},
        }]}))
