#!/usr/bin/env python3
"""Get the currently selected email in Outlook via accessibility API + Graph API.

Uses two sources from the Outlook UI:
  1. Message list row (selected) → subject, sender name, time
  2. Reading pane header → exact sender email, subject

These are combined to query the Graph API and find the exact message.
"""

import json
import os
import re
import subprocess
import sys
from datetime import datetime, timedelta, timezone

import requests
from auth import get_token, GRAPH_ENDPOINT

HS_BIN = "/Applications/Hammerspoon.app/Contents/Frameworks/hs/hs"


def _hs_get_selection():
    """Read the selected message-list row description via Hammerspoon.

    New Outlook renders the message list inside a web-content accessibility
    subtree that a plain osascript/JXA walk can't reach. Hammerspoon can (native
    AX API + AXManualAccessibility), and it already holds Accessibility
    permission — so we shell out to it. This also means Alfred itself doesn't
    need Accessibility permission. Returns '' if Hammerspoon/module unavailable.
    """
    if not os.path.exists(HS_BIN):
        return ""
    try:
        r = subprocess.run(
            [HS_BIN, "-c", "return require('outlook_rsvp').getSelection()"],
            capture_output=True, text=True, timeout=15,
        )
        return r.stdout.strip() if r.returncode == 0 else ""
    except Exception:
        return ""

# JXA script: locate the selected email's info by SEARCHING Outlook's
# accessibility tree by role/description, rather than by hardcoded element
# indices (which break whenever the window layout shifts).
#
# It flattens the tree in depth-first order, pruning heavy subtrees (folder
# sidebar, message-list body, email body) and decoration leaves (buttons,
# images), and stops the instant the timestamp after the From header is seen.
# It derives the fields from stable landmarks in the reading-pane header:
#   - senderEmail: element with desc "messageHeaderFromContent"
#   - subject:     nearest preceding AXStaticText with a value
#   - timestamp:   first following value that looks like a time
JXA_GET_SELECTED = r"""
function getRole(el)  { try { return el.role(); } catch(e) { return ""; } }
function getDesc(el)  { try { var d = el.description(); return d === null ? "" : d; } catch(e) { return ""; } }
function getVal(el)   { try { var v = el.value(); return (v === null || v === undefined) ? "" : String(v); } catch(e) { return ""; } }

// Each accessibility property read is an Apple Events round-trip, so the walk
// prunes subtrees that hold nothing we need and stops the instant it has the
// From header plus its trailing timestamp.
var flat = [];        // {role, desc, val} in DFS order, header region only
var listDesc = "";    // selected message-list row description (for snoozer)
var MAX_DEPTH = 14;
var STOP = {};        // sentinel thrown to unwind the recursion early
var sawFrom = false;  // set once the "From" header has been visited

// Leaf/decoration roles we never need. Pruning these removes most of the
// Apple Events round-trips the walk would pay: the ~50 toolbar AXButtons, and
// the 8-level self-nested AXSearchField subtree at the top of the window.
// (Header fields are AXStaticText, so skipping AXTextField is safe.)
var SKIP = {AXButton: 1, AXImage: 1, AXDisclosureTriangle: 1, AXScrollBar: 1,
            AXSplitter: 1, AXValueIndicator: 1, AXOutline: 1, AXWebArea: 1,
            AXTextField: 1};

function walk(el, depth) {
    if (depth > MAX_DEPTH) return;
    var role = getRole(el);

    if (SKIP[role]) return;

    // Message-list body: don't descend into all 364 rows. Read only the
    // SELECTED row (one AXSelectedRows call) to get its description — snoozer
    // uses this to detect "Conversation," and snooze the whole thread.
    if (role === "AXTable") {
        try {
            var sel = el.selectedRows();
            if (sel && sel.length) {
                try { listDesc = sel[0].uiElements()[0].description(); }
                catch(e) { listDesc = getDesc(sel[0]); }
            }
        } catch(e) {}
        return;
    }

    var desc = getDesc(el);
    var val = getVal(el);
    flat.push({role: role, desc: desc, val: val});

    if (desc === "messageHeaderFromContent") sawFrom = true;
    // The timestamp is the first time-like value AFTER the From header.
    // Capture it and unwind — everything past it (recipients, body) is noise.
    else if (sawFrom && /\d{1,2}:\d{2}\s*[AP]M/.test(val)) throw STOP;

    var kids;
    try { kids = el.uiElements(); } catch(e) { return; }
    for (var i = 0; i < kids.length; i++) walk(kids[i], depth + 1);
}

var se = Application("System Events");
var outlook = se.processes["Microsoft Outlook"];
var subject = "", senderEmail = "", timestamp = "";

try {
    walk(outlook.windows[0], 0);
} catch(e) { if (e !== STOP) { /* ignore */ } }

// Sender email — from the "messageHeaderFromContent" landmark.
var fromIdx = -1;
for (var i = 0; i < flat.length; i++) {
    if (flat[i].desc === "messageHeaderFromContent") { fromIdx = i; break; }
}
if (fromIdx >= 0) {
    var m = flat[fromIdx].val.match(/[\w.+-]+@[\w.-]+\.[\w.-]+/);
    if (m) senderEmail = m[0];

    // Subject — nearest preceding AXStaticText that carries a value.
    for (var j = fromIdx - 1; j >= 0; j--) {
        if (flat[j].role === "AXStaticText" && flat[j].val) { subject = flat[j].val; break; }
    }

    // Timestamp — first following value that looks like a clock time.
    for (var k = fromIdx; k < flat.length; k++) {
        if (flat[k].val && /\d{1,2}:\d{2}\s*[AP]M/.test(flat[k].val)) { timestamp = flat[k].val; break; }
    }
}

JSON.stringify({listDesc: listDesc, subject: subject, senderEmail: senderEmail, timestamp: timestamp});
"""


def get_selected_info():
    """Use JXA to read info about the selected email from Outlook's UI."""
    result = subprocess.run(
        ["osascript", "-l", "JavaScript", "-e", JXA_GET_SELECTED],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"JXA error: {result.stderr.strip()}", file=sys.stderr)
        info = None
    else:
        try:
            info = json.loads(result.stdout.strip())
        except json.JSONDecodeError:
            print(f"Could not parse JXA output: {result.stdout[:200]}", file=sys.stderr)
            info = None

    # New Outlook: the JXA walk can't reach the web-content message list, so it
    # comes back empty. Fall back to Hammerspoon, which can read the selection.
    if not info or not (info.get("subject") or info.get("senderEmail") or info.get("listDesc")):
        desc = _hs_get_selection()
        if desc:
            info = {"listDesc": desc, "subject": "", "senderEmail": "", "timestamp": ""}
    return info


def parse_ui_timestamp(ts_str):
    """Parse Outlook UI timestamps into a local naive datetime. Handles:
      'Today at 2:47 PM', 'Yesterday at 10:30 AM',
      'Mon 3/9/2026 at 10:30 AM', '3/9/2026 at 10:30 AM',
      'Wednesday, July 8, 2026 at 11:25 AM'."""
    if not ts_str:
        return None

    # Extract the time portion (e.g., "2:47 PM")
    time_match = re.search(r"(\d{1,2}:\d{2}\s*[AP]M)", ts_str)
    if not time_match:
        return None
    time_str = time_match.group(1)

    # Determine the date portion
    now = datetime.now()
    ts_lower = ts_str.lower()
    if "today" in ts_lower:
        date_part = now.date()
    elif "yesterday" in ts_lower:
        date_part = (now - timedelta(days=1)).date()
    else:
        # Try a numeric date like "3/9/2026" or "Mon 3/9/2026"
        date_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", ts_str)
        if date_match:
            date_part = datetime.strptime(date_match.group(1), "%m/%d/%Y").date()
        else:
            # Try a written date like "July 8, 2026"
            month_match = re.search(r"([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})", ts_str)
            if month_match:
                try:
                    date_part = datetime.strptime(
                        f"{month_match.group(1)} {month_match.group(2)} {month_match.group(3)}",
                        "%B %d %Y",
                    ).date()
                except ValueError:
                    return None
            else:
                return None

    # Combine date + time in local time
    time_part = datetime.strptime(time_str.strip(), "%I:%M %p").time()
    local_dt = datetime.combine(date_part, time_part)
    return local_dt


def pick_best_match(messages, ui_timestamp_str):
    """From a list of candidate messages, pick the one closest to the UI timestamp."""
    if not messages:
        return None
    if len(messages) == 1:
        return messages[0]

    ui_dt = parse_ui_timestamp(ui_timestamp_str)
    if not ui_dt:
        # Can't parse timestamp, return the most recent
        return messages[0]

    best = None
    best_diff = None
    for msg in messages:
        received = msg.get("receivedDateTime", "")
        try:
            # Graph returns UTC times like "2026-03-10T18:47:08Z"
            msg_utc = datetime.fromisoformat(received.replace("Z", "+00:00"))
            # Convert to local naive for comparison (UI shows local time)
            msg_local = msg_utc.astimezone().replace(tzinfo=None)
            diff = abs((msg_local - ui_dt).total_seconds())
            if best_diff is None or diff < best_diff:
                best_diff = diff
                best = msg
        except (ValueError, TypeError):
            continue

    return best or messages[0]


def find_message(token, subject, sender_email, ui_timestamp_str):
    """Find the exact message via Graph API using subject + sender + timestamp."""
    headers = {"Authorization": f"Bearer {token}"}

    # Escape single quotes in subject for OData filter
    safe_subject = subject.replace("'", "''")

    # Try exact filter: subject + sender
    params = {
        "$top": 10,
        "$select": "id,subject,from,receivedDateTime,conversationId,isRead,internetMessageId",
        "$orderby": "receivedDateTime desc",
    }

    if sender_email:
        params["$filter"] = f"subject eq '{safe_subject}' and from/emailAddress/address eq '{sender_email}'"
    else:
        params["$filter"] = f"subject eq '{safe_subject}'"

    response = requests.get(
        f"{GRAPH_ENDPOINT}/me/messages", headers=headers, params=params
    )

    if response.status_code == 200:
        messages = response.json().get("value", [])
        if messages:
            return pick_best_match(messages, ui_timestamp_str)

    # Fallback: search by subject only (handles special characters)
    response = requests.get(
        f"{GRAPH_ENDPOINT}/me/messages",
        headers=headers,
        params={
            "$search": f'"subject:{subject}"',
            "$top": 10,
            "$select": "id,subject,from,receivedDateTime,conversationId,isRead,internetMessageId",
        },
    )
    if response.status_code != 200:
        print(f"API error {response.status_code}: {response.text}", file=sys.stderr)
        return None

    messages = response.json().get("value", [])
    # If we have sender_email, filter results client-side
    if sender_email:
        messages = [
            m for m in messages
            if m.get("from", {}).get("emailAddress", {}).get("address", "").lower() == sender_email.lower()
        ]
    return pick_best_match(messages, ui_timestamp_str)


def main():
    info = get_selected_info()
    if not info or (not info.get("subject") and not info.get("listDesc")):
        print("No email selected in Outlook.")
        return

    subject = info.get("subject", "")
    sender_email = info.get("senderEmail", "")
    list_desc = info.get("listDesc", "")
    timestamp = info.get("timestamp", "")

    # If reading pane subject is empty, parse from list description
    if not subject:
        match = re.search(r"Subject:\s*(.+?)(?:,\s{4}|\s{4})", list_desc)
        if match:
            subject = match.group(1).strip().rstrip(",")

    if not subject:
        print(f"Could not determine subject.")
        print(f"  List description: {list_desc[:120]}...")
        return

    print(f"Subject:  {subject}")
    print(f"Sender:   {sender_email or '(unknown)'}")
    print(f"Time:     {timestamp or '(unknown)'}")

    token = get_token()
    msg = find_message(token, subject, sender_email, timestamp)
    if not msg:
        print("\nCould not find this message via the Graph API.")
        return

    sender = msg.get("from", {}).get("emailAddress", {})
    print(f"\nFound message:")
    print(f"  ID:         {msg['id'][:50]}...")
    print(f"  Subject:    {msg['subject']}")
    print(f"  From:       {sender.get('name', '')} <{sender.get('address', '')}>")
    print(f"  Received:   {msg['receivedDateTime']}")
    print(f"  Read:       {msg['isRead']}")
    print(f"  Message-ID: {msg.get('internetMessageId', 'N/A')}")


if __name__ == "__main__":
    main()
