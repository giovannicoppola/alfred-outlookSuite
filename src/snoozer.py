"""
A new version of the OL SNOOZER script, to be included in the alfred-OutlookSuite
Sunny ☀️   🌡️+70°F (feels +70°F, 59%) 🌬️↓2mph 🌕&m Fri Jun  2 08:42:48 2023
W22Q2 – 153 ➡️ 211 – 22 ❇️ 343
"""

"""
# Script updated on Mon Apr 11 10:17:50 2022  – Sunny ☀️   🌡️+48°F (feels +45°F, 39%) 🌬️↙9mph 🌔
# to use exchange id instead of database ID, so that it will be possible to use it across computers.
# Partly cloudy ⛅️  🌡️+69°F (feels +69°F, 41%) 🌬️↘13mph 🌔 Tue Apr 12 17:03:32 2022
# tried to fetch the message ID from the sqlite database. To my knowledge is the only way to obtain a message ID which is shared across computers. The unique id I have been using is not.
# it seems there is no way to get this via applescript

# also, removed the categories (labels) part which I don't use for now. These are in v01 of this script, in the same folder.
"""
from subprocess import Popen, PIPE, check_output
import sys
import time
from util import log, fetchRecordID
import json
from consts import *

SNOOZE_CACHE_FILE = os.path.join(os.getenv('alfred_workflow_data', '/tmp'), 'snooze_selection_cache.json')
SNOOZE_CACHE_FRESH_SECONDS = 120


def _read_snapshot():
    """Return the email identity captured by captureSnooze.py while the date
    picker was open, or None if there's no fresh snapshot. Using the snapshot
    (rather than the live selection read after Enter) ensures the intended email
    is snoozed even if the selection drifted afterward."""
    try:
        with open(SNOOZE_CACHE_FILE, "r") as f:
            cache = json.load(f)
    except Exception:
        return None
    if (time.time() - float(cache.get("ts", 0))) >= SNOOZE_CACHE_FRESH_SECONDS:
        return None
    return {
        "subject": cache.get("subject", ""),
        "senderEmail": cache.get("senderEmail", ""),
        "timestamp": cache.get("timestamp", ""),
        "listDesc": cache.get("listDesc", ""),
    }


def is_new_outlook():
    """Check if the new Outlook is running (no AppleScript exchange account support)."""
    try:
        result = check_output(
            ['osascript', '-e', 'tell application "Microsoft Outlook" to get first exchange account'],
            stderr=PIPE
        ).decode('utf-8').strip()
        return False
    except Exception:
        return True


def _write_snoozes(data):
    """Write the snooze store atomically so an interrupted write can't corrupt
    it (a corrupt store loses every pending snooze)."""
    tmp = OUTLOOK_SNOOZER_FILE + ".tmp"
    with open(tmp, "w") as f:
        json.dump(data, f, indent=4)
    os.replace(tmp, OUTLOOK_SNOOZER_FILE)


def snooze_via_graph(snooze_date):
    """Snooze the selected email(s) via Microsoft Graph API."""
    # Add bundled lib/ for msal/requests, and src/ for auth/get_selected
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_path = os.path.join(script_dir, "lib")
    for p in [lib_path, script_dir]:
        if p not in sys.path:
            sys.path.insert(0, p)

    from auth import get_token, GRAPH_ENDPOINT
    from get_selected import get_selected_info, find_message
    import re
    import requests

    # Prefer the snapshot captured (synchronously) while Alfred's date picker was
    # open, when the selection was frozen. Fall back to a live read only if no
    # fresh snapshot exists (capture skipped/failed) — that read can race a
    # post-Enter selection change, which is the bug the snapshot avoids.
    info = _read_snapshot()
    if not info or (not info.get("subject") and not info.get("listDesc")):
        info = get_selected_info()
    if not info or (not info.get("subject") and not info.get("listDesc")):
        log("No email selected in Outlook.")
        return None

    subject = info.get("subject", "")
    sender_email = info.get("senderEmail", "")
    timestamp = info.get("timestamp", "")
    list_desc = info.get("listDesc", "")

    if not subject:
        # Row fields are separated by comma + 2+ spaces; stop the subject there
        # (or at end of string). Subject commas themselves use a single space.
        match = re.search(r"Subject:\s*(.+?)(?:,\s{2,}|$)", list_desc)
        if match:
            subject = match.group(1).strip().rstrip(",")

    if not subject:
        log("Could not determine subject.")
        return None

    token = get_token()

    # Determine if this is a conversation or single message
    is_conversation = "Conversation," in list_desc

    msg = find_message(token, subject, sender_email, timestamp)
    if not msg:
        log("Could not find this message via the Graph API.")
        return None

    # Find Snoozed folder
    resp = requests.get(
        f"{GRAPH_ENDPOINT}/me/mailFolders",
        headers={"Authorization": f"Bearer {token}"},
        params={"$filter": "displayName eq 'Snoozed'", "$select": "id"},
    )
    folders = resp.json().get("value", []) if resp.status_code == 200 else []
    if folders:
        snoozed_folder_id = folders[0]["id"]
    else:
        # Auto-create the Snoozed folder so snoozing works out of the box
        # instead of silently doing nothing when it doesn't exist.
        create_resp = requests.post(
            f"{GRAPH_ENDPOINT}/me/mailFolders",
            headers={"Authorization": f"Bearer {token}"},
            json={"displayName": "Snoozed"},
        )
        if create_resp.status_code not in (200, 201):
            log(f"Snoozed folder not found and could not be created (HTTP {create_resp.status_code}).")
            return None
        snoozed_folder_id = create_resp.json()["id"]

    # Collect messages to move
    messages_to_move = []
    if is_conversation:
        conv_id = msg.get("conversationId")
        if conv_id:
            resp = requests.get(
                f"{GRAPH_ENDPOINT}/me/mailFolders/inbox/messages",
                headers={"Authorization": f"Bearer {token}"},
                params={
                    "$filter": f"conversationId eq '{conv_id}'",
                    "$select": "id,subject,from,receivedDateTime,internetMessageId",
                    "$top": 50,
                },
            )
            if resp.status_code == 200:
                messages_to_move = resp.json().get("value", [])
    if not messages_to_move:
        messages_to_move = [msg]

    # Move and collect internetMessageIds. A moved-but-untracked message would
    # sit in Snoozed forever (never resurfaced), so guarantee we capture an
    # internetMessageId for every successful move.
    message_ids = []
    for m in messages_to_move:
        move_resp = requests.post(
            f"{GRAPH_ENDPOINT}/me/messages/{m['id']}/move",
            headers={"Authorization": f"Bearer {token}"},
            json={"destinationId": snoozed_folder_id},
        )
        if move_resp.status_code != 201:
            log(f"Failed to move message {m.get('id')} to Snoozed (HTTP {move_resp.status_code}).")
            continue
        try:
            moved = move_resp.json()
        except Exception:
            moved = {}
        inet_id = m.get("internetMessageId") or moved.get("internetMessageId")
        if not inet_id:
            # Last resort: fetch it from the moved message so it stays tracked.
            new_id = moved.get("id") or m.get("id")
            g = requests.get(
                f"{GRAPH_ENDPOINT}/me/messages/{new_id}",
                headers={"Authorization": f"Bearer {token}"},
                params={"$select": "internetMessageId"},
            )
            if g.status_code == 200:
                inet_id = g.json().get("internetMessageId")
        if inet_id:
            message_ids.append(inet_id)
        else:
            log(f"WARNING: moved message {m.get('id')} to Snoozed but could not obtain its "
                "internetMessageId; it will NOT auto-unsnooze.")

    # Update snoozer.json
    if os.path.exists(OUTLOOK_SNOOZER_FILE):
        with open(OUTLOOK_SNOOZER_FILE, "r") as f:
            mySnoozes = json.load(f)
    else:
        mySnoozes = {}

    for mid in message_ids:
        mySnoozes[mid] = snooze_date

    _write_snoozes(mySnoozes)

    try:
        os.remove(SNOOZE_CACHE_FILE)
    except OSError:
        pass

    log(f"Snoozer file updated ({len(message_ids)} messages via Graph API)")
    return snooze_date


def snooze_via_applescript(snooze_date):
    """Snooze the selected email(s) via AppleScript (legacy Outlook)."""
    scpt = '''
        tell application "Microsoft Outlook"

            activate

            if view of the first main window is not equal to "mail view" then
                set view of the main window 1 to mail view

            end if

            set msgSet to selection

            set myIDs to {}

            repeat with aMessage in msgSet

                set msgID to id of aMessage
                set end of myIDs to msgID
                move aMessage to folder "Snoozed"



            end repeat

        end tell

        set AppleScript's text item delimiters to ":::"
        set myIDs to myIDs as text
        set AppleScript's text item delimiters to ""
        return myIDs

    end run '''

    command = ['osascript', '-e', scpt]
    myOutput = check_output(command).decode('utf-8').strip()

    myIDs = myOutput.split(":::")

    if os.path.exists(OUTLOOK_SNOOZER_FILE):
        with open(OUTLOOK_SNOOZER_FILE, "r") as f:
            mySnoozes = json.load(f)
    else:
        mySnoozes = {}

    for myID in myIDs:
        myRecordID = fetchRecordID(myID)
        mySnoozes[myRecordID] = snooze_date

    _write_snoozes(mySnoozes)

    log("Snoozer file updated")
    return snooze_date


def main():

    MY_SNOOZE_DATE = sys.argv[1]

    if is_new_outlook():
        result = snooze_via_graph(MY_SNOOZE_DATE)
    else:
        result = snooze_via_applescript(MY_SNOOZE_DATE)

    if result:
        print(result)


if __name__ == '__main__':
    main()


