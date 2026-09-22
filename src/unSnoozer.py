"""
A new version of the OL SNOOZER script, to be included in the alfred-OutlookSuite
Sunny ☀️   🌡️+70°F (feels +70°F, 59%) 🌬️↓2mph 🌕&m Fri Jun  2 08:42:48 2023
W22Q2 – 153 ➡️ 211 – 22 ❇️ 343
"""

# Snooze by date applescript, Monday, May 25, 2020, 1:17 PM
# modified on Saturday, May 30, 2020, 1:25 PM to overwrite the ID and date files and avoid duplications
## Thursday, November 19 2020, 9:45AM added count of unsnoozing tomorrow, which is helpful to plan next day
## Friday, Decemeber 11 2020, added multiple days

# on Mon Apr 11 10:17:50 2022  Sunny ☀️   🌡️+48°F (feels +45°F, 39%) 🌬️↙9mph 🌔
# tried to use exchange id instead of database ID, so that it will be possible to use it across computers, however the exchange ID changes if the message changes folder.

# Partly cloudy ⛅️  🌡️+69°F (feels +69°F, 41%) 🌬️↘13mph 🌔 Tue Apr 12 17:03:32 2022
# tried to fetch the message ID from the sqlite database. To my knowledge is the only way to obtain a message ID which is shared across computers. The unique id I have been using is not.
#### I copied the applescript wihtin alfred as opposed to run it from the bash because I don't have sudo permissions to make the applescript executable.

# Sunny ☀️   🌡️+87°F (feels +86°F, 35%) 🌬️←7mph 🌕&m Fri Jun  2 13:04:58 2023
#W22Q2 – 153 ➡️ 211 – 22 ❇️ 343
# streamlined version to be included in the alfred-outlookSuite workflow


from subprocess import Popen, PIPE, check_output
from util import log, fetchEmailID
from consts import *
import json
from datetime import date


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


def _prune_snoozes(ids_to_remove):
    """Remove the given keys from the snooze store, re-reading it first so a
    snooze added since we loaded it isn't clobbered."""
    if not ids_to_remove:
        return
    if os.path.exists(OUTLOOK_SNOOZER_FILE):
        with open(OUTLOOK_SNOOZER_FILE, "r") as f:
            data = json.load(f)
    else:
        data = {}
    for k in ids_to_remove:
        data.pop(k, None)
    _write_snoozes(data)


def unsnooze_via_graph(myuniIDs, search_date):
    """Move snoozed messages back to Inbox via Graph API.

    Returns (processed_ids, moved_count, inbox_count, snoozed_count) where
    processed_ids are the entries safe to remove from the store: those moved
    back, plus those no longer in Snoozed (already gone). Entries that hit a
    transient error are NOT included, so they're retried next run rather than
    silently dropped."""
    import sys
    # Add bundled lib/ for msal/requests, and src/ for auth
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_path = os.path.join(script_dir, "lib")
    for p in [lib_path, script_dir]:
        if p not in sys.path:
            sys.path.insert(0, p)
    from auth import get_token, GRAPH_ENDPOINT, MissingConfigError
    import requests

    try:
        token = get_token()
    except MissingConfigError:
        log("Unsnooze skipped: Microsoft sign-in not configured.")
        return [], 0, 0, 0

    # Find Snoozed folder ID
    resp = requests.get(
        f"{GRAPH_ENDPOINT}/me/mailFolders",
        headers={"Authorization": f"Bearer {token}"},
        params={"$filter": "displayName eq 'Snoozed'", "$select": "id"},
    )
    folders = resp.json().get("value", []) if resp.status_code == 200 else []
    if not folders:
        log("Snoozed folder not found via Graph API.")
        return [], 0, 0, 0

    snoozed_folder_id = folders[0]["id"]

    moved = 0
    processed_ids = []
    for inet_id in myuniIDs:
        # Find the message in the Snoozed folder by internetMessageId
        safe_id = inet_id.replace("'", "''")
        resp = requests.get(
            f"{GRAPH_ENDPOINT}/me/mailFolders/{snoozed_folder_id}/messages",
            headers={"Authorization": f"Bearer {token}"},
            params={
                "$filter": f"internetMessageId eq '{safe_id}'",
                "$select": "id,subject",
                "$top": 1,
            },
        )
        if resp.status_code != 200:
            log(f"Error searching for {inet_id}: {resp.status_code}")
            continue
        msgs = resp.json().get("value", [])
        if not msgs:
            # No longer in Snoozed (already moved/deleted): stop tracking it.
            log(f"Message not found in Snoozed (dropping from store): {inet_id}")
            processed_ids.append(inet_id)
            continue

        # Move to Inbox
        move_resp = requests.post(
            f"{GRAPH_ENDPOINT}/me/messages/{msgs[0]['id']}/move",
            headers={"Authorization": f"Bearer {token}"},
            json={"destinationId": "inbox"},
        )
        if move_resp.status_code == 201:
            moved += 1
            processed_ids.append(inet_id)
        else:
            log(f"Failed to move {inet_id}: {move_resp.status_code}")

    # Get current counts
    inbox_resp = requests.get(
        f"{GRAPH_ENDPOINT}/me/mailFolders/inbox",
        headers={"Authorization": f"Bearer {token}"},
        params={"$select": "totalItemCount"},
    )
    inbox_count = inbox_resp.json().get("totalItemCount", 0) if inbox_resp.status_code == 200 else 0

    snoozed_resp = requests.get(
        f"{GRAPH_ENDPOINT}/me/mailFolders/{snoozed_folder_id}",
        headers={"Authorization": f"Bearer {token}"},
        params={"$select": "totalItemCount"},
    )
    snoozed_count = snoozed_resp.json().get("totalItemCount", 0) if snoozed_resp.status_code == 200 else 0

    return processed_ids, moved, inbox_count, snoozed_count


def unsnooze_via_applescript(myuniIDs):
    """Move snoozed messages back to Inbox via AppleScript (legacy Outlook)."""
    nToUnsnooze = str(len(myuniIDs))

    myIDs = []
    for myuniID in myuniIDs:
        myEmailID = fetchEmailID(myuniID)
        myIDs.append(myEmailID)

    scpt = '''

    on run {myIDs, OUTLOOK_LOG_FILE,nToUnsnooze, BEEUSER, BEETOKEN, BEEGOAL, BEEMINDER}
        set AppleScript's text item delimiters to ","
        set IDlist to every text item of myIDs
        set AppleScript's text item delimiters to {""} -- Reset delimiters

        tell application "Microsoft Outlook"
            set mailAccount to first exchange account
            set myMailbox to folder "Snoozed" of mailAccount
            set theMsgs to (every message of myMailbox)


            repeat with aMsg in theMsgs -- going through all the messages in the Snoozed folder
                set msgID to (id of aMsg) as string

                repeat with i in IDlist
                    if msgID is (i as string) then
                        #log "found one"
                        move aMsg to folder "Inbox" of mailAccount
                    end if
                end repeat
            end repeat
        set InboxCount to (count messages in folder "Inbox" of mailAccount)
	    set SnoozeCount to (count messages in folder "Snoozed" of mailAccount)

        #logging
        set LogText to ("ToInbox=" & nToUnsnooze & ", Snoozed=" & SnoozeCount & ", TotalInbox=" & InboxCount)

        #set myLogPath to OUTLOOK_LOG_FILE
        do shell script "echo " & LogText & ", $(date) >> " & quoted form of OUTLOOK_LOG_FILE

        end tell



        ####BEEMINDER section
	    #beeminder POST syntax:
	    #data = [
	    #  ('auth_token', 'YOUR-BEEMINDER-AUTH-TOKEN-GOES-HERE'),
	    #  ('value', currentNoteCount),
	    #  ('comment', 'from EvernoteScript!'),
	    #]

        if BEEMINDER is equal to "1" then

	        set theURL to "https://www.beeminder.com/api/v1/users/" & BEEUSER & "/goals/" & BEEGOAL & "/datapoints.json" & " -d " & "auth_token=" & BEETOKEN & " -d " & "value=" & InboxCount & " -d " & "comment=from+snoozeScript"
	        #	log quoted form of the theURL


                # posting to beeminder board
                do shell script "curl -X POST " & theURL
        end if
        return InboxCount
    end run
            '''

    IDString = ",".join(myIDs)
    command = ['osascript', '-e', scpt, IDString, OUTLOOK_LOG_FILE, nToUnsnooze, BEEUSER, BEETOKEN, BEEGOAL, BEEMINDER]
    currentInboxCount = check_output(command).decode('utf-8').strip()
    return nToUnsnooze, currentInboxCount


def post_to_beeminder(inbox_count):
    """Post the current inbox count to Beeminder, if configured (BEEMINDER on
    plus BEEUSER/BEETOKEN/BEEGOAL). Mirrors the legacy AppleScript behavior for
    the Graph API path. Best-effort: failures are logged, never fatal."""
    if BEEMINDER != "1" or not (BEEUSER and BEETOKEN and BEEGOAL):
        return
    import sys
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for p in [os.path.join(script_dir, "lib"), script_dir]:
        if p not in sys.path:
            sys.path.insert(0, p)
    import requests
    try:
        resp = requests.post(
            f"https://www.beeminder.com/api/v1/users/{BEEUSER}/goals/{BEEGOAL}/datapoints.json",
            data={
                "auth_token": BEETOKEN,
                "value": inbox_count,
                "comment": "from alfred-outlookSuite unsnooze",
            },
            timeout=10,
        )
        if resp.status_code in (200, 201):
            log(f"Posted inbox count {inbox_count} to Beeminder goal '{BEEGOAL}'.")
        else:
            log(f"Beeminder post failed (HTTP {resp.status_code}): {resp.text[:150]}")
    except Exception as e:
        log(f"Beeminder post error: {e}")


def main():
    # Get the current date
    today = date.today()

    # Convert today's date to the same format as the dictionary values
    search_date = today.strftime("%Y-%m-%d")

    #reading in the JSON file with the snooze information
    if os.path.exists(OUTLOOK_SNOOZER_FILE):
        with open(OUTLOOK_SNOOZER_FILE, "r") as f:
            mySnoozes = json.load(f)
    else:
        mySnoozes = {}

    # Entries due today or earlier. We do NOT prune the store yet: pruning only
    # happens for entries actually handled, AFTER the move, so a failed move
    # (network/auth/API error) leaves them tracked for the next run instead of
    # stranding the emails in Snoozed forever.
    myuniIDs = [key for key, value in mySnoozes.items() if value <= search_date]

    nToUnsnooze = str(len(myuniIDs))

    if not myuniIDs:
        result = {"items": [{
            "title": "No emails to unsnooze today",
            "subtitle": "ready to use outlookSuite now",
            "arg": "0",
            "icon": {"path": "icons/done.png"}
        }]}
        print(json.dumps(result))
        return

    if is_new_outlook():
        processed_ids, moved, inbox_count, snoozed_count = unsnooze_via_graph(myuniIDs, search_date)
        currentInboxCount = str(inbox_count)

        # Prune only the entries we actually handled, now that the move is done.
        _prune_snoozes(processed_ids)

        # Post the current inbox count to Beeminder (if configured)
        post_to_beeminder(inbox_count)

        # Log
        log_text = f"ToInbox={moved}, Snoozed={snoozed_count}, TotalInbox={inbox_count}"
        try:
            with open(OUTLOOK_LOG_FILE, "a") as f:
                from datetime import datetime
                f.write(f"{log_text}, {datetime.now()}\n")
        except Exception:
            pass

        result = {"items": [{
            "title": f"Done! {moved} emails unsnoozed to Inbox ({currentInboxCount} total)",
            "subtitle": "ready to use outlookSuite now",
            "arg": currentInboxCount,
            "icon": {"path": "icons/done.png"}
        }]}
    else:
        # AppleScript moves everything it can in one call; prune the due entries
        # only after it returns without raising.
        nToUnsnooze, currentInboxCount = unsnooze_via_applescript(myuniIDs)
        _prune_snoozes(myuniIDs)
        result = {"items": [{
            "title": f"Done! {nToUnsnooze} emails unsnoozed to Inbox ({currentInboxCount} total)",
            "subtitle": "ready to use outlookSuite now",
            "arg": currentInboxCount,
            "icon": {"path": "icons/done.png"}
        }]}

    print(json.dumps(result))


if __name__ == '__main__':
    main()
