"""
Save the selected email as .eml and archive it.
Supports both legacy (AppleScript) and new (Graph API) Outlook.

Usage: python3 saveEmail.py "/path/to/destination/folder"
Output: prints the final file path to stdout (for clipboard)

For new Outlook: captureSelected.py must run BEFORE the folder picker
appears (at keyword trigger time) to cache the selected email info.
"""

import sys
import os
import re
import json
from subprocess import check_output, PIPE
from consts import *
from util import log

CACHE_FILE = os.path.join(os.getenv('alfred_workflow_data', '/tmp'), 'selected_email_cache.json')

# Archive (move the message to Archive) only when explicitly requested. The `ols`
# keyword/hotkey saves only; the `olsa` keyword sets SAVE_ARCHIVE=1 to also archive.
SHOULD_ARCHIVE = os.getenv('SAVE_ARCHIVE') == '1'


def is_new_outlook():
    try:
        check_output(
            ['osascript', '-e', 'tell application "Microsoft Outlook" to get first exchange account'],
            stderr=PIPE
        ).decode('utf-8').strip()
        return False
    except Exception:
        return True


def save_via_graph(dest_folder):
    """Save selected email as .eml via Graph API, archive it, return final path.
    Reads the message ID from the cache file written by captureSelected.py."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_path = os.path.join(script_dir, "lib")
    for p in [lib_path, script_dir]:
        if p not in sys.path:
            sys.path.insert(0, p)

    from auth import get_token, GRAPH_ENDPOINT
    import requests

    # Read cached email info (written by captureSelected.py at keyword trigger time)
    if not os.path.exists(CACHE_FILE):
        log("No cached email info found. Make sure captureSelected.py runs before the folder picker.")
        return None

    with open(CACHE_FILE, "r") as f:
        cache = json.load(f)

    msg_id = cache.get("id")
    if not msg_id:
        log("Cached email info missing message ID.")
        return None

    token = get_token()

    # Download raw MIME content
    resp = requests.get(
        f"{GRAPH_ENDPOINT}/me/messages/{msg_id}/$value",
        headers={"Authorization": f"Bearer {token}"},
    )
    if resp.status_code != 200:
        log(f"Error downloading MIME (HTTP {resp.status_code})")
        return None

    # Build filename: date_subject_id.eml (matching legacy format)
    received = cache.get("receivedDateTime", "")
    date_prefix = received[:10] if received else "unknown"
    safe_subject = re.sub(r'[<>?|:/]', '', cache.get("subject", ""))
    short_id = msg_id[:20]
    filename = f"{date_prefix}_{safe_subject}_{short_id}.eml"

    # Save to destination folder
    os.makedirs(dest_folder, exist_ok=True)
    filepath = os.path.join(dest_folder, filename)
    with open(filepath, "wb") as f:
        f.write(resp.content)

    log(f"Saved: {filepath}")

    # Archive the message (only when requested via SAVE_ARCHIVE=1)
    if SHOULD_ARCHIVE:
        archive_id = None
        folder_resp = requests.get(
            f"{GRAPH_ENDPOINT}/me/mailFolders",
            headers={"Authorization": f"Bearer {token}"},
            params={"$filter": "displayName eq 'Archive'", "$select": "id"},
        )
        if folder_resp.status_code == 200:
            folders = folder_resp.json().get("value", [])
            if folders:
                archive_id = folders[0]["id"]

        if archive_id:
            move_resp = requests.post(
                f"{GRAPH_ENDPOINT}/me/messages/{msg_id}/move",
                headers={"Authorization": f"Bearer {token}"},
                json={"destinationId": archive_id},
            )
            if move_resp.status_code == 201:
                log("Archived message.")
            else:
                log(f"Archive failed (HTTP {move_resp.status_code})")
        else:
            log("Archive folder not found, skipping archive.")

    # Clean up cache file
    try:
        os.remove(CACHE_FILE)
    except OSError:
        pass

    return filepath


def save_via_applescript(dest_folder):
    """Save selected email as .eml via AppleScript, optionally archive, return final path."""
    staging_folder = "Containers/com.microsoft.Outlook/Data/Library/Application Support/com.microsoft.Outlook/SavingFolder"
    library_path = os.path.expanduser("~/Library")
    staging_path = os.path.join(library_path, staging_folder)

    archive_block = '''
        set messageAccount to first exchange account
        try
            repeat with aMessage in messages_
                move aMessage to folder named "Archive" of messageAccount
            end repeat
        end try
''' if SHOULD_ARCHIVE else ""

    scpt = f'''
on run argv
    set folderPath to "{staging_path}"

    tell application "System Events"
        if not (exists folder folderPath) then
            do shell script "mkdir -p " & quoted form of folderPath
        end if
    end tell

    set theFileName to POSIX file folderPath

    tell application "Microsoft Outlook"
        set messages_ to the selection

        repeat with i from 1 to number of items in messages_
            set theMsg to item i of messages_

            try
                set myID to id of theMsg
                set dateObj to time received of theMsg
            end try

            set theMonth to text -1 thru -2 of ("0" & (month of dateObj as number))
            set theDay to text -1 thru -2 of ("0" & day of dateObj)
            set theYear to year of dateObj
            set dateStamp to (((theYear as string) & "-" & theMonth as string) & "-" & theDay as string)

            set mySubject to subject of theMsg
            set myEditSubject to do shell script "echo " & quoted form of mySubject & " | sed 's|[<>?|:/]||g'"
            set fullName to dateStamp & "_" & myEditSubject & "_" & myID & ".eml"
            set textPath to (folderPath & "/" & fullName) as string
            save theMsg in textPath

        end repeat
{archive_block}
    end tell

    set sourceFolder to theFileName as alias
    set destinationFolder to POSIX file (first item of argv as text)
    tell application "Finder"
        move entire contents of sourceFolder to destinationFolder
    end tell
    set finalPath to first item of argv as text & "/" & fullName
    return finalPath
end run
'''

    command = ['osascript', '-e', scpt, dest_folder]
    output = check_output(command).decode('utf-8').strip()
    return output


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 saveEmail.py /path/to/destination", file=sys.stderr)
        sys.exit(1)

    dest_folder = sys.argv[1]

    if is_new_outlook():
        result = save_via_graph(dest_folder)
    else:
        result = save_via_applescript(dest_folder)

    if result:
        print(result)


if __name__ == '__main__':
    main()
