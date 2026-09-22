"""
OUTLOOK-SUITE
A suite of tools for Outlook
Light rain, mist 🌦   🌡️+60°F (feels +60°F, 93%) 🌬️↙6mph 🌑&m Sat May 20 11:20:32 2023
W20Q2 – 140 ➡️ 224 – 9 ❇️ 356

"""

import sys
import os
import sqlite3
from subprocess import check_output, PIPE
from consts import *
from util import log, timestampToText,  checkingTime, checkTimespan, checkJSON
import json
import subprocess
from datetime import datetime, timedelta


CACHE_DIR = os.getenv('alfred_workflow_cache', '/tmp')
CACHE_FILE = os.path.join(CACHE_DIR, 'graph_messages_cache.json')
CACHE_MINUTES = int(os.getenv('CACHE_MINUTES', '30'))
# Max messages to pull across paginated Graph responses. Graph caps a single
# page at 1000; we follow @odata.nextLink until we hit this ceiling so searches
# reach beyond the newest page. Covers the primary mailbox only (not the
# online archive, which Graph does not expose).
MAX_RESULTS = int(os.getenv('MAX_RESULTS', '200'))
PAGE_SIZE = 50

#initializing JSON output
result = {"items": [], "variables":{}}


def is_new_outlook():
    """Check if the new Outlook is running (no AppleScript exchange account support)."""
    try:
        check_output(
            ['osascript', '-e', 'tell application "Microsoft Outlook" to get first exchange account'],
            stderr=PIPE
        ).decode('utf-8').strip()
        return False
    except Exception:
        return True


def _init_graph():
    """Add lib/ and src/ to sys.path, return (token, GRAPH_ENDPOINT, requests)."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lib_path = os.path.join(script_dir, "lib")
    for p in [lib_path, script_dir]:
        if p not in sys.path:
            sys.path.insert(0, p)
    from auth import get_token, GRAPH_ENDPOINT
    import requests
    token = get_token()
    return token, GRAPH_ENDPOINT, requests


def _graph_get_all(url, headers, params, requests, max_results=MAX_RESULTS):
    """GET a messages endpoint, following @odata.nextLink up to max_results.
    Returns (messages, status_code, error_text). On the first non-200, stops
    and returns what it has plus the error."""
    messages = []
    first_status = None
    while url and len(messages) < max_results:
        resp = requests.get(url, headers=headers, params=params)
        if first_status is None:
            first_status = resp.status_code
        if resp.status_code != 200:
            return messages, resp.status_code, resp.text[:300]
        data = resp.json()
        messages.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None  # nextLink already encodes params
    return messages[:max_results], first_status, None


def _graph_get_folder_map(token, GRAPH_ENDPOINT, requests):
    """Build a {folder_id: folder_name} map of the top-level mail folders.

    Only top-level folders are mapped: walking the full nested tree is many
    sequential API calls (tens of seconds on a deep mailbox) and this runs on
    every search. Nested subfolders are therefore not resolvable via folder:
    (they report "folder not found"); this is a documented limitation.
    """
    folder_map = {}
    url = f"{GRAPH_ENDPOINT}/me/mailFolders"
    while url:
        resp = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                            params={"$select": "id,displayName", "$top": 100})
        if resp.status_code != 200:
            break
        data = resp.json()
        for f in data.get("value", []):
            folder_map[f["id"]] = f["displayName"]
        url = data.get("@odata.nextLink")
    return folder_map


def _graph_find_folder_id(token, GRAPH_ENDPOINT, requests, folder_name):
    """Find a folder ID by display name."""
    resp = requests.get(
        f"{GRAPH_ENDPOINT}/me/mailFolders",
        headers={"Authorization": f"Bearer {token}"},
        params={"$filter": f"displayName eq '{folder_name}'", "$select": "id"},
    )
    if resp.status_code == 200:
        folders = resp.json().get("value", [])
        if folders:
            return folders[0]["id"]
    return None


def _load_cache():
    """Load cached messages if the cache exists and hasn't expired."""
    if not os.path.exists(CACHE_FILE):
        return None
    try:
        with open(CACHE_FILE, "r") as f:
            cache = json.load(f)
        cached_at = datetime.fromisoformat(cache["timestamp"])
        if datetime.now() - cached_at > timedelta(minutes=CACHE_MINUTES):
            return None
        return cache
    except Exception:
        return None


def _save_cache(messages, folder_map):
    """Save messages and folder map to cache with a timestamp."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache = {
        "timestamp": datetime.now().isoformat(),
        "messages": messages,
        "folder_map": folder_map,
    }
    with open(CACHE_FILE, "w") as f:
        json.dump(cache, f)


def _filter_cached(messages, gq, folder_name_to_id):
    """Apply filter-only query conditions to cached messages client-side."""
    for filt in gq['filters']:
        if "hasAttachments eq true" in filt:
            messages = [m for m in messages if m.get("hasAttachments")]
        elif "isRead eq true" in filt:
            messages = [m for m in messages if m.get("isRead")]
        elif "isRead eq false" in filt:
            messages = [m for m in messages if not m.get("isRead")]
        elif "importance eq 'high'" in filt:
            messages = [m for m in messages if m.get("importance") == "high"]
        elif "importance eq 'low'" in filt:
            messages = [m for m in messages if m.get("importance") == "low"]
        elif "receivedDateTime ge" in filt:
            since_str = filt.split("ge ")[1]
            messages = [m for m in messages if m.get("receivedDateTime", "") >= since_str]

    if gq['folder_name']:
        fid = folder_name_to_id.get(gq['folder_name'].casefold())
        if fid:
            messages = [m for m in messages if m.get("parentFolderId") == fid]

    if gq['exclude_terms']:
        for term in gq['exclude_terms']:
            term_lower = term.lower()
            messages = [m for m in messages
                        if term_lower not in m.get("subject", "").lower()
                        and term_lower not in m.get("bodyPreview", "").lower()]

    reverse = gq['sort_order'] != 'ASC'
    messages.sort(key=lambda m: m.get("receivedDateTime", ""), reverse=reverse)
    return messages


def compileGraphQuery(myQuery, myContacts):
    """Parse user query into Graph API parameters.
    Returns a dict with keys: search_terms, filters, folder_id, folder_name,
    sort_order, exclude_terms, is_thread, thread_topic.
    """
    MYITER = os.getenv('myIter')
    MYSOURCE = os.getenv('mySource')

    search_parts = []    # $search KQL terms (all) — used only to detect "is this a search?"
    filters = []         # $filter OData terms
    exclude_terms = []   # client-side exclusion
    match_criteria = []  # client-side strict match: (type, value) for from/to/cc/subject
    text_parts = []      # free-text terms — MUST go to $search (body isn't readable client-side)
    narrow_parts = []    # from/subject $search terms (selective; also client-filtered)
    broad_parts = []     # to/cc $search terms (match broadly; also client-filtered)
    sort_order = 'DESC'
    folder_name = None
    is_thread = False
    thread_topic = None

    # Graph $search is relevance-based, NOT a strict filter, and ANDing terms
    # dilutes it (e.g. from:X AND sometext can push all of X's mail out of the
    # fetched window). Strategy: send the term that can ONLY be matched via the
    # API (free text, against the body) to $search; fetch a candidate set with a
    # single selective term otherwise; then enforce every from/to/cc/subject
    # term precisely client-side via match_criteria.
    def add_from(val):
        search_parts.append(f'"from:{val}"'); narrow_parts.append(f'"from:{val}"')
        match_criteria.append(("from", val))
    def add_to(val):
        search_parts.append(f'"to:{val}"'); broad_parts.append(f'"to:{val}"')
        match_criteria.append(("to", val))
    def add_cc(val):
        search_parts.append(f'"cc:{val}"'); broad_parts.append(f'"cc:{val}"')
        match_criteria.append(("cc", val))
    def add_subject(val):
        search_parts.append(f'"subject:{val}"'); narrow_parts.append(f'"subject:{val}"')
        match_criteria.append(("subject", val))
    def add_text(val):
        # Free text is matched by $search against the full body, so it is NOT
        # client-filtered (bodyPreview is only the first ~250 chars).
        search_parts.append(f'"{val}"'); text_parts.append(f'"{val}"')

    myElements = myQuery.split()
    for myElement in myElements:
        if myElement.startswith("from:"):
            val_raw = myElement.split(":")[1]
            if myElement == 'from:me':
                add_from(MYSELF)
            elif "_" in val_raw:
                # An underscore means the user gave the full name (e.g.
                # from:john_appleseed) — search directly, don't autocomplete.
                add_from(val_raw.strip().replace("_", " "))
            elif MYITER is None or (MYITER == '1' and MYSOURCE != "contacts"):
                if CONTACT_AUTOCOMPLETE != "None":
                    showContacts(val_raw, myContacts, myQuery, "from")
                else:
                    add_from(val_raw.strip())
            else:
                add_from(val_raw.strip())

        elif myElement.startswith("to:"):
            val_raw = myElement.split(":")[1]
            if myElement == 'to:me':
                add_to(MYSELF)
            elif "_" in val_raw:
                add_to(val_raw.strip().replace("_", " "))
            elif MYITER is None:
                if CONTACT_AUTOCOMPLETE != "None":
                    showContacts(val_raw, myContacts, myQuery, "to")
                else:
                    add_to(val_raw.strip())
            else:
                add_to(val_raw.strip())

        elif myElement.startswith("sq:"):
            if MYITER is None:
                showSavedQueries(myElement.split(":")[1], SAVED_QUERIES, myQuery)

        elif myElement.startswith("cc:"):
            if myElement == 'cc:me':
                add_cc(MYSELF)
            else:
                add_cc(myElement.split(":")[1].strip().replace("_", " "))

        elif myElement == "has:attach":
            filters.append("hasAttachments eq true")

        elif myElement.startswith("subject:"):
            add_subject(myElement.split(":")[1].strip())

        elif myElement.startswith("since:"):
            val = myElement.split(":")[1].strip()
            # Parse days/weeks/months
            import re
            match = re.match(r'^(\d+)([wm]?)$', val)
            if match:
                num, unit = int(match.group(1)), match.group(2)
                if unit == 'w':
                    num *= 7
                elif unit == 'm':
                    num *= 30
            else:
                num = int(val)
            since_dt = datetime.now() - timedelta(days=num)
            filters.append(f"receivedDateTime ge {since_dt.strftime('%Y-%m-%dT00:00:00Z')}")

        elif myElement.startswith("is:"):
            val = myElement.split(":")[1].strip()
            if val == 'read':
                filters.append("isRead eq true")
            elif val == 'unread':
                filters.append("isRead eq false")
            elif val == 'important':
                filters.append("importance eq 'high'")
            elif val == 'unimportant':
                filters.append("importance eq 'low'")

        elif myElement.startswith("folder:"):
            folder_name = myElement.split(":")[1].strip().replace("_", " ")

        elif myElement.startswith("account:"):
            pass  # account filtering not available in Graph API single-user context

        elif myElement == "--a":
            sort_order = 'ASC'

        elif myElement.startswith("-"):
            exclude_terms.append(myElement[1:].strip())

        else:
            # default: free text search in subject and body
            add_text(myElement)

    # Choose what to send to $search:
    #   - Free text MUST be searched (can't verify body client-side) → send all of it.
    #   - Otherwise fetch a candidate set with a SINGLE selective term (avoids
    #     AND-dilution); from/subject preferred, else to/cc. All from/to/cc/subject
    #     terms are still enforced precisely client-side via match_criteria.
    if text_parts:
        api_search_parts = text_parts
    elif narrow_parts:
        api_search_parts = narrow_parts[:1]
    else:
        api_search_parts = broad_parts[:1]

    return {
        'search_terms': search_parts,
        'api_search_parts': api_search_parts,
        'match_criteria': match_criteria,
        'filters': filters,
        'folder_name': folder_name,
        'sort_order': sort_order,
        'exclude_terms': exclude_terms,
    }


def _recipient_haystack(recipients):
    parts = []
    for r in recipients or []:
        ea = r.get("emailAddress", {})
        parts.append(f"{ea.get('name', '')} {ea.get('address', '')}")
    return " ".join(parts).lower()


def _msg_matches_criteria(m, match_criteria):
    """Strictly verify a message against from/to/cc/subject criteria client-side,
    since Graph $search only ranks by relevance and does not filter precisely."""
    for mtype, val in match_criteria:
        v = val.lower()
        if mtype == "from":
            ea = m.get("from", {}).get("emailAddress", {})
            if v not in f"{ea.get('name', '')} {ea.get('address', '')}".lower():
                return False
        elif mtype == "to":
            if v not in _recipient_haystack(m.get("toRecipients")):
                return False
        elif mtype == "cc":
            if v not in _recipient_haystack(m.get("ccRecipients")):
                return False
        elif mtype == "subject":
            if v not in m.get("subject", "").lower():
                return False
    return True


def handleGraph(myQuery, myContacts, is_thread=False, thread_topic=None):
    """Execute a mail search via Graph API and output Alfred JSON."""
    token, GRAPH_ENDPOINT, requests = _init_graph()

    folder_map = _graph_get_folder_map(token, GRAPH_ENDPOINT, requests)
    # Build reverse map: name -> id (case-insensitive)
    folder_name_to_id = {v.casefold(): k for k, v in folder_map.items()}

    # Thread view: search by conversationId
    if is_thread and thread_topic:
        # First find a message with this subject to get its conversationId
        resp = requests.get(
            f"{GRAPH_ENDPOINT}/me/messages",
            headers={"Authorization": f"Bearer {token}"},
            params={
                "$search": f'"subject:{thread_topic}"',
                "$select": "conversationId",
                "$top": 1,
            },
        )
        conv_id = None
        if resp.status_code == 200:
            msgs = resp.json().get("value", [])
            if msgs:
                conv_id = msgs[0].get("conversationId")
        if conv_id:
            safe_conv = conv_id.replace("'", "''")
            resp = requests.get(
                f"{GRAPH_ENDPOINT}/me/messages",
                headers={"Authorization": f"Bearer {token}"},
                params={
                    "$filter": f"conversationId eq '{safe_conv}'",
                    "$select": "id,subject,from,toRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview,conversationId,parentFolderId,webLink",
                    "$top": 50,
                },
            )
            messages = resp.json().get("value", []) if resp.status_code == 200 else []
            # Sort ascending for thread view
            messages.sort(key=lambda m: m.get("receivedDateTime", ""))
        else:
            messages = []
        _format_graph_results(messages, folder_map)
        return

    # Normal search
    gq = compileGraphQuery(myQuery, myContacts)

    select_fields = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,isRead,importance,bodyPreview,conversationId,parentFolderId,webLink"

    # --- Cache logic ---
    # Only the bare "no query" recent listing is served from the local cache.
    # Filter/folder queries (is:, since:, has:, folder:, -exclude) must hit the
    # API so they search server-side across the whole mailbox/folder rather than
    # just the recent cached window (which would miss older/archived mail and
    # can't reach folder-scoped results).
    is_bare_listing = not (gq['search_terms'] or gq['filters']
                           or gq['folder_name'] or gq['exclude_terms'])
    if is_bare_listing:
        cache = _load_cache()
        if cache:
            log("Using cached messages")
            cached_folder_map = cache.get("folder_map", {})
            cached_name_to_id = {v.casefold(): k for k, v in cached_folder_map.items()}
            messages = cache["messages"]

            excl_ids = set()
            for ef in EXCL_FOLDERS:
                fid = cached_name_to_id.get(ef.strip().casefold())
                if fid:
                    excl_ids.add(fid)
            if excl_ids:
                messages = [m for m in messages if m.get("parentFolderId") not in excl_ids]

            reverse = gq['sort_order'] != 'ASC'
            messages.sort(key=lambda m: m.get("receivedDateTime", ""), reverse=reverse)
            _format_graph_results(messages, cached_folder_map)
            return

    # --- API path ---
    # Determine endpoint (specific folder or all messages)
    if gq['folder_name']:
        fid = folder_name_to_id.get(gq['folder_name'].casefold())
        if fid:
            base_url = f"{GRAPH_ENDPOINT}/me/mailFolders/{fid}/messages"
        else:
            # Folder genuinely not found — say so rather than silently
            # searching the whole mailbox (which looks like the filter is ignored).
            result["items"].append({
                "title": f"Folder not found: {gq['folder_name']}",
                "subtitle": "Check the name — use underscores for spaces (e.g. folder:sent_items)",
                "arg": "", "icon": {"path": "icons/Warning.png"},
            })
            result['variables'] = {"mySource": "mailList"}
            print(json.dumps(result))
            return
    else:
        base_url = f"{GRAPH_ENDPOINT}/me/messages"

    # Graph API doesn't support $search + $filter together well.
    # Strategy: use $search if we have search terms, $filter if we have filters only,
    # or $search with client-side post-filtering for filters.
    params = {"$select": select_fields, "$top": PAGE_SIZE}

    if gq['search_terms'] and not gq['filters']:
        # Pure search
        params["$search"] = " AND ".join(gq['api_search_parts'])
    elif gq['filters'] and not gq['search_terms']:
        # Pure filter
        params["$filter"] = " and ".join(gq['filters'])
        # Some filters can't be combined with $orderby receivedDateTime server-side
        # (Graph returns 400 "InefficientFilter"): hasAttachments and importance.
        # For those, skip $orderby and let the client-side sort handle ordering.
        unsortable = any(("hasAttachments" in f) or ("importance" in f)
                         for f in gq['filters'])
        if not unsortable:
            order_dir = "asc" if gq['sort_order'] == 'ASC' else "desc"
            params["$orderby"] = f"receivedDateTime {order_dir}"
    elif gq['search_terms'] and gq['filters']:
        # Search + filter: use $search, post-filter on client side
        params["$search"] = " AND ".join(gq['api_search_parts'])
    else:
        # No query at all — just list recent messages
        order_dir = "asc" if gq['sort_order'] == 'ASC' else "desc"
        params["$orderby"] = f"receivedDateTime {order_dir}"

    log(f"Graph query: {base_url} params={params}")

    messages, status, err = _graph_get_all(
        base_url, {"Authorization": f"Bearer {token}"}, params, requests
    )
    if status != 200:
        log(f"Graph API error {status}: {err}")
        result["items"].append({
            "title": "Graph API error",
            "subtitle": f"HTTP {status}",
            "arg": "", "icon": {"path": "icons/Warning.png"}
        })
        result['variables'] = {"mySource": "mailList"}
        print(json.dumps(result))
        return

    # Save baseline to cache (no-query listing from all messages)
    if not gq['search_terms'] and not gq['filters'] and not gq['folder_name']:
        _save_cache(messages, folder_map)
        log(f"Cached {len(messages)} messages (TTL={CACHE_MINUTES}min)")

    # Client-side post-filtering for $filter conditions when $search was used
    if gq['search_terms'] and gq['filters']:
        for filt in gq['filters']:
            if "hasAttachments eq true" in filt:
                messages = [m for m in messages if m.get("hasAttachments")]
            elif "isRead eq true" in filt:
                messages = [m for m in messages if m.get("isRead")]
            elif "isRead eq false" in filt:
                messages = [m for m in messages if not m.get("isRead")]
            elif "importance eq 'high'" in filt:
                messages = [m for m in messages if m.get("importance") == "high"]
            elif "importance eq 'low'" in filt:
                messages = [m for m in messages if m.get("importance") == "low"]
            elif "receivedDateTime ge" in filt:
                since_str = filt.split("ge ")[1]
                messages = [m for m in messages if m.get("receivedDateTime", "") >= since_str]

    # Strict client-side match for from/to/cc/subject: Graph $search only ranks by
    # relevance, so a fetched result set may include non-matching senders/recipients
    # (or a broad term like to:me may have crowded out the real matches). Enforce
    # every such term precisely against the actual message fields.
    if gq['match_criteria']:
        messages = [m for m in messages if _msg_matches_criteria(m, gq['match_criteria'])]

    # Client-side exclusion
    if gq['exclude_terms']:
        for term in gq['exclude_terms']:
            term_lower = term.lower()
            messages = [m for m in messages
                        if term_lower not in m.get("subject", "").lower()
                        and term_lower not in m.get("bodyPreview", "").lower()]

    # Exclude configured folders (e.g. Deleted Items) from results — including
    # search/filter queries, so e.g. is:unread doesn't surface trash. Skip the
    # exclusion for a folder the user explicitly asked for (folder:Deleted_Items).
    excl_ids = set()
    for ef in EXCL_FOLDERS:
        fid = folder_name_to_id.get(ef.strip().casefold())
        if fid:
            excl_ids.add(fid)
    if gq['folder_name']:
        excl_ids.discard(folder_name_to_id.get(gq['folder_name'].casefold()))
    if excl_ids:
        messages = [m for m in messages if m.get("parentFolderId") not in excl_ids]

    # Sort client-side if $orderby wasn't used ($search returns by relevance,
    # and some $filter combinations don't support $orderby)
    if gq['search_terms'] or "$orderby" not in params:
        reverse = gq['sort_order'] != 'ASC'
        messages.sort(key=lambda m: m.get("receivedDateTime", ""), reverse=reverse)

    empty_title, empty_subtitle = _no_match_message(gq)
    _format_graph_results(messages, folder_map, empty_title, empty_subtitle)


def _no_match_message(gq):
    """Craft a targeted 'no matches' message reflecting the query (folder,
    sender/recipient/subject terms, and filters)."""
    parts = []
    for mtype, val in gq.get('match_criteria', []):
        parts.append(f"{mtype}:{val}")
    filt_labels = {
        "hasAttachments eq true": "has:attach",
        "isRead eq true": "is:read",
        "isRead eq false": "is:unread",
        "importance eq 'high'": "is:important",
        "importance eq 'low'": "is:unimportant",
    }
    for f in gq.get('filters', []):
        parts.append(filt_labels.get(f, "since:…" if f.startswith("receivedDateTime") else f))
    terms = " ".join(parts)
    folder = gq.get('folder_name')
    if folder:
        title = f"No matches in folder: {folder}"
        subtitle = f"for {terms}" if terms else "This folder has no matching messages"
    elif terms:
        title, subtitle = "No matches found", f"for {terms}"
    else:
        title, subtitle = "No matches found", "Try a different query"
    return title, subtitle


def _format_graph_results(messages, folder_map,
                          empty_title="No matches found",
                          empty_subtitle="Try a different query"):
    """Format Graph API messages into Alfred JSON output."""
    totCount = len(messages)
    myCounter = 0

    for m in messages:
        myCounter += 1
        subject = m.get("subject", "(no subject)")
        MySubjectClean = cleanSubject(subject)
        sender = m.get("from", {}).get("emailAddress", {})
        sender_str = sender.get("name", sender.get("address", ""))
        received = m.get("receivedDateTime", "")

        # Parse datetime for display
        try:
            dt = datetime.fromisoformat(received.replace("Z", "+00:00"))
            dt_local = dt.astimezone()
            timeRec = dt_local.strftime("%Y-%m-%d %H:%M")
            timeRecF = dt_local.strftime("%A, %B %d, %Y %I:%M %p")
        except (ValueError, TypeError):
            timeRec = received[:16] if received else ""
            timeRecF = received if received else ""

        # Icons
        readIcon = '' if m.get("isRead") else '⭐'
        importance = m.get("importance", "normal")
        urgentIcon = '‼️' if importance == "high" else ('🔽' if importance == "low" else '')
        attIcon = '📎' if m.get("hasAttachments") else ''

        # Folder name
        folder_id = m.get("parentFolderId", "")
        folder_name = folder_map.get(folder_id, "")
        SnoozeString = "💤" if folder_name == "Snoozed" else ""

        # Recipients for preview
        to_list = ", ".join([r.get("emailAddress", {}).get("name", r.get("emailAddress", {}).get("address", ""))
                            for r in m.get("toRecipients", [])])

        myPreview = f"{timeRecF}\nFrom: {sender_str}\nTo: {to_list}\n\nSubject: {MySubjectClean}\n\n{m.get('bodyPreview', '')}"

        # arg = webLink (opens message in Outlook/browser)
        webLink = m.get("webLink", "")
        conv_topic = subject

        result["items"].append({
            "title": f"{readIcon}{MySubjectClean}{urgentIcon}",
            'subtitle': f"{myCounter}/{totCount:,} {attIcon} [{folder_name}] From: {sender_str} {timeRec} {SnoozeString}",
            'valid': True,
            "quicklookurl": '',
            # Reset myIter so opening a message routes to the "open" branch,
            # not the "recur" (re-search) branch used for autocomplete acceptance.
            'variables': {"myIter": ""},
            "mods": {
                "control": {
                    "valid": 'true',
                    "subtitle": f"🧵 filter entire thread",
                    "arg": conv_topic,
                    'variables': {
                        "mySource": 'thread',
                        "threadTopic": conv_topic
                    }
                },
                "shift": {
                    "valid": 'true',
                    "subtitle": f"👀 show preview in large font",
                    "arg": myPreview
                }},
            "icon": {"path": ""},
            'arg': webLink
        })

    if not messages:
        result["items"].append({
            "title": empty_title,
            "subtitle": empty_subtitle,
            "arg": "",
            "icon": {"path": "icons/Warning.png"}
        })
    result['variables'] = {"mySource": "mailList"}
    print(json.dumps(result))


def fetchFolder ():
    """
    A function to fetch the folder name knowing its ID.
    Because going to the database using the function is too slow (and because the folders change rarely), decided to store the match in a JSON file in the data folder, and to update it every 30 days
    
    db = sqlite3.connect(OUTLOOK_DB_FILE)
    db.row_factory = sqlite3.Row
    rs = db.execute(f"SELECT Folder_Name from Folders WHERE Record_RecordID = {myFolderID}").fetchone()
    return rs[0]
    """

    with open(OUTLOOK_FOLDER_KEY_FILE, "r") as f:
        d = json.load(f)
    return d


def fetchAccounts ():
    
    with open(OUTLOOK_ACCOUNT_KEY_FILE, "r") as f:
        d = json.load(f)
    return d

def fetchContacts ():
    
    if CONTACT_AUTOCOMPLETE == "Database":
        OUTLOOK_CONTACTS_FILE = OUTLOOK_CONTACTS_LIST_FILE
    elif CONTACT_AUTOCOMPLETE in ["AddressBook","None"]:
        OUTLOOK_CONTACTS_FILE = OUTLOOK_CONTACTS_BOOK_FILE

    with open(OUTLOOK_CONTACTS_FILE, "r") as f:
        d = json.load(f)
    return d


def cleanSubject(mySubject):
    """
    a function to delete from subject used-defined strings which can take important real estate when results are returned.
    """
    for substring in WEED_TX:
            mySubject = mySubject.replace (substring, "")
    return mySubject

def showContacts(MY_INPUT, MY_CONTACTS,myQuery,myDirection):
        # If the contact source is unavailable/empty (e.g. legacy "Database"
        # autocomplete on new Outlook, where the cache is never built), skip
        # autocomplete and let the normal search run instead of dead-ending.
        if not MY_CONTACTS:
            return
        MYOUTPUT = {"items": []}
        mySubset = [i for i in MY_CONTACTS if MY_INPUT.casefold() in i.casefold()]
        myQueryQ = " ".join([w for w in myQuery.split() if not w.startswith(f"{myDirection}:")])


        # adding a complete contact if the user selects it from the list
        if mySubset:
            for thisContact in mySubset:
                thisContact_ = thisContact.replace(" ", "_")
                MYOUTPUT["items"].append({
                "title": f"{thisContact}",
                "subtitle": MY_INPUT,
                "arg": f"{myQueryQ} {myDirection}:{thisContact_} ",
                "variables" : {
                    
                    "myIter": True,
                    #"myQuery": myQuery,
                    "mySource": 'contacts',
                    
                    },
                "icon": {
                        "path": f"icons/contact.png"
                    }
                })
        else:
            MYOUTPUT["items"].append({
            "title": "no contacts matching",
            "subtitle": "try another query?",
            "variables" : {
                    
                    "myArg": MY_INPUT+" "
                    },
            "arg": "",
            "icon": {
                    "path": f"icons/Warning.png"
                }
            })
        print (json.dumps(MYOUTPUT))
        exit()


def showSavedQueries(MY_INPUT, SAVED_QUERIES,myQuery):
        MYOUTPUT = {"items": []}
        mySubset = [i for i in SAVED_QUERIES if MY_INPUT.casefold() in i['Name'].casefold()]
        myQueryQ = " ".join([w for w in myQuery.split() if not w.startswith("sq:")])


        # adding a complete contact if the user selects it from the list
        if mySubset:
            for thisSQ in mySubset:
                
                MYOUTPUT["items"].append({
                "title": f"{thisSQ['Name']}",
                "subtitle": thisSQ['Query'],
                "arg": f"{myQueryQ} {thisSQ['Query']} ",
                "variables" : {
                    
                    "myIter": True,
                    "mySource": 'saved_queries',
                    
                    },
                "icon": {
                        "path": f"icons/savedSearch.png"
                    }
                })
        else:
            MYOUTPUT["items"].append({
            "title": "no saved searches matching",
            "subtitle": "try another query?",
            "variables" : {
                    
                    "myArg": MY_INPUT+" "
                    },
            "arg": "",
            "icon": {
                    "path": f"icons/Warning.png"
                }
            })
        print (json.dumps(MYOUTPUT))
        exit()


def compileSQL(myQuery,myFolderKeys,myAccountKeys, myContacts):
    """
    a function to parse the user's input and generate an SQL query that can be used to query the database
    """
    
    myElements = myQuery.split()
    # if len(myElements) > 1:
    conditions = []
    DEFAULT_SORT = 'DESC'
    MYITER = os.getenv('myIter')
    MYSOURCE = os.getenv('mySource')
    
    
    for myElement in myElements:
        #log (f"myIter = {MYITER}")
        if myElement.startswith("from:"): #user is searching by sender
            if myElement == 'from:me': #use the user-defined name string
                    myString = MYSELF
            
            elif MYITER == None or (MYITER == '1' and MYSOURCE != "contacts"): #not coming from the contact autocomplete
                if CONTACT_AUTOCOMPLETE != "None":
                    showContacts (myElement.split(":")[1],myContacts,myQuery,"from")
                else: #no autocomplete
                    myString = myElement.split(":")[1].strip()
                    myString = myString.replace("_"," ")    
            else: #no showing contacts
                myString = myElement.split(":")[1].strip()
                myString = myString.replace("_"," ")
                    
            conditions.append (f"Message_SenderList LIKE '%{myString.strip()}%'")
                
        elif myElement.startswith("to:"): #user is searching by sender
            
            if myElement == 'to:me': #use the user-defined name string
                myString = MYSELF
            
            elif MYITER == None: #not coming from the contact autocomplete
                if CONTACT_AUTOCOMPLETE != "None":
                    showContacts (myElement.split(":")[1],myContacts,myQuery,"to")
                else: #no autocomplete
                    myString = myElement.split(":")[1].strip()
                    myString = myString.replace("_"," ")    
            else: #no showing contacts
                myString = myElement.split(":")[1].strip()
                myString = myString.replace("_"," ")

            conditions.append (f"Message_RecipientList LIKE '%{myString.strip()}%'")

        elif myElement.startswith("sq:"): #user is entering a saved query
            
            if MYITER == None: #not coming from the sq autocomplete
                showSavedQueries (myElement.split(":")[1],SAVED_QUERIES,myQuery)
                
            # else: #no showing contacts
            #     myString = myElement.split(":")[1].strip()
            #     myString = myString.replace("_"," ")

            # conditions.append (f"Message_RecipientList LIKE '%{myString.strip()}%'")

        elif myElement.startswith("cc:"): #user is searching by sender
            
            if myElement == 'cc:me': #use the user-defined name string
                myString = MYSELF
            
            else: 
                myString = myElement.split(":")[1].strip()
                myString = myString.replace("_"," ")
            conditions.append (f"Message_CCRecipientAddressList LIKE '%{myString}%'")

        elif myElement == ("has:attach"): #user is searching for messages with attachments
                 
           
           conditions.append (f"Message_HasAttachment = 1")
        
        elif myElement.startswith("subject:"): #user is searching by subject
            
           
           myString = myElement.split(":")[1].strip()
           conditions.append (f"Message_NormalizedSubject LIKE '%{myString}%'")
        
        elif myElement.startswith("since:"): #user wants emails from the past xxx days
            
           myString = myElement.split(":")[1].strip()
           timeSpan = checkTimespan (myString)
           conditions.append (f"Message_TimeReceived > {timeSpan}")
        

        elif myElement.startswith("is:"): #read or unread
            
            myString = myElement.split(":")[1].strip()
           
            if myString == 'read':
                conditions.append (f"Message_ReadFlag = 1")
            
            elif myString == 'unread':
                conditions.append (f"Message_ReadFlag = 0")
            
            elif myString == 'important':
                conditions.append (f"Record_Priority = 1")
            elif myString == 'unimportant':
                conditions.append (f"Record_Priority = 5")

        elif myElement.startswith("is:"): #read or unread
            
            myString = myElement.split(":")[1].strip()
           
            if myString == 'read':
                conditions.append (f"Message_ReadFlag = 1")
            
            elif myString == 'unread':
                conditions.append (f"Message_ReadFlag = 0")

        elif myElement.startswith("folder:"): #user is searching by subject
           
           myString = myElement.split(":")[1].strip()
           myString = myString.replace("_"," ")
           myFolderID = next((key for key, value in myFolderKeys.items() if value.casefold() == myString.casefold()), None)
           
           if myFolderID:
            
            conditions.append (f"Record_FolderID = {myFolderID}")
        
        elif myElement.startswith("account:"): #user is searching by exchange account
           
           myString = myElement.split(":")[1].strip()
           myString = myString.replace("_"," ")
           myAccountID = next((key for key, value in myAccountKeys.items() if myString.casefold() in value.casefold()), None)
           
           if myAccountID:
            
            conditions.append (f"Record_AccountUID = {myAccountID}")
        
        elif myElement == ("--a"): #user wants to sort by decreasing date
           DEFAULT_SORT = 'ASC'
        
        elif myElement.startswith("-"): #user wants to exclude a word from preview or subject
            
           
           myString = myElement[1:].strip()
           conditions.append(f"(Message_Preview NOT LIKE '%{myString}%' AND Message_NormalizedSubject NOT LIKE '%{myString}%')")
        


        else: #default search: subject and preview (might want to customize that in workflow configuration)
            conditions.append(f"(Message_Preview LIKE '%{myElement}%' OR Message_NormalizedSubject LIKE '%{myElement}%')")
        
    
    if EXCL_FOLDERS:
        for myExclFolder in EXCL_FOLDERS:
            #try: 
            myFolderID = next((key for key, value in myFolderKeys.items() if value == myExclFolder.strip()), None)
                
            #except:
            #    myFolderKeys = fetchFolder ()
            #   myFolderID = next((key for key, value in myFolderKeys.items() if value == myExclFolder.strip()), None)
  
            if myFolderID:
                conditions.append(f"Record_FolderID <> {myFolderID}")
    
    
    # joining all the conditions
    conditions_str = " AND ".join(conditions)
    conditions_str = f" WHERE {conditions_str}"
    
    SQL_SORT = f'ORDER BY Message_TimeSent {DEFAULT_SORT}'
    sql = f"SELECT * FROM Mail {conditions_str} {SQL_SORT}"
    log (sql)
        
    return sql
    

def handle(mySQL):
    

    db = sqlite3.connect(OUTLOOK_DB_FILE)
    db.row_factory = sqlite3.Row
    
    rs = db.execute(mySQL).fetchall()
    totCount = len(rs)
    
    myCounter = 0
    attIcon = ''
    for r in rs:
        myCounter += 1
        timeRec = timestampToText (r['Message_TimeSent'],"%Y-%m-%d %H:%M")
        timeRecF = timestampToText (r['Message_TimeSent'],"%A, %B %d, %Y %I:%M %p")
        if r['Message_ReadFlag'] == 0:
            readIcon = '⭐'
        else:
            readIcon = ''

        if r['Record_Priority'] == 1:
            urgentIcon = '‼️'
        elif r['Record_Priority'] == 5:
            urgentIcon = '🔽'
        
        else:
        
            urgentIcon = ''

        if r['Message_HasAttachment'] == 1:
            attIcon = '📎'
        if r['Message_NormalizedSubject']:
            MySubjectClean = cleanSubject(r['Message_NormalizedSubject'] )
        
        myPreview = f"{timeRecF}\nFrom: {r['Message_SenderList']}\nTo: {r['Message_RecipientList']}\n\nSubject: {MySubjectClean}\n\n{r['Message_Preview']}"
        myMessagePath = f"{OUTLOOK_MSG_FOLDER}{r['PathToDataFile']}"

        try: 
            myFolder = myFolderKeys [str(r['Record_FolderID'])]
        except:
            myFolderKeys = fetchFolder ()
            myFolder = myFolderKeys [str(r['Record_FolderID'])]
        if myFolder == "Snoozed":
            SnoozeString = "💤"
        else:
            SnoozeString = ""
        result["items"].append({
            "title": f"{readIcon}{MySubjectClean}{urgentIcon}",
            
            'subtitle': f"{myCounter}/{totCount:,} {attIcon} [{myFolder}] From: {r['Message_SenderList']} {timeRec} {SnoozeString}",
            'valid': True,
            "quicklookurl": '',
            # Reset myIter so opening a message routes to the "open" branch,
            # not the "recur" (re-search) branch used for autocomplete acceptance.
            'variables': {"myIter": ""},
                "mods": {

                "control": {
                    "valid": 'true',
                    "subtitle": f"🧵 filter entire thread",
                    "arg": r['Message_ThreadTopic'],
                    'variables': {
                        "mySource": 'thread',
                        "threadTopic": r['Message_ThreadTopic']
                    }
                },
                "shift": {
                    "valid": 'true',
                    "subtitle": f"👀 show preview in large font",
                    "arg": myPreview
                }},
            "icon": {
                "path": f""
            },
            'arg': myMessagePath
                }) 
        
    if not rs:
        result["items"].append({
            "title": "No matches in your library",
            "subtitle": "Try a different query",
            "arg": "",
            "icon": {
                "path": "icons/Warning.png"
                }
            
                })
    result['variables'] = {"mySource": "mailList"}                   
    print (json.dumps(result))


def main():
    MYSOURCE = os.getenv('mySource')
    log (SAVED_QUERIES)

    # Check if the snooze file has been updated today
    if checkJSON(OUTLOOK_SNOOZER_FILE):
        log("The JSON file has been updated today.")
    else:
        log("The JSON file has not been updated today. Updating....")
        # Discard the child's stdout/stderr: unSnoozer prints its own Alfred
        # JSON, which would otherwise corrupt this Script Filter's output.
        subprocess.run(["/usr/bin/python3", "unSnoozer.py"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    myQuery = sys.argv[1]

    if is_new_outlook():
        # --- Graph API path ---
        log("New Outlook detected, using Graph API")

        # Contacts are still useful for autocomplete
        try:
            myContacts
        except:
            try:
                myContacts = fetchContacts()
            except Exception:
                myContacts = []

        if MYSOURCE == "thread":
            thread_topic = os.getenv('threadTopic')
            handleGraph(myQuery, myContacts, is_thread=True, thread_topic=thread_topic)
        else:
            handleGraph(myQuery, myContacts)
    else:
        # --- Legacy SQLite path ---
        log("Legacy Outlook detected, using SQLite")
        checkingTime()
        try:
            myFolderKeys
        except:
            myFolderKeys = fetchFolder()

        try:
            myAccountKeys
        except:
            myAccountKeys = fetchAccounts()

        try:
            myContacts
        except:
            myContacts = fetchContacts()

        if MYSOURCE == "thread":
            myQuery = os.getenv('threadTopic')
            mySQL = f"SELECT * FROM Mail WHERE Message_ThreadTopic = '{myQuery}' ORDER BY Message_TimeSent ASC"
        elif MYSOURCE == "contacts":
            mySQL = compileSQL(myQuery, myFolderKeys, myAccountKeys, myContacts)
        elif myQuery:
            mySQL = compileSQL(myQuery, myFolderKeys, myAccountKeys, myContacts)
        else:
            ExclFolder = ""
            if EXCL_FOLDERS:
                for myExclFolder in EXCL_FOLDERS:
                    myFolderID = next((key for key, value in myFolderKeys.items() if value == myExclFolder.strip()), None)
                if myFolderID:
                    ExclFolder = f"WHERE Record_FolderID <> {myFolderID}"
            mySQL = f"SELECT * FROM Mail {ExclFolder} ORDER BY Message_TimeSent DESC"

        handle(mySQL)





if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        _sd = os.path.dirname(os.path.abspath(__file__))
        for _p in (os.path.join(_sd, "lib"), _sd):
            if _p not in sys.path:
                sys.path.insert(0, _p)
        from auth import MissingConfigError, setup_required_items_json
        if isinstance(e, MissingConfigError):
            print(setup_required_items_json())
        else:
            raise
