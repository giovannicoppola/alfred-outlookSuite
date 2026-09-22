-- outlook_rsvp.lua
-- RSVP to calendar invitations in Outlook for Mac
-- Auto-detects legacy vs new Outlook and uses the correct AX interactions.
-- Actions: accept, tentative, decline, remove
-- Default: no response sent to organizer (silent RSVP)
--
-- Legacy Outlook:
--   - RSVP via AXPopUpButton (Accept/Tentative/Decline) → AXShowMenu → menu item
--   - Header path: splitter group 1 of splitter group 1 of splitter group 1, search by desc
--
-- New Outlook:
--   - RSVP via AXButton (Accept/Tentative/Decline) → AXPress
--   - Silent/send via AXCheckBox "Email organizer switch" (value 1=send, 0=silent)
--   - Header path: splitter group 1 > splitter group 1 > splitter group 1, search by desc

local M = {}

local function runAppleScript(script, successMsg, failMsg)
    local ok, _, raw = hs.osascript.applescript(script)
    if ok then
        print(successMsg)
        return true
    else
        local errMsg = failMsg
        if raw and raw.NSLocalizedDescription then
            errMsg = errMsg .. ": " .. raw.NSLocalizedDescription
        end
        print("ERROR: " .. errMsg)
        return false
    end
end

-- Detect new vs legacy Outlook at runtime.
-- Legacy Outlook's front window has a command-ribbon AXToolbar (many buttons).
-- New Outlook (16.11x+) exposes only a vestigial title-bar AXToolbar with NO
-- child elements, so the presence of *any* AXToolbar is no longer a reliable
-- signal — we require a NON-EMPTY toolbar to call it legacy. (Checking for any
-- toolbar used to work but broke when New Outlook started shipping the empty
-- title-bar toolbar, misdetecting New as legacy and running the wrong path.)
local function detectNewOutlook()
    local script = [[
        tell application "System Events"
            tell process "Microsoft Outlook"
                set w to front window
                set isNew to true
                repeat with el in UI elements of w
                    if role of el is "AXToolbar" then
                        if (count of UI elements of el) > 0 then
                            set isNew to false
                            exit repeat
                        end if
                    end if
                end repeat
                return isNew
            end tell
        end tell
    ]]
    local ok, result = hs.osascript.applescript(script)
    if ok then return result end
    return true
end

-- Independent confirmation that this is *legacy* Outlook, used as a hard safety
-- net before any recursive hs.axuielement walk (see removeLegacy). Unlike
-- detectNewOutlook's UI heuristic, this probes the defining difference directly:
-- legacy Outlook implements the AppleScript object model (answers "count of
-- accounts" with a number); New Outlook does not and raises an error. Returns
-- true ONLY when we get a positive legacy answer, so any ambiguity is treated
-- as "not legacy" and the dangerous walk is skipped.
local function isLegacyOutlook()
    local ok, result = hs.osascript.applescript(
        [[tell application "Microsoft Outlook" to return count of accounts]])
    return ok and type(result) == "number"
end

---------------------------------------------------------------------------
-- Legacy Outlook
---------------------------------------------------------------------------

local FIND_HEADER_LEGACY = [[
        set sg to splitter group 1 of splitter group 1 of splitter group 1 of front window
        set headerArea to missing value
        repeat with i from 1 to count of UI elements of sg
            if description of UI element i of sg is "Message Header Details" then
                set headerArea to UI element i of sg
                exit repeat
            end if
        end repeat
        if headerArea is missing value then error "Message Header Details not found"
]]

local function rsvpLegacy(action, sendResponse)
    local titles = {
        accept    = "Accept",
        tentative = "Tentative",
        decline   = "Decline",
    }

    local title = titles[action]
    if not title then
        print("ERROR: Unknown RSVP action: " .. tostring(action))
        return
    end

    local menuItem = sendResponse and "Respond Without Comments" or "Do Not Send a Response"
    local suffix = sendResponse and " (notified)" or ""

    local script = string.format([[
        tell application "System Events"
            tell process "Microsoft Outlook"
%s
                set btn to first pop up button of headerArea whose title is "%s"
                perform action "AXShowMenu" of btn
                delay 0.2
                click menu item "%s" of menu 1 of btn
            end tell
        end tell
    ]], FIND_HEADER_LEGACY, title, menuItem)

    runAppleScript(script, "Meeting " .. action .. "ed" .. suffix,
        "No '" .. title .. "' button found — is a meeting invite selected?")
end

local function removeLegacy()
    -- HARD SAFETY NET (independent of detectNewOutlook): the recursive
    -- hs.axuielement walk below hangs Hammerspoon's main thread — freezing the
    -- keyboard and wedging Outlook — if it ever runs against New Outlook's
    -- Electron web-content subtree. Before walking, confirm via an independent
    -- signal that this really is legacy Outlook; if not, refuse rather than hang.
    if not isLegacyOutlook() then
        print("ERROR: Not legacy Outlook — refusing recursive AX walk (would hang). "
            .. "This is the New Outlook remove path's job; check detection.")
        return
    end

    local app = hs.application.get("Microsoft Outlook")
    if not app then
        print("ERROR: Outlook not running")
        return
    end

    local axApp = hs.axuielement.applicationElement(app)
    local win = axApp:attributeValue("AXFocusedWindow")
    if not win then
        print("ERROR: No focused window")
        return
    end

    local function findRemoveButton(el, depth)
        if depth > 15 then return nil end
        local role = el:attributeValue("AXRole") or ""
        local desc = el:attributeValue("AXDescription") or ""

        if role == "AXButton" then
            if desc == "No Response Required" or desc:find("^Remove") then
                local size = el:attributeValue("AXSize")
                -- Make sure it's a real button (not tiny 2x2 ones in tables)
                if size and size.w > 10 and size.h > 10 then
                    return el
                end
            end
        end

        local children = el:attributeValue("AXChildren")
        if children then
            for _, child in ipairs(children) do
                local found = findRemoveButton(child, depth + 1)
                if found then return found end
            end
        end
        return nil
    end

    local btn = findRemoveButton(win, 0)
    if not btn then
        print("ERROR: No 'Remove' or 'No Response Required' button found — is a meeting invite selected?")
        return
    end

    -- Click using mouse position (AXPress doesn't work for these buttons)
    local pos = btn:attributeValue("AXPosition")
    local size = btn:attributeValue("AXSize")
    if pos and size then
        local currentPos = hs.mouse.absolutePosition()
        local x = pos.x + size.w / 2
        local y = pos.y + size.h / 2
        hs.mouse.absolutePosition({x = x, y = y})
        hs.eventtap.leftClick({x = x, y = y})
        hs.timer.usleep(100000)
        hs.mouse.absolutePosition(currentPos)
    end

    print("Dismissed")
end

---------------------------------------------------------------------------
-- New Outlook
---------------------------------------------------------------------------

-- Navigate to the "Message Header Details" scroll area. New Outlook nests the
-- reading pane under a variable number of splitter groups (build-dependent), so
-- we descend through nested split groups until we find the level that contains
-- it, bounded to a few hops. This stays within the chrome (no web-content tree
-- walk), so it can't trigger the AX hang.
local FIND_HEADER_NEW = [[
        set sg to splitter group 1 of splitter group 1 of splitter group 1 of front window
        set headerArea to missing value
        repeat 8 times
            set nextSG to missing value
            repeat with i from 1 to count of UI elements of sg
                set el to UI element i of sg
                set elDesc to ""
                try
                    set elDesc to description of el
                end try
                if elDesc is "Message Header Details" then
                    set headerArea to el
                    exit repeat
                end if
                if (nextSG is missing value) and (role of el is "AXSplitGroup") then
                    set nextSG to el
                end if
            end repeat
            if headerArea is not missing value then exit repeat
            if nextSG is missing value then exit repeat
            set sg to nextSG
        end repeat
        if headerArea is missing value then error "Message Header Details not found"
]]

local function setEmailOrganizerToggle(sendResponse)
    local desiredValue = sendResponse and 1 or 0
    return string.format([[
                try
                    set cb to first checkbox of headerArea whose title is "Email organizer switch"
                    if (value of cb as integer) is not %d then
                        click cb
                        delay 0.15
                    end if
                end try
]], desiredValue)
end

-- True when (x, y) lands on some connected display.
local function pointOnScreen(x, y)
    for _, s in ipairs(hs.screen.allScreens()) do
        local f = s:fullFrame()
        if x >= f.x and x < f.x + f.w and y >= f.y and y < f.y + f.h then
            return true
        end
    end
    return false
end

-- New Outlook RSVP/Remove work by synthesizing a real mouse click at the
-- button's global screen coordinates (AXPress does not delete/RSVP on these web
-- buttons). If the Outlook window is stranded off-screen — e.g. a disconnected
-- external monitor left it parked in coordinate space with no display — macOS
-- clamps the click to the screen edge and it either does nothing or hits the
-- wrong thing. Refuse and tell the user instead of clicking blind.
local function clickAt(cx, cy)
    if not pointOnScreen(cx, cy) then
        hs.alert.show("Outlook is off-screen — move its window back on-screen (Window ▸ Fill), then retry.")
        print("ERROR: RSVP target (" .. cx .. "," .. cy .. ") is off every display; refusing to click.")
        return false
    end
    local currentPos = hs.mouse.absolutePosition()
    hs.eventtap.leftClick({ x = cx, y = cy })
    hs.mouse.absolutePosition(currentPos)
    return true
end

-- RSVP to a meeting invite in New Outlook.
--
-- Same constraint as removeNew: NO recursive AX tree walk (it hangs on New
-- Outlook's web-content subtree). We navigate to the bounded header area via
-- System Events, optionally toggle the "Email organizer switch" there, then read
-- the RSVP button's screen rect and dispatch a real mouse click at its center.
local function rsvpNew(action, sendResponse)
    local descs = {
        accept    = "Accept",
        tentative = "Tentative",
        decline   = "Decline",
    }

    local desc = descs[action]
    if not desc then
        print("ERROR: Unknown RSVP action: " .. tostring(action))
        return
    end

    -- In the header area: set the organizer-notify checkbox, then return the
    -- target button's rect. Button is matched by description or title == desc.
    local script = string.format([[
        tell application "System Events"
            tell process "Microsoft Outlook"
%s
%s
                set btn to missing value
                repeat with i from 1 to count of UI elements of headerArea
                    set el to UI element i of headerArea
                    if role of el is "AXButton" then
                        set d to ""
                        set t to ""
                        try
                            set d to description of el
                        end try
                        try
                            set t to title of el
                        end try
                        if d is "%s" or t is "%s" then
                            set btn to el
                            exit repeat
                        end if
                    end if
                end repeat
                if btn is missing value then error "button not found"
                set p to position of btn
                set s to size of btn
                return ((item 1 of p) as text) & "," & ((item 2 of p) as text) & "," & ((item 1 of s) as text) & "," & ((item 2 of s) as text)
            end tell
        end tell
    ]], FIND_HEADER_NEW, setEmailOrganizerToggle(sendResponse), desc, desc)

    local ok, rect = hs.osascript.applescript(script)
    if not ok or type(rect) ~= "string" then
        print("ERROR: No '" .. desc .. "' button found — is a meeting invite selected?")
        return
    end

    local x, y, w, h = rect:match("(-?%d+),(-?%d+),(-?%d+),(-?%d+)")
    if not x then
        print("ERROR: could not parse button rect: " .. tostring(rect))
        return
    end
    local cx = tonumber(x) + tonumber(w) / 2
    local cy = tonumber(y) + tonumber(h) / 2

    if not clickAt(cx, cy) then return end

    local suffix = sendResponse and " (notified)" or ""
    print("Meeting " .. action .. "ed" .. suffix)
end

-- Remove a cancelled / no-response-required meeting from the calendar.
--
-- IMPORTANT: New Outlook is Electron/web-based. Recursively walking its AX tree
-- with hs.axuielement hangs indefinitely on a node in the web-content subtree
-- (it blocks Hammerspoon's main thread, freezing the keyboard and wedging
-- Outlook). So we must NOT walk the tree. Instead we navigate directly to the
-- bounded "Message Header Details" area via System Events to read the target
-- button's screen rect, then dispatch a REAL mouse click at its center.
-- AXPress does not actually delete the invite on these buttons; a real click does.
local function removeNew()
    -- Locate the button within the (bounded) header area and return its rect.
    local script = string.format([[
        tell application "System Events"
            tell process "Microsoft Outlook"
%s
                set btn to missing value
                repeat with i from 1 to count of UI elements of headerArea
                    set el to UI element i of headerArea
                    if role of el is "AXButton" then
                        set d to ""
                        try
                            set d to description of el
                        end try
                        if d starts with "Remove" or d is "No Response Required" then
                            set btn to el
                            exit repeat
                        end if
                    end if
                end repeat
                if btn is missing value then error "Remove button not found"
                set p to position of btn
                set s to size of btn
                return ((item 1 of p) as text) & "," & ((item 2 of p) as text) & "," & ((item 1 of s) as text) & "," & ((item 2 of s) as text)
            end tell
        end tell
    ]], FIND_HEADER_NEW)

    local ok, rect = hs.osascript.applescript(script)
    if not ok or type(rect) ~= "string" then
        print("ERROR: No 'Remove' button found — is a cancelled meeting selected?")
        return
    end

    local x, y, w, h = rect:match("(-?%d+),(-?%d+),(-?%d+),(-?%d+)")
    if not x then
        print("ERROR: could not parse button rect: " .. tostring(rect))
        return
    end
    local cx = tonumber(x) + tonumber(w) / 2
    local cy = tonumber(y) + tonumber(h) / 2

    if not clickAt(cx, cy) then return end

    print("Removed from calendar")
end

---------------------------------------------------------------------------
-- Public API
---------------------------------------------------------------------------

-- Single entry point: action = "accept"|"tentative"|"decline"|"remove", mode = "silent"|"send"
function M.run(action, mode)
    local isNew = detectNewOutlook()

    if action == "remove" then
        if isNew then removeNew() else removeLegacy() end
        return
    end

    local sendResponse = (mode == "send")
    if isNew then
        rsvpNew(action, sendResponse)
    else
        rsvpLegacy(action, sendResponse)
    end
end

-- Invoked via Alfred workflows using the hs CLI:
--
--   hs -c "require('outlook_rsvp').run('accept', 'silent')"
--   hs -c "require('outlook_rsvp').run('accept', 'send')"
--   hs -c "require('outlook_rsvp').run('tentative', 'silent')"
--   hs -c "require('outlook_rsvp').run('tentative', 'send')"
--   hs -c "require('outlook_rsvp').run('decline', 'silent')"
--   hs -c "require('outlook_rsvp').run('decline', 'send')"
--   hs -c "require('outlook_rsvp').run('remove')"
--
-- (hs = /Applications/Hammerspoon.app/Contents/Frameworks/hs/hs)

-- Read the description of the currently selected message-list row.
--
-- Used by the Alfred save (`ols`) and snooze (`olz`) scripts: New Outlook renders
-- the message list inside a web-content accessibility subtree that a plain
-- osascript/JXA walk can't reach, so those scripts shell out to `hs` and let
-- Hammerspoon read the selection via the native AX API (with AXManualAccessibility
-- set). This also means Alfred itself does not need Accessibility permission —
-- Hammerspoon does the reading.
--
--   hs -c "return require('outlook_rsvp').getSelection()"
--
-- Returns the selected row's description string (sender/subject/time), or "".
function M.getSelection()
    local app = hs.application.get("com.microsoft.Outlook")
    if not app then return "" end
    local ax = hs.axuielement.applicationElement(app)
    if not ax then return "" end
    pcall(function() ax:setAttributeValue("AXManualAccessibility", true) end)
    local win = ax:attributeValue("AXFocusedWindow")
    if not win then return "" end

    local tbl
    local function find(el, d)
        if d > 10 or tbl then return end
        local r = el:attributeValue("AXRole") or ""
        local ds = el:attributeValue("AXDescription") or ""
        if (r == "AXTable" and ds == "Message List")
            or (r == "AXOutline" and ds == "Message List Outline View") then
            tbl = el
            return
        end
        local ch = el:attributeValue("AXChildren")
        if ch then for _, c in ipairs(ch) do find(c, d + 1) end end
    end
    find(win, 0)
    if not tbl then return "" end

    local sel = tbl:attributeValue("AXSelectedRows")
    if not sel or #sel == 0 then return "" end
    local cells = sel[1]:attributeValue("AXChildren")
    local desc = cells and cells[1] and cells[1]:attributeValue("AXDescription")
    return desc or sel[1]:attributeValue("AXDescription") or ""
end

-- Optional keyboard shortcuts (opt-in). Enable once from ~/.hammerspoon/init.lua:
--   require("outlook_rsvp").bindDefaultHotkeys()
-- Binds Cmd+Shift+A/T/D/X to accept/tentative/decline/remove, and only acts
-- while Outlook is frontmost. Kept OUT of module load on purpose: the Alfred
-- `rsvp` action does `require('outlook_rsvp').run(...)`, so auto-binding on
-- require would grab these global combos as a side effect of that path too.
function M.bindDefaultHotkeys()
    if M._hotkeysBound then return end  -- idempotent
    M._hotkeysBound = true
    local function guarded(action, mode)
        return function()
            local app = hs.application.frontmostApplication()
            if app and app:bundleID() == "com.microsoft.Outlook" then
                M.run(action, mode)
            end
        end
    end
    hs.hotkey.bind({ "cmd", "shift" }, "a", guarded("accept", "silent"))
    hs.hotkey.bind({ "cmd", "shift" }, "t", guarded("tentative", "silent"))
    hs.hotkey.bind({ "cmd", "shift" }, "d", guarded("decline", "silent"))
    hs.hotkey.bind({ "cmd", "shift" }, "x", guarded("remove"))
end

return M
