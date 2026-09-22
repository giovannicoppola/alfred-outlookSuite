-- Vim-style j/k navigation in Outlook message list
-- Auto-detects legacy vs new Outlook and uses the correct AX detection.
--
-- When the Outlook message list is focused:
--   j/k  → Down/Up arrows (navigate messages)
--   h    → Left arrow (collapse conversation)
--   l    → toggle: expand a collapsed conversation, or collapse an open one
--          (works even when selection is on a message inside the thread)
--   f/u  → Page Down/Page Up (scroll message list)
--   gg   → Go to first message (Home)
--   G    → Go to last message (End)
-- In compose, search, reading pane, or any other app, keys type normally.
--
-- New Outlook quirks:
--   Home/End/PageUp/PageDown scroll the view without moving selection.
--   gg  → Cmd+Opt+Up (jumps to top)
--   G   → Cmd+Opt+Down (jumps to last message)
--   f/u → burst of arrow keys (new Outlook ignores PageUp/PageDown for selection)
--
-- Detection uses a two-tier check:
--   1. Fast: is Outlook the frontmost app?
--   2. AX:  is the focused UI element the message list?
--
-- Legacy Outlook: AXOutline with AXDescription "Message List Outline View"
-- New Outlook:    AXTable with AXDescription "Message List"
--                 (focused element may be AXWindow, so we search for the table)

local OUTLOOK_BUNDLE = "com.microsoft.Outlook"

-- Temporary: set true to log why keys pass through. Toggle from console:
--   _G.DEBUG_OUTLOOK_KEYS = true
DEBUG_OUTLOOK_KEYS = DEBUG_OUTLOOK_KEYS or false

-- Bare keys: vim-style navigation (QWERTY h/j/k/l)
local bareKeyMap = {
    [4]  = "left",  -- h (collapse conversation)
    [38] = "down",  -- j
    [40] = "up",    -- k
    [37] = "right", -- l (expand conversation)
}

-- Ctrl+ keys: page navigation
-- Legacy: send PageDown/PageUp; New: send burst of arrows
local ctrlKeyMap = {
    [3]  = "down", -- Ctrl+f (page down)
    [32] = "up",   -- Ctrl+u (page up)
}
local ctrlKeyMapLegacy = {
    [3]  = "pagedown", -- Ctrl+f
    [32] = "pageup",   -- Ctrl+u
}

local PAGE_JUMP_COUNT = 10 -- number of arrow keys per "page" in new Outlook

-- gg detection: track last 'g' press time
local G_KEYCODE = 5  -- QWERTY g = Gallium g
local GG_TIMEOUT = 0.3 -- seconds to press g twice
local lastGTime = 0

-- New Outlook (Chromium/WebView) no longer auto-exposes its accessibility
-- tree. Until AXManualAccessibility is set on the app, AXFocusedUIElement
-- reports a bare AXWindow and the tree is empty, so message-list detection
-- fails and j/k/h/l fall through as normal typing. Set it once per launch.
local accessibilityEnabled = false
local function ensureAccessibility(app)
    if accessibilityEnabled then return end
    local axApp = hs.axuielement.applicationElement(app)
    if not axApp then return end
    axApp:setAttributeValue("AXManualAccessibility", true)
    -- Only latch as done if the set actually took (the watcher may fire before
    -- the app is ready); otherwise the next keystroke retries.
    if axApp:attributeValue("AXManualAccessibility") == true then
        accessibilityEnabled = true
    end
end

-- Detect whether the running Outlook is new or legacy.
-- Decide by which message-list element the window exposes:
--   New:    AXTable   with AXDescription "Message List"
--   Legacy: AXOutline with AXDescription "Message List Outline View"
-- (The old heuristic keyed off the absence of an AXToolbar, but current
-- New Outlook builds DO expose an AXToolbar, so that misdetected new as
-- legacy and broke j/k/h/l. The message-list role is the reliable signal.)
-- Caches the result until Outlook relaunches (see watcher below).
local outlookVersionCache = nil
local function isNewOutlook(app)
    if outlookVersionCache ~= nil then return outlookVersionCache end
    ensureAccessibility(app)
    local axApp = hs.axuielement.applicationElement(app)
    local win = axApp:attributeValue("AXFocusedWindow")
    if not win then return false end

    local function classify(el, depth)
        if depth > 10 then return nil end
        local r = el:attributeValue("AXRole") or ""
        local d = el:attributeValue("AXDescription") or ""
        if r == "AXTable" and d == "Message List" then return true end
        if r == "AXOutline" and d == "Message List Outline View" then return false end
        local children = el:attributeValue("AXChildren")
        if children then
            for _, child in ipairs(children) do
                local v = classify(child, depth + 1)
                if v ~= nil then return v end
            end
        end
        return nil
    end

    local isNew = classify(win, 0)
    if isNew == nil then return false end -- can't tell yet; don't cache, retry next key
    outlookVersionCache = isNew
    return isNew
end

-- Invalidate caches when Outlook relaunches, and re-enable accessibility.
local watcher = hs.application.watcher.new(function(name, event, app)
    if app and app:bundleID() == OUTLOOK_BUNDLE then
        if event == hs.application.watcher.terminated then
            outlookVersionCache = nil
            accessibilityEnabled = false
        elseif event == hs.application.watcher.launched then
            outlookVersionCache = nil
            accessibilityEnabled = false
            ensureAccessibility(app)
        end
    end
end)
watcher:start()

-- Enable accessibility for an already-running Outlook at load time.
local runningOutlook = hs.application.get(OUTLOOK_BUNDLE)
if runningOutlook then
    ensureAccessibility(runningOutlook)
end

local function isMessageListFocusedLegacy()
    local focused = hs.axuielement.systemWideElement():attributeValue("AXFocusedUIElement")
    if not focused then return false end
    local role = focused:attributeValue("AXRole")
    local desc = focused:attributeValue("AXDescription")
    return role == "AXOutline" and desc == "Message List Outline View"
end

local function isMessageListFocusedNew(app)
    local focused = hs.axuielement.systemWideElement():attributeValue("AXFocusedUIElement")
    if not focused then return false end
    local role = focused:attributeValue("AXRole")
    local desc = focused:attributeValue("AXDescription")

    if role == "AXTable" and desc == "Message List" then
        return true
    end

    -- New Outlook may report AXWindow as focused; check the AXTable directly
    if role == "AXWindow" then
        local axApp = hs.axuielement.applicationElement(app)
        local win = axApp:attributeValue("AXFocusedWindow")
        if not win then return false end

        local function findTable(el, depth)
            if depth > 8 then return nil end
            local r = el:attributeValue("AXRole") or ""
            local d = el:attributeValue("AXDescription") or ""
            if r == "AXTable" and d == "Message List" then return el end
            local children = el:attributeValue("AXChildren")
            if children then
                for _, child in ipairs(children) do
                    local found = findTable(child, depth + 1)
                    if found then return found end
                end
            end
            return nil
        end

        local tbl = findTable(win, 0)
        if tbl and tbl:attributeValue("AXFocused") then
            return true
        end
    end

    return false
end

local function isMessageListFocused()
    local app = hs.application.frontmostApplication()
    if not app or app:bundleID() ~= OUTLOOK_BUNDLE then
        return false
    end

    if isNewOutlook(app) then
        return isMessageListFocusedNew(app)
    else
        return isMessageListFocusedLegacy()
    end
end

-- Locate the message-list container (AXTable new / AXOutline legacy).
local function findMessageList(app)
    local axApp = hs.axuielement.applicationElement(app)
    local win = axApp:attributeValue("AXFocusedWindow")
    if not win then return nil end
    local result
    local function walk(el, depth)
        if depth > 10 or result then return end
        local r = el:attributeValue("AXRole") or ""
        local d = el:attributeValue("AXDescription") or ""
        if (r == "AXTable" and d == "Message List")
            or (r == "AXOutline" and d == "Message List Outline View") then
            result = el
            return
        end
        local children = el:attributeValue("AXChildren")
        if children then for _, c in ipairs(children) do walk(c, depth + 1) end end
    end
    walk(win, 0)
    return result
end

-- Info about the currently selected row: a stable fingerprint, whether it is a
-- conversation (expandable), and its disclosure state if the OS exposes one.
local function selectedRowInfo(app)
    local tbl = findMessageList(app)
    if not tbl then return nil end
    local sel = tbl:attributeValue("AXSelectedRows")
    local row = sel and sel[1]
    if not row then return nil end
    local idx = row:attributeValue("AXIndex")
    local cells = row:attributeValue("AXChildren")
    local desc = (cells and cells[1] and cells[1]:attributeValue("AXDescription"))
        or row:attributeValue("AXDescription") or ""
    local disclosing = row:attributeValue("AXDisclosing") -- boolean on legacy, nil on new
    return {
        fp = tostring(idx) .. "|" .. desc:sub(1, 40),
        isConv = (disclosing ~= nil) or (desc:find("Conversation,") ~= nil),
        disclosing = disclosing,
    }
end

local tap = hs.eventtap.new({ hs.eventtap.event.types.keyDown }, function(event)
    local keycode = event:getKeyCode()
    local flags = event:getFlags()

    local app = hs.application.frontmostApplication()
    if not app or app:bundleID() ~= OUTLOOK_BUNDLE then
        return false
    end
    if DEBUG_OUTLOOK_KEYS then print(string.format("[okeys] tap fired: keycode=%d", keycode)) end
    local newOutlook = isNewOutlook(app)

    -- Shift+g (G): go to last message
    if keycode == G_KEYCODE and flags.shift and not flags.cmd and not flags.alt and not flags.ctrl and not flags.fn then
        if isMessageListFocused() then
            if newOutlook then
                hs.eventtap.keyStroke({"cmd", "alt"}, "down", 0)
                -- Skip past "Load more conversations" row at the bottom
                hs.eventtap.keyStroke({}, "up", 0)
            else
                hs.eventtap.keyStroke({}, "end", 0)
            end
            return true
        end
        return false
    end

    -- gg: go to first message (bare g pressed twice quickly)
    if keycode == G_KEYCODE and not flags.cmd and not flags.alt and not flags.ctrl and not flags.shift and not flags.fn then
        if isMessageListFocused() then
            local now = hs.timer.secondsSinceEpoch()
            if (now - lastGTime) < GG_TIMEOUT then
                if newOutlook then
                    hs.eventtap.keyStroke({"cmd", "alt"}, "up", 0)
                else
                    hs.eventtap.keyStroke({}, "home", 0)
                end
                lastGTime = 0
                return true
            end
            lastGTime = now
            return true -- consume first g
        end
        return false
    end

    -- Ctrl+ keys: page navigation
    if flags.ctrl and not flags.cmd and not flags.alt and not flags.shift and not flags.fn then
        if newOutlook then
            local direction = ctrlKeyMap[keycode]
            if direction and isMessageListFocused() then
                for _ = 1, PAGE_JUMP_COUNT do
                    hs.eventtap.keyStroke({}, direction, 0)
                end
                return true
            end
        else
            local action = ctrlKeyMapLegacy[keycode]
            if action and isMessageListFocused() then
                hs.eventtap.keyStroke({}, action, 0)
                return true
            end
        end
        return false
    end

    -- Bare keys: j/k/h/l navigation
    local bareAction = bareKeyMap[keycode]
    if not bareAction then
        if DEBUG_OUTLOOK_KEYS then print(string.format("[okeys] keycode %d not in bareKeyMap, passing through", keycode)) end
        return false
    end

    if flags.cmd or flags.alt or flags.ctrl or flags.shift or flags.fn then
        if DEBUG_OUTLOOK_KEYS then print("[okeys] modifier held, passing through") end
        return false
    end

    if DEBUG_OUTLOOK_KEYS then
        local f = hs.axuielement.systemWideElement():attributeValue("AXFocusedUIElement")
        print(string.format("[okeys] bare '%s' (kc=%d) newOutlook=%s focusedRole=%s focusedDesc=[%s] listFocused=%s",
            bareAction, keycode, tostring(newOutlook),
            f and (f:attributeValue("AXRole") or "?") or "nil",
            f and (f:attributeValue("AXDescription") or "") or "",
            tostring(isMessageListFocused())))
    end

    if isMessageListFocused() then
        -- l is a toggle: expand a collapsed conversation, collapse an open one
        -- (even when selection is on a message inside the thread, matching h).
        if bareAction == "right" then
            local info = selectedRowInfo(app)
            local dir
            if info and info.disclosing ~= nil then
                -- Legacy exposes real state: collapse if open, else expand.
                dir = info.disclosing and "left" or "right"
            elseif info and info.isConv then
                -- New Outlook, on a conversation header → expand.
                dir = "right"
            else
                -- New Outlook, on a message row (inside an open thread) → collapse.
                dir = "left"
            end
            hs.eventtap.keyStroke({}, dir, 0)
            return true
        end
        hs.eventtap.keyStroke({}, bareAction, 0)
        return true
    end

    return false
end)

tap:start()

return {
    start = function() tap:start() end,
    stop = function() tap:stop() end,
}
