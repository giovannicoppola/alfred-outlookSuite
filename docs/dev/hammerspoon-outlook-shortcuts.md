# Hammerspoon Configuration

## Modules

| File | Description |
|------|-------------|
| `init.lua` | Entry point. Loads modules, binds hotkeys. |
| `outlook_keys.lua` | Vim-style navigation in Outlook message list |
| `outlook_rsvp.lua` | RSVP to calendar invitations via AX automation |
| `capslock_remap.lua` | Caps Lock remapping (requires Caps Lock → Control in System Settings) |
| `Spoons/ClipboardImageSaver.spoon` | Save clipboard images to disk |

## Outlook Shortcuts

All Outlook shortcuts auto-detect **legacy** vs **new** Outlook for Mac and use the correct AX interactions. No configuration needed — just works with whichever version is running.

### Message List Navigation

Requires the message list to be focused (use Ctrl+Shift+M to focus it).

Keys use Gallium keyboard layout positions (QWERTY y/n/j/w map to Gallium j/k/h/l).

| Shortcut | Action |
|----------|--------|
| `j` | Move down (next message) |
| `k` | Move up (previous message) |
| `h` | Collapse conversation |
| `l` | Expand conversation |
| `Ctrl+f` | Page down (~10 messages) |
| `Ctrl+u` | Page up (~10 messages) |
| `gg` | Go to first message |
| `G` | Go to last message |

### RSVP (Outlook must be frontmost)

| Shortcut | Action |
|----------|--------|
| `Cmd+Shift+A` | Accept silently (no notification to organizer) |
| `Cmd+Shift+T` | Tentative silently |
| `Cmd+Shift+D` | Decline silently |
| `Cmd+Shift+X` | Remove from calendar |

Select a meeting invitation in the reading pane before pressing.

### UI

| Shortcut | Action |
|----------|--------|
| `Ctrl+Shift+M` | Focus the message list |
| `Cmd+/` | Toggle left sidebar (folder tree) |
| `Cmd+\` | Toggle right panel (My Day / calendar / tasks) |

### Other

| Shortcut | Action |
|----------|--------|
| `Cmd+Option+2` | Capture screen area and save image (ClipboardImageSaver) |

## Legacy vs New Outlook Detection

Detection is based on the presence of `AXToolbar` in the window's top-level AX children:
- **Legacy**: has `AXToolbar` → uses `AXOutline` / `AXPopUpButton` interactions
- **New**: no `AXToolbar` → uses `AXTable` / `AXButton` interactions

The result is cached and invalidated when Outlook relaunches.

### Key AX Differences

| Element | Legacy | New |
|---------|--------|-----|
| Message list | `AXOutline` desc="Message List Outline View" | `AXTable` desc="Message List" |
| RSVP buttons | `AXPopUpButton` → show menu → click item | `AXButton` → AXPress |
| Silent/send toggle | Menu item in popup | `AXCheckBox` "Email organizer switch" |
| Page navigation | Home/End/PageUp/PageDown | Burst of arrow keys (Home/End don't move selection) |
| Go to first | Home | Cmd+Opt+Up then Down (skip date header) |
| Go to last | End | Cmd+Opt+Down then Up (skip "Load more conversations") |

## RSVP via CLI

The RSVP module can also be invoked directly via the `hs` CLI:

```sh
hs -c "require('outlook_rsvp').run('accept', 'silent')"
hs -c "require('outlook_rsvp').run('accept', 'send')"
hs -c "require('outlook_rsvp').run('tentative', 'silent')"
hs -c "require('outlook_rsvp').run('tentative', 'send')"
hs -c "require('outlook_rsvp').run('decline', 'silent')"
hs -c "require('outlook_rsvp').run('decline', 'send')"
hs -c "require('outlook_rsvp').run('remove')"
```

(`hs` = `/Applications/Hammerspoon.app/Contents/Frameworks/hs/hs`)
