# alfred-outlook
***A suite of tools to interact with Microsoft Outlook via Alfred***



- *Note*: Many of the functions of this package will not work with the new, electron-based, 'New Outlook' interface, as AppleScript is not supported. Uncheck `New Outlook` in the `File` menu in Outlook to switch to the legacy version.

<a href="https://github.com/giovannicoppola/alfred-outlook/releases/latest/">
<img alt="Downloads"
src="https://img.shields.io/github/downloads/giovannicoppola/alfred-outlook/total?color=purple&label=Downloads"><br/>
</a>

![](images/alfred-outlookSuite.png)

<!-- MarkdownTOC autolink="true" bracket="round" depth="3" autoanchor="true" -->

- [Motivation](#motivation)
- [New Outlook vs Legacy Outlook](#new-vs-legacy)
- [RSVP to meeting invitations (New Outlook)](#rsvp)
- [Requesting admin access (New Outlook)](#admin-access)
- [Setting up](#setting-up)
- [Basic Usage](#usage)
- [Known Issues](#known-issues)
- [Acknowledgments](#acknowledgments)
- [Changelog](#changelog)
- [Feedback](#feedback)

<!-- /MarkdownTOC -->


<h1 id="motivation">Motivation ✅</h1>

- Quickly list, search, and filter your Microsoft Outlook emails 
- Perform basic tasks like: email snoozing, quick email drafting, email saving. 



<h1 id="new-vs-legacy">New Outlook vs Legacy Outlook 🆚</h1>

Microsoft's **New Outlook** (the Electron-based interface) dropped AppleScript and the local SQLite database that the original workflow relied on. To keep working, the New Outlook build talks to Microsoft's **Graph API** (OAuth2) instead, while the **released** workflow still uses AppleScript + SQLite against **Legacy Outlook**.

You can switch between the two interfaces via the `New Outlook` toggle in Outlook's `File` menu.

> ⚠️ **Heads up — New Outlook is a moving target.** New Outlook for Mac changed its
> UI, accessibility structure, and behavior **multiple times during development of
> this workflow, and is very likely to change again.** Features that read Outlook's
> on-screen state through the macOS accessibility interface — detecting the selected
> email for **saving (`ols`)** and **snoozing (`olz`)**, the **RSVP** actions, and
> calendar capture — are the most exposed and can break after an Outlook update
> (typically failing silently or doing nothing). If something that worked before
> suddenly stops, an Outlook update is the usual cause; it may need a fix in the
> workflow's detection code. The pure Graph API features (search, filters, drafts)
> are more stable, as they don't depend on Outlook's UI.

### Feature comparison

| Feature | Legacy Outlook (released) | New Outlook (this version) | Notes |
|---|---|---|---|
| Email search (free text, `from:`/`to:`/`cc:`, `subject:`, `has:attach`, `is:read`/`is:unread`, `is:important`, `since:`, `folder:`, `-text` to exclude, `--a`, `sq:`) | ✅ | ✅ | Legacy matches subject+preview; New Outlook uses Graph `$search` (matches the **full body** + subject/sender/recipients) with `$filter`, some filters resolved client-side. `-text` exclusion filters on subject+preview only. |
| `account:` filter (multiple Exchange accounts) | ✅ | ❌ | Graph API is single-account only |
| Open email from results | ✅ Opens in Outlook desktop app (local `.olk` file) | ⚠️ Opens `webLink` in browser (Outlook Web) | No API to open a specific message in the New Outlook desktop app |
| Thread view (`^↩`) | ✅ | ✅ | New version uses `conversationId` |
| Show email fields in large font (`⇧↩`) | ✅ | ✅ | |
| Draft new email (`em`) | ✅ Opens draft in desktop app | ✅ Saves a draft silently in Drafts (not opened) + notification | New version is for jotting a subject to finish later, not composing now |
| Save email as file (`ols` = save only, `olsa` = save + archive) | ✅ | ✅ | New version is a two-step capture-then-save flow; hotkey needs a 2s delay |
| Email snoozing (`olz`) | ✅ | ✅ | `snoozer.json` is cross-compatible between versions |
| Conversation snooze | ✅ | ✅ | |
| Unsnoozing (auto + `olu`) | ✅ | ✅ | |
| Beeminder posting on unsnooze | ✅ | ✅ | Posts the current inbox count to your Beeminder goal when emails are unsnoozed |
| Snoozed overview (`checkSnoozed`) | ✅ | ✅ | Reads `snoozer.json` only |
| Contact autocomplete (AddressBook / Database) | ✅ | ✅ | |
| Folder database refresh (`outlook::refresh`) | ✅ | ➖ | Not needed; New Outlook queries folders live via Graph |
| Browse calendar events / meeting notes | ❌ | ✅ | New-only; uses Graph `calendarView` |
| RSVP to meeting invitations (accept/tentative/decline, silent or notify; remove) | ❌ | ✅ | New-only; `rsvp` keyword. Requires [Hammerspoon](https://www.hammerspoon.org/) (optional dependency) |
| Vim-style message-list navigation (`j`/`k`/`h`/`l`, `gg`/`G`) | ✅ | ✅ | Bundled Hammerspoon module (`outlook_keys.lua`), installed via `olhs`; not an Alfred keyword/hotkey. Auto-detects Legacy/New; `l` toggles a conversation open/closed |
| Authentication | None (local database) | OAuth2 via Microsoft Entra app registration | New version requires one-time browser login |
| Extra dependencies | None (built-in `sqlite3`) | `msal`, `requests` (bundled in `lib/`); Hammerspoon (optional, only for RSVP) | |

✅ = supported · ⚠️ = supported with differences · ❌ = not available · ➖ = not applicable

**Search coverage caveats (New Outlook):**
- **Primary mailbox only.** Search reaches the messages in your live mailbox. It does **not** cover the **Online Archive** (In-Place Archive) — Microsoft's Graph API does not expose the archive mailbox. If your institution auto-archives older mail, those messages won't appear here even though Outlook's own search finds them. For older/archived mail, use Outlook's built-in search.
- **Result ceiling.** Results are paged up to `MAX_RESULTS` (default 200). Filter-only queries are additionally served from a local cache of the 50 most recent messages (see `CACHE_MINUTES`), so e.g. `has:attach` only scans those 50.
- **Relevance vs. date.** Free-text/sender/subject searches use Graph `$search`, which returns by *relevance*; results are re-sorted by date client-side, so a very old match may fall outside the fetched set.

<h1 id="rsvp">RSVP to meeting invitations (New Outlook) 📅</h1>

Respond to the meeting invitation currently open in Outlook's reading pane
without leaving the keyboard.

1. Select/open the invitation in Outlook.
2. Trigger the `rsvp` keyword (or assign it a hotkey) and pick a response:
   Accept, Tentative, or Decline — each **silent** (no response sent) or
   **notify** (sends a response) — or **Remove from calendar** for a cancelled
   invite.

You can also drive RSVP straight from the keyboard: add
`require("outlook_rsvp").bindDefaultHotkeys()` to your `~/.hammerspoon/init.lua`
to enable **⌘⇧A** (accept), **⌘⇧T** (tentative), **⌘⇧D** (decline), and
**⌘⇧X** (remove) — active only while Outlook is frontmost.

**This feature requires [Hammerspoon](https://www.hammerspoon.org/)** (an
optional dependency — everything else in the workflow works without it). RSVP
buttons live in Outlook's UI with no Graph API equivalent, so the response is
performed by driving Outlook's accessibility interface from Hammerspoon.

### One-time Hammerspoon setup (RSVP only)
- Install [Hammerspoon](https://www.hammerspoon.org/) and grant it Accessibility
  permission (System Settings → Privacy & Security → Accessibility).
- Run the **`olhs`** keyword in Alfred. It copies the bundled
  `outlook_rsvp.lua` into `~/.hammerspoon/` for you and reports what (if anything)
  is still missing — Hammerspoon install, the `hs` IPC line, the module load, and
  the Accessibility grant.
- Add these two lines to your `~/.hammerspoon/init.lua` (the workflow will **not**
  edit your Hammerspoon config for you), then reload Hammerspoon:

  ```lua
  require("hs.ipc")        -- required: enables the `hs` command-line calls
  require("outlook_rsvp")  -- optional: only for in-Hammerspoon hotkeys
  ```
- The module auto-detects legacy vs new Outlook via accessibility inspection; no
  further configuration is needed. Re-run `olhs` any time to re-check the setup.

> Note: the vim-style message-list navigation (`j`/`k`/`h`/`l`, `gg`/`G`) is a
> pure Hammerspoon feature — it relies on bare single-key hotkeys, which Alfred
> can't express — but its module (`outlook_keys.lua`) is now **bundled** and
> installed by the same `olhs` setup. Add `require("outlook_keys")` to your
> `init.lua` to load it whenever Hammerspoon starts.

> ⚠️ **Fragility caveat.** RSVP works by driving New Outlook's *accessibility
> interface* — it locates the response buttons ("Accept", "Decline", "Remove",
> the "Email organizer" toggle, etc.) by their position and description in the
> message-header area. Microsoft changes New Outlook's UI/accessibility structure
> between builds, so these actions **will break from time to time** (a button
> stops being found, or the wrong element is clicked) and need re-tweaking in
> `outlook_rsvp.lua`. This is inherent to any AX-based automation of New Outlook —
> there is no Graph API for responding to invitations. If an RSVP action silently
> does nothing, that's the likely cause; check Hammerspoon's console for a
> "button not found" message. The core Graph-API mail features are unaffected.

The New Outlook build reaches your mailbox through the Microsoft Graph API, which in a corporate tenant requires your IT team to register a public client app and give you two credentials (Client ID, Tenant ID — no client secret). You can't enable this yourself.

**👉 See [ADMIN_ACCESS.md](ADMIN_ACCESS.md)** for a plain-English explanation of what to ask for, why it's safe (delegated, your-mailbox-only access), and a copy-paste request you can send straight to IT.

Once you have the two credentials, follow [SETUP.md](SETUP.md) to finish configuration. The legacy Outlook workflow needs none of this.

<h1 id="setting-up">Setting up ⚙️</h1>

### Needed
- Alfred 5 with Powerpack license
- **Hammerspoon required for selected-email actions (New Outlook only).** The actions that operate on the *currently selected* email — **Email Saving (`ols`)** and **Email Snoozing (`olz`)** — must read which message is selected. New Outlook renders its message list inside a web-content accessibility subtree that a plain script can't reach, so the workflow delegates the read to **[Hammerspoon](https://www.hammerspoon.org/)** (the same optional dependency used for RSVP), which can read it via the native accessibility API. Install Hammerspoon, grant **it** Accessibility permission, and load the bundled `outlook_rsvp.lua` module (see the [RSVP setup](#rsvp) — same module provides `getSelection`). With that in place, **Alfred itself does not need Accessibility permission**. Without Hammerspoon, `ols`/`olz` on New Outlook fail with *"No cached email info found"*. Search, drafting, filters, and other Graph-only actions need none of this.
- *Note*: Many of the functions of this package will not work with the 'New Outlook' interface, as AppleScript is not supported. 
- *Note*: Your institution might have an automatic archival policy. In that case older mail messages that have been archived will not be returned in alfred-outlook queries
- Python3 (howto [here](https://www.freecodecamp.org/news/python-version-on-mac-update/))
- Download `alfred-outlook` [latest release](https://github.com/giovannicoppola/alfred-outlook/releases/latest)



## Default settings 
- In Alfred, open the 'Configure Workflow' menu in `alfred-outlook` preferences
- *Optional*:	
	- set the keyword (or hotkey) to launch `alfred-outlook` (default: `olk`) 
	- set the keyword (or hotkey) to force-refresh the folder list (default: `outlook::refresh`)
	- set the keyword (or hotkey) to force-the unsnoozing script (default: `olu`)
	- set your name, so that it can be used in `from:me`, `to:me`, and `cc:me` searches
	- *Text to hide in subject* Enter here comma-separated text strings that you don't want to see in the 'Subject' result (e.g. [External], <External> etc.), therefore saving important real estate in Alfred's output.
	- *Folders to exclude* List here, comma-separated, the folders you would like to exclude from the search. Default: `Deleted Items`. 
	- *Folders list refresh rate*	- The folder list is generated periodically (default: 30 days) to improve performance, as folders change less often. Enter your preferred number of days here.
	- *Snooze file location* - Leave empty if you use this Workflow on one computer (it will be saved in the Workflow's data folder). Enter a shared folder location here in case you want to share the file snoozing information across computers using the same Outlook account.
	- *Beeminder info* if you use [Beeminder](https://www.beeminder.com/) to track your inbox, enter your account information here and Alfred will post the number of messages in inbox every time an unsnooze script is run
	- *Contact autocomplete* Choose between `None` (no autocomplete, default), `AddressBook`: autocomplete list generated from the Address Book, or `Database`: autocomplete list generated from the senders in the database. 
	- *Saved Queries* enter here queries that can be inboked using the `sq:` trigger. Enter multiple queries in this format: `RecentUnread = to:me is:unread since:20`, separated by semicolon (`;`)	


<h1 id="usage">Basic Usage 📖</h1>

## Search your email 🔍

- standard search:
	- **Legacy Outlook:** search strings match the subject and the preview (first 250 characters).
	- **New Outlook:** search strings use Graph `$search`, which matches the **full message body** (plus subject, sender, and recipients), not just the preview.
- gmail-like search strings, listed below, supported (remember to set the MYSELF variable in the Workflow Configuration). 
	- `from:` (including `from:me`). Replace space with underscore if you want to specify the whole name (e.g. `from:john_appleseed`)
	- `to:` (including `to:me`)
	- `cc:` (including `cc:me`)
	- `subject:`,  `has:attach`, `is:unread`, `is:read`, `is:important`, `is:unimportant`  
	- `folder:` (note: replace spaces with underscores if the folder name contains spaces, e.g. `folder:sent_items`)
	- `-text` to exclude text (drops results whose **subject or preview** contains the text; in New Outlook this is a client-side filter on the preview, so it does not see body-only occurrences)
	- `account:` to filter by Exchange account
	- `--a` to sort by increasing date (oldest first)
	- `since:n` will return email received in the last `n` days. `w` and `m` are supported for months and weeks, respectively (e.g. `since:2w`).
	- `sq:` will open a list of saved queries
- Once an email of interest has been identified, the following actions are possible:
	- ↩️Enter will open the email in Outlook
	- ^-↩️ (control-enter) will show all the messages in the thread
	- ⇧-↩️ (shift-enter) will show in large font (and copy to clipboard)  the fields: From, To, Subject, and Preview from the selected email. 
 
## Draft a new email ⭐
- use a keyword (default: `em`) or a hotkey to launch, followed by text. Alfred creates a draft email with the entered text as the subject and saves it in the `Drafts` folder.
- Think of this as a quick way to **jot down an email to write later** — capture the subject/reminder now, then open Outlook when you're ready to compose the body and send it.
- **Legacy Outlook** opens the new draft in the desktop app. **New Outlook** saves the draft *silently* (it is not opened) and shows a "Draft saved to Outlook" notification — there is no reliable way to open a specific New Outlook draft for composing from here.
![](images/draft.png)

## Email Saving 💾
- Two keywords, while Outlook is frontmost and an email is selected:
	- **`ols`** — **save only**: saves the email as a file, leaving it in place.
	- **`olsa`** — **save and archive**: saves the email *and* moves it to the `Archive` folder.
- Choose the destination folder using Alfred's file filter
- Save your email there. File will be renamed to include date and exclude special characters. A Markdown link to that email is copied to the clipboard. 



## Email Snoozing 💤
- Make sure you have, or create, a `Snoozed` folder in your main Outlook account. 
- use a keyword (default: `olz`) or a hotkey to launch, while Outloook is the frontmost application and an email is selected. 
- enter the number of days you want to snooze your email. Alfred will show the corresponding date, and the number of emails already snoozed for that date. 
- Selecting the result will 1) Snooze the email until the desired date and 2) move that email to the `Snoozed` folder. 
- the `checkSnoozed` keyword will show in large font (and copy to clipboard) an overview of all the snoozed email (one line per day) 
- Unsnoozing will happen once a day, as soon as `alfred-outlook` is launched. You can force the unsnooze script using a keyword (default: `olu`) or a hotkey. 

![](images/snooze.png)


## Folder database refresh 🔄
- will occur according to the rate in days set in `alfred-outlook` preferences
- `outlook::refresh` to force folder database refresh


<h1 id="known-issues">Limitations & known issues ⚠️</h1>

**New Outlook (Graph API):**
- **Online Archive is not searchable.** Search covers your primary mailbox only. Auto-archived / In-Place Archive mail is not exposed by the Graph API, so older messages that Outlook's own search finds may return nothing here. Use Outlook's built-in search for archived mail.
- **`account:` filter is not supported** — the Graph API operates on a single mailbox.
- **Opening a message** opens its `webLink` in the browser (Outlook Web), not the desktop app — there is no API to open a specific message in the New Outlook desktop app.
- Search results are capped at `MAX_RESULTS` (default 200) and `$search` orders by relevance before client-side date sorting — see the [search caveats](#new-vs-legacy) above.

**Legacy Outlook (AppleScript + SQLite):**
- Requires the classic interface — uncheck `New Outlook` in Outlook's `File` menu.
- Your institution's archival policy may move older mail out of the local database, in which case it won't be returned in queries.

Not much other testing has been done — please report anything you find!



<h1 id="acknowledgments">Acknowledgments 😀</h1>

- Thanks to the [Alfred forum](https://www.alfredforum.com) community!
- Thanks to [@xeric](https://github.com/xeric) for the first version and the inspiration to build upon it. 
- Abacus icon: https://www.flaticon.com/free-icon/abacus_1046277
- Snooze icon: https://www.flaticon.com/free-icon/snooze_3602333
- https://www.flaticon.com/free-icon/chapter_6348276
- https://www.flaticon.com/free-icon/diskette_907027
- https://www.flaticon.com/free-icon/save_4371273
	
<h1 id="changelog">Changelog 🧰</h1>

- 2026-08-27 version 0.10.0: **New Outlook (Microsoft Graph API) support** with automatic Legacy/New detection across all commands; calendar event browsing & meeting notes; RSVP to meeting invitations (`rsvp`, requires Hammerspoon); bundled `msal`/`requests` dependencies
- 2023-09-08 version 0.9.4: bug fix
- 2023-06-10 version 0.9.3: contact autocomplete, saved queries
- 2023-06-07 version 0.9.1: support for multiple Exchange accounts
- 2023-06-03 version 0.9: complete rewrite, renamed `OutlookSuite`
- 2022-12-05 version 0.3 (Alfred 5)
- 2022-06-27 version 0.2 (Python 3)


<h1 id="feedback">Feedback 🧐</h1>

Feedback welcome! If you notice a bug, or have ideas for new features, please feel free to get in touch either here, or on the [Alfred](https://www.alfredforum.com) forum. 

