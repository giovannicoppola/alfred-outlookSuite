# alfred-outlookSuite
***A suite of tools to interact with Microsoft Outlook via Alfred***



- *Note*: Many of the functions of this package will not work with the new, electron-based, 'New Outlook' interface, as AppleScript is not supported. Uncheck `New Outlook` in the `File` menu in Outlook to switch to the legacy version.

<a href="https://github.com/giovannicoppola/alfred-outlookSuite/releases/latest/">
<img alt="Downloads"
src="https://img.shields.io/github/downloads/giovannicoppola/alfred-outlookSuite/total?color=purple&label=Downloads"><br/>
</a>
<a href="https://alfred.app/workflows/giovanni/outlook-suite/">
<img alt="Gallery Downloads"
src="https://img.shields.io/badge/dynamic/json?url=https%3A%2F%2Fraw.githubusercontent.com%2Fgiovannicoppola%2Falfred-gallery-downloads%2Fmain%2Fdownloads.json&query=%24.outlook-suite%5B0%5D.display&label=Gallery%20Downloads&color=blue&logo=alfred"><br/>
</a>

![](images/alfred-outlookSuite.png)

<!-- MarkdownTOC autolink="true" bracket="round" depth="3" autoanchor="true" -->

- [Motivation](#motivation)
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



<h1 id="setting-up">Setting up ⚙️</h1>

### Needed
- Alfred 5 with Powerpack license
- *Note*: Many of the functions of this package will not work with the 'New Outlook' interface, as AppleScript is not supported. 
- *Note*: Your institution might have an automatic archival policy. In that case older mail messages that have been archived will not be returned in alfred-outlookSuite queries
- Python3 (howto [here](https://www.freecodecamp.org/news/python-version-on-mac-update/))
- Download `alfred-outlookSuite` [latest release](https://github.com/giovannicoppola/alfred-outlookSuite/releases/latest)



## Default settings 
- In Alfred, open the 'Configure Workflow' menu in `alfred-outlookSuite` preferences
- *Optional*:	
	- set the keyword (or hotkey) to launch `alfred-outlookSuite` (default: `olk`) 
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

- standard search: one or more search strings will search both the subject of an email, and its preview (first 250 characters). 
- gmail-like search strings, listed below, supported (remember to set the MYSELF variable in the Workflow Configuration). 
	- `from:` (including `from:me`). Replace space with underscore if you want to specify the whole name (e.g. `from:john_appleseed`)
	- `to:` (including `to:me`)
	- `cc:` (including `cc:me`)
	- `subject:`,  `has:attach`, `is:unread`, `is:read`, `is:important`, `is:unimportant`  
	- `folder:` (note: replace spaces with underscores if the folder name contains spaces, e.g. `folder:sent_items`)
	- `-text` to exclude text 
	- `account:` to filter by Exchange account
	- `--a` to sort by increasing date (oldest first)
	- `since:n` will return email received in the last `n` days. `w` and `m` are supported for months and weeks, respectively (e.g. `since:2w`).
	- `sq:` will open a list of saved queries
- Once an email of interest has been identified, the following actions are possible:
	- ↩️Enter will open the email in Outlook
	- ^-↩️ (control-enter) will show all the messages in the thread
	- ⇧-↩️ (shift-enter) will show in large font (and copy to clipboard)  the fields: From, To, Subject, and Preview from the selected email. 
 
## Draft a new email ⭐
- use a keyword (default: `em`) or a hotkey to launch, followed by text. Alfred will create a draft email with subject =  entered text, and save it  in the `Drafts` folder. 
![](images/draft.png)

## Email Saving 💾
- use a keyword (default: `ols`) or a hotkey to launch, while Outloook is the frontmost application and an email is selected. 
- Choose the destination folder using Alfred's file filter
- Save your email there. File will be renamed to include date and exclude special characters. A Markdown link to that email is copied to the clipboard. 



## Email Snoozing 💤
- Make sure you have, or create, a `Snoozed` folder in your main Outlook account. 
- use a keyword (default: `olz`) or a hotkey to launch, while Outloook is the frontmost application and an email is selected. 
- enter the number of days you want to snooze your email. Alfred will show the corresponding date, and the number of emails already snoozed for that date. 
- Selecting the result will 1) Snooze the email until the desired date and 2) move that email to the `Snoozed` folder. 
- the `checkSnoozed` keyword will show in large font (and copy to clipboard) an overview of all the snoozed email (one line per day) 
- Unsnoozing will happen once a day, as soon as `alfred-outlookSuite` is launched. You can force the unsnooze script using a keyword (default: `olu`) or a hotkey. 

![](images/snooze.png)


## Folder database refresh 🔄
- will occur according to the rate in days set in `alfred-outlookSuite` preferences
- `outlook::refresh` to force folder database refresh


<h1 id="known-issues">Limitations & known issues ⚠️</h1>

- None for now, but I have not done extensive testing, let me know if you see anything!
- Again, a lot of this will not work with the electron interface (`New Outlook`). Uncheck `New Outlook` in the `File` menu in Outlook to switch to the previous version.



<h1 id="acknowledgments">Acknowledgments 😀</h1>

- Thanks to the [Alfred forum](https://www.alfredforum.com) community!
- Thanks to [@xeric](https://github.com/xeric) for the first version and the inspiration to build upon it. 
- Abacus icon: https://www.flaticon.com/free-icon/abacus_1046277
- Snooze icon: https://www.flaticon.com/free-icon/snooze_3602333
- https://www.flaticon.com/free-icon/chapter_6348276
- https://www.flaticon.com/free-icon/diskette_907027
- https://www.flaticon.com/free-icon/save_4371273
	
<h1 id="changelog">Changelog 🧰</h1>

- 2023-09-08 version 0.9.4: bug fix
- 2023-06-10 version 0.9.3: contact autocomplete, saved queries
- 2023-06-07 version 0.9.1: support for multiple Exchange accounts
- 2023-06-03 version 0.9: complete rewrite, renamed `OutlookSuite`
- 2022-12-05 version 0.3 (Alfred 5)
- 2022-06-27 version 0.2 (Python 3)


<h1 id="feedback">Feedback 🧐</h1>

Feedback welcome! If you notice a bug, or have ideas for new features, please feel free to get in touch either here, or on the [Alfred](https://www.alfredforum.com) forum. 

