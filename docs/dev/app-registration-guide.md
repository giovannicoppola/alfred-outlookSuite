# Entra App Registration — Form Answers

Below are suggested values for your organization's Entra app registration form.

---

## Application Name

`Outlook Keyboard Shortcuts`

## Description

A local desktop productivity tool that manages Outlook emails (archive, delete, flag, move, snooze, create drafts) and reads calendar events via keyboard shortcuts, using the Microsoft Graph API. This replaces an existing AppleScript-based workflow that is incompatible with the new Outlook for Mac. The app runs locally on the user's Mac, only accesses the signed-in user's own mailbox and calendar, and stores no data externally.

## App Owner

*(your name)*

## App Developer/Engineer

*(your name)*

## Is the App Hosted by Your Organization?

Yes — the app runs locally on a user-managed Mac. It makes API calls only to Microsoft Graph (`https://graph.microsoft.com`), which is within the Microsoft 365 / organizational tenant environment. No external servers or third-party hosting is involved.

## Design/Config Document

Attach the file **`app-config-document.md`** (created separately in this folder). It covers all required details: API permissions, redirect URIs, platform configuration, authentication flow, and token configuration.

## Vendor Supplied Config/Info Links

- Microsoft Graph API documentation: `https://learn.microsoft.com/en-us/graph/overview`
- MSAL Python library (Microsoft official): `https://learn.microsoft.com/en-us/entra/msal/python/`
- Graph API Mail permissions reference: `https://learn.microsoft.com/en-us/graph/permissions-reference#mail-permissions`
- Graph API Calendar permissions reference: `https://learn.microsoft.com/en-us/graph/permissions-reference#calendars-permissions`

## SDER Note

Check with your IT contact whether this type of single-user, read/write-own-mailbox app requires SDER approval. Given the limited scope (no application permissions, no access to other users' data, no external hosting, read-only calendar access), it may be exempt — but confirm before submitting.
