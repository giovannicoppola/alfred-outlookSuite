# Requesting Access from Your IT / System Administrator

The **New Outlook** version of this workflow talks to your mailbox through Microsoft's official **Graph API** instead of the old local database. In a corporate environment, you can't turn that access on yourself — your IT team has to register a small "application" in Microsoft Entra (formerly Azure AD) and hand you two credentials (a Client ID and a Tenant ID).

This page explains, in plain English, exactly what to ask for and why. There's a copy-paste request at the bottom you can send to IT as-is.

---

## Why this is needed

- The New (Electron-based) Outlook for Mac removed the local AppleScript/database access the legacy workflow relied on.
- The only supported way to read your mail programmatically now is the Microsoft Graph API.
- Graph access requires an **app registration** in your organization's tenant. Only an administrator can create one (or approve it).

You are **not** asking for any special or elevated access to other people's data. The app acts *as you*, sees only *your* mailbox and calendar, and only after *you* personally sign in with your normal corporate login (including MFA). This is the same "delegated" model used by any third-party mail or calendar app.

---

## What to ask IT to create

Ask them to register **one application** in Microsoft Entra with these settings:

| Setting | Value to request | Plain-English meaning |
|---|---|---|
| Application name | e.g. `Alfred Outlook (personal use)` | A label so IT can find it later |
| Account type | **Single tenant** (this organization only) | Nobody outside your company can use it |
| Platform | **Mobile and desktop applications** | It runs on your Mac, not a web server |
| Redirect URI | `http://localhost:5050` | Where the sign-in flow returns *on your own machine* |
| Allow public client flows | **Yes** (no client secret) | Lets this desktop app sign in with PKCE instead of a stored secret |

### Permissions to request (all **Delegated**)

"Delegated" is the key word — it means the app can only ever do what *you* are already allowed to do, and only while *you* are signed in. It gets **no** standing access to anyone's mailbox.

| Permission | What it allows | Why the workflow needs it |
|---|---|---|
| `User.Read` | Read your own basic profile | Sign-in |
| `Mail.ReadWrite` | Read/move/flag **your** email | Search, snooze (move to a folder), save, draft |
| `Calendars.Read` | Read **your** calendar | Browse events / meeting notes |

> If your IT team is cautious and you don't need snoozing or saving, you can ask for the read-only `Mail.Read` instead of `Mail.ReadWrite` — but snooze/save/draft features will not work.

Depending on company policy, an admin may also need to click **"Grant admin consent"** on those permissions so you aren't blocked at sign-in.

---

## What IT gives back to you

You should receive **two values** (both safe to receive by your normal internal channels):

1. **Client ID** (a.k.a. Application ID) — looks like `9e88dd7f-29b7-4eaa-9dbe-e61485390310`
2. **Tenant ID** (a.k.a. Directory ID) — looks like `3e9aadf8-6a16-490f-8dcd-c68860caae0b`

Ideally there's **no client secret** — a public-client app signs in with PKCE, so there's nothing to store or rotate. But if your organization's standard is a **confidential app with a client secret** instead, that works too: you'll get a third value (the secret) to add to `config.json` (see [SETUP.md](SETUP.md)), and in that case the app does **not** need "Allow public client flows" enabled.

---

## What you do with those two values

Put them into a `config.json` file in the workflow's `src/` folder:

```json
{
    "client_id": "YOUR_CLIENT_ID_HERE",
    "tenant_id": "YOUR_TENANT_ID_HERE",
    "scopes": ["Mail.ReadWrite", "Calendars.Read"]
}
```

This file is **gitignored** and must never be committed or shared. The first time you run a search, your browser opens the normal Microsoft sign-in once; after that it's silent. Full technical steps are in [SETUP.md](SETUP.md).

---

## Copy-paste request for your IT team

> **Subject: Request to register a personal Microsoft Entra app for Outlook (Graph API, delegated)**
>
> Hi,
>
> I'd like to use an Alfred (macOS) workflow to search and manage **my own** Outlook mail and calendar. The New Outlook for Mac no longer supports the local automation it used to, so it needs the Microsoft Graph API, which requires an app registration in our tenant.
>
> Could you please register a single-tenant app with the following, or let me know the approved process?
>
> - **Name:** Alfred Outlook (personal use)
> - **Account type:** Single tenant
> - **Platform:** Mobile and desktop applications
> - **Redirect URI:** `http://localhost:5050`
> - **Allow public client flows:** Yes (public client using PKCE — no client secret needed)
> - **Delegated API permissions:** `User.Read`, `Mail.ReadWrite`, `Calendars.Read`
>   (all *delegated* — the app acts only as me, only after I sign in with MFA; no application-level or mailbox-wide access)
> - Please **grant admin consent** for these permissions if our policy requires it.
>
> When it's set up, I'll need the **Client ID** and **Tenant ID**.
>
> Thanks!

---

## If IT says no

Some organizations don't allow personal app registrations at all. In that case:

- Ask whether there's an **already-approved** internal client ID for Graph/delegated mail access you can reuse.
- If nothing is available, the Graph-based (New Outlook) features can't run. Your only fallback is to switch Outlook back to **Legacy** mode (`File` menu → uncheck *New Outlook*) and use the released legacy workflow, which needs no admin access. See the [feature comparison](README.md#new-vs-legacy) for what differs between the two.
