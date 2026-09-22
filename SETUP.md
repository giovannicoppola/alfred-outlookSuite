# Setup Guide (New Outlook / Graph API)

This guide covers configuring the workflow for the **New Outlook for Mac**, which uses the Microsoft Graph API. If you use **Legacy Outlook**, none of this is required — just install the workflow.

Before starting, get a **Client ID** and **Tenant ID** from your IT team. Most setups need no client secret (a public-client app using PKCE); but if IT gives you a **client secret** with a confidential app, that works too — just add it in Step 1. See [ADMIN_ACCESS.md](ADMIN_ACCESS.md) for how to request them.

---

## Prerequisites

- macOS with [Alfred](https://www.alfredapp.com) + Powerpack
- Python 3.9+ — macOS's built-in `/usr/bin/python3` works (the bundled dependencies are pure-Python; no compiler or Homebrew needed)
- Microsoft Outlook for Mac (New or Legacy)
- Accessibility permission for Alfred (System Settings → Privacy & Security → Accessibility)
- Client ID and Tenant ID from your admin (see [ADMIN_ACCESS.md](ADMIN_ACCESS.md))

---

## Step 1: Create the config file

Create `src/config.json` inside the workflow with the credentials from your admin:

```json
{
    "client_id": "YOUR_CLIENT_ID_HERE",
    "tenant_id": "YOUR_TENANT_ID_HERE",
    "scopes": ["Mail.ReadWrite", "Calendars.Read"]
}
```

> **Given a client secret?** If your admin registered a *confidential* app and handed you a secret, add a `"client_secret": "YOUR_SECRET_HERE"` line to the JSON above. The workflow detects it and uses the confidential flow automatically — and the registration then does **not** need "Allow public client flows" enabled. Leave it out for the public-client (PKCE) path.

**This file is gitignored and must never be committed or shared.** The same applies to `src/token_cache.json`, which is generated automatically after you sign in.

---

## Step 2: Confirm bundled dependencies

The Graph API path needs `msal` and `requests`, which are bundled in `src/lib/`. The scripts add that folder to their path automatically, so nothing to install in most cases.

If `src/lib/` is missing (e.g. a fresh clone), install the dependencies into it:

```bash
pip install --target=src/lib msal requests
```

> Note: this also pulls in `cryptography`/`cffi`/`pycparser` (via `PyJWT[crypto]`), which the public-client sign-in flow does **not** use. They ship compiled, architecture-specific binaries, so to keep `src/lib` pure-Python (and working on both Apple Silicon and Intel), delete those three folders afterward.

---

## Step 3: First-run authentication

The first search (keyword `olk`) triggers a one-time sign-in:

1. Your browser opens the Microsoft login page.
2. Sign in with your organization credentials (including MFA if required).
3. The browser redirects to `http://localhost:5050`; a temporary local server captures the response.
4. You see "Authentication complete. You can close this tab."
5. Results appear in Alfred.

A `src/token_cache.json` file is created. Subsequent runs use the cached token silently. When the refresh token eventually expires, the browser login repeats once.

---

## Step 4: Grant Accessibility permission

Features that read the selected email or calendar event (snooze, save, event notes) use the macOS Accessibility API:

1. Open **System Settings → Privacy & Security → Accessibility**.
2. Ensure **Alfred** is listed and enabled.

Without this, selection-based scripts fail silently.

---

## Troubleshooting

**"Config file not found"** — create `src/config.json` (Step 1).

**Browser opens but sign-in fails** — verify the redirect URI in the app registration is exactly `http://localhost:5050`, that the app allows public client flows (platform "Mobile and desktop applications"), that port 5050 is free, and that admin consent was granted for the delegated permissions.

**No results, or an older email that Outlook finds is missing** — search covers your primary mailbox only. The **Online Archive** is not exposed by the Graph API. Use Outlook's built-in search for archived mail. See the [limitations](README.md#known-issues).

**Sign-in fails with `AADSTS7000218` ("request body must contain client_assertion or client_secret")** — your app registration is a *confidential* app that requires a secret, but no `client_secret` is in `config.json`. Either add the `client_secret` (Step 1), or have your admin enable **"Allow public client flows"** on the registration.

**Login keeps prompting / token errors** — delete `src/token_cache.json` and re-authenticate; the browser sign-in runs once more and caches a fresh token. (On the public-client path there's no secret to expire; on the confidential path, a *rotated/expired client secret* also causes this — update `config.json` with the new secret.)

**Works in Terminal but not from Alfred** — confirm Alfred has Accessibility permission (Step 4).
