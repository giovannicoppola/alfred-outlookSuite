# Outlook Keyboard Shortcuts — App Configuration Document

## 1. Overview

| Item | Detail |
|------|--------|
| App Name | Outlook Keyboard Shortcuts |
| Purpose | Manage Outlook emails (archive, delete, flag, move, snooze, create drafts) and read calendar events via keyboard shortcuts |
| App Type | Public client (desktop application) |
| Hosting | Local — runs on the user's macOS machine, no external servers |
| Data Storage | No external data storage; only a local OAuth token cache on the user's machine |
| Users | Single user (the app owner/developer) |

## 2. Platform Configuration

| Setting | Value |
|---------|-------|
| Platform | Mobile and desktop applications |
| Supported account types | Single tenant (your organization only) |
| Allow public client flows | Yes |

## 3. Redirect URIs

| URI | Purpose |
|-----|---------|
| `http://localhost` | MSAL interactive login — opens browser, redirects back to localhost after authentication |

No web, iOS, or Android redirect URIs are needed.

## 4. API Permissions

All permissions are **delegated** (the app acts on behalf of the signed-in user, never as an independent application).

| Permission | Type | Justification |
|------------|------|---------------|
| `User.Read` | Delegated | Required for sign-in and basic profile |
| `Mail.Read` | Delegated | Read email metadata (subject, sender, date, read status) |
| `Mail.ReadWrite` | Delegated | Move emails between folders, delete, flag/unflag, mark read/unread, create drafts |
| `Calendars.Read` | Delegated | Read calendar events (today's schedule, current meeting details) |

### Permissions NOT requested

| Permission | Reason not needed |
|------------|-------------------|
| `Mail.Send` | The app does not send emails (drafts are created, not sent) |
| `Mail.Read` (Application) | No application-level access — only delegated |
| `Calendars.ReadWrite` | The app only reads calendar events, never creates or modifies them |
| `Contacts.*` | No contacts access needed |

## 5. Authentication Details

| Setting | Value |
|---------|-------|
| Authentication flow | Authorization code with PKCE (Proof Key for Code Exchange) |
| Client type | Public client (no client secret) |
| Client secret | None — not required for public client apps |
| Certificate | None |
| Library | MSAL for Python (Microsoft Authentication Library, official Microsoft SDK) |

### Authentication flow description

1. On first run, the app opens the system browser to the Microsoft login page.
2. The user signs in with their organization credentials (including MFA if configured).
3. Microsoft redirects to `http://localhost` with an authorization code.
4. MSAL exchanges the code for access and refresh tokens using PKCE.
5. Tokens are cached locally in a JSON file on the user's machine.
6. On subsequent runs, MSAL uses the cached refresh token for silent authentication (no browser needed).
7. If the refresh token expires, the interactive browser login repeats.

## 6. Token Configuration

| Setting | Value |
|---------|-------|
| Token type | Bearer (OAuth 2.0) |
| Token endpoint | `https://login.microsoftonline.com/{tenant-id}/oauth2/v2.0/token` |
| Token cache location | Local file: `token_cache.json` in the app directory on the user's machine |
| Token lifetime | Default Microsoft settings (typically ~1 hour for access tokens, longer for refresh tokens) |
| Optional claims | None required |

## 7. API Endpoints Used

All calls go to `https://graph.microsoft.com/v1.0`:

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `/me/messages` | GET | List recent emails |
| `/me/messages/{id}` | GET | Get a specific email's details |
| `/me/messages/{id}` | PATCH | Flag/unflag, mark read/unread |
| `/me/messages/{id}` | DELETE | Delete an email |
| `/me/messages/{id}/move` | POST | Move email to archive or another folder |
| `/me/mailFolders` | GET | List available mail folders |
| `/me/messages/{id}/$value` | GET | Download full email as raw MIME (.eml file) |
| `/me/messages/{id}/attachments` | GET | List and download email attachments |
| `/me/messages` | POST | Create a draft email with a subject line |
| `/me/calendarView` | GET | List calendar events for a time range (today, this hour) |
| `/me/events/{id}` | GET | Get details of a specific calendar event |

## 8. Security Considerations

- **No client secret**: Public client app using PKCE — no secrets to manage or rotate.
- **No application permissions**: Only delegated permissions, so the app can never access another user's mailbox.
- **No external data transfer**: All API calls go to Microsoft Graph within the M365 tenant. No data is sent to third-party services.
- **Local execution only**: The app runs as a Python script on the user's Mac, triggered by Alfred keyboard shortcuts.
- **Revocable**: IT can delete the app registration at any time to revoke all access. The user can revoke consent from their Microsoft account.

## 9. Support Plan

This is a personal productivity tool maintained by the app owner. No end-user support, knowledge base articles, or escalation paths are required. If the app registration is revoked or permissions are changed, the app simply stops working with no impact on other users or systems.

## 10. Architecture Diagram

```
[User presses hotkey]
        |
        v
[Alfred Workflow on macOS]
        |
        v
[Python script extracts message ID from Outlook "Copy Link"]
        |
        v
[MSAL authenticates via cached token or browser login]
        |
        v
[HTTP request to https://graph.microsoft.com/v1.0/me/...]
        |
        v
[Microsoft Graph processes request within organizational M365 tenant]
        |
        v
[Email is archived / deleted / flagged / moved / draft created]
[OR: Calendar events are returned and displayed]
```
