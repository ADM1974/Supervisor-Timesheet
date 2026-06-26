# Timesheet Approvals — Supervisor app (Microsoft 365 sign-in)

The supervisor's app: sign in with a work 365 account, see the **submitted**
timesheets for the sites you supervise, and **Approve** or **Reject** them. It
reads and updates the **same Timesheets list** as the contractor + staff apps.

```
Supervisor phone → "Sign in with Microsoft" (MSAL) → ID token
   → Netlify Function verifies the token (signature, tenant, audience)
   → works out the supervisor's sites (Sites.ManagerEmail = their email)
   → lists Submitted rows for those sites, grouped per person/week
   → Approve/Reject updates the rows (app-only Graph connection)
```

## How a supervisor is matched to their sites
**Just use the `Manager` People Picker** already on the Sites list — pick each
site's supervisor from your 365 directory. The app resolves that person's email
automatically and shows them only their sites. Nothing to type.

Optional fallback: if you ever add a **`ManagerEmail`** (Single line of text)
column and fill it in, that's matched too — handy if a picked person can't be
resolved. You don't need it unless the People Picker matching misses someone.

> The People-Picker email resolution can only be fully confirmed once the app is
> running against your live SharePoint — verify it on first deploy (sign in as a
> supervisor and check you see your sites). If anyone's sites don't show, add the
> `ManagerEmail` column with their email as the fallback.

## One-time setup

### 1. Create the "Supervisor Timesheet" app registration (admin)
1. entra.microsoft.com → **App registrations → + New registration**.
2. Name `Supervisor Timesheet`; account type **"this organizational directory only"**.
3. **Redirect URI:** platform **Single-page application (SPA)** → your Netlify site
   URL (add after the first deploy, step 2).
4. Register → copy the **Application (client) ID** → this is `SUP_CLIENT_ID`. No secret.

### 2. Deploy (Git only — the `jose` dependency needs the build)
1. Put this folder in a **new GitHub repo** (e.g. `supervisor-timesheet`).
2. Netlify → **Add new site → Import an existing project → GitHub** → pick the repo.
3. Build command: empty. Publish directory: `.`  → Deploy. Note the URL.

### 3. Wire sign-in
- Put the site URL into the registration's **SPA redirect URI**.
- Edit **`config.js`** → paste the **client id** into `clientId` → commit (auto-deploys).

### 4. Environment variables (Netlify → this site)
Five are the **same values as your other sites** (copy them), plus the new one:
```
TENANT_ID        3efd78a4-4c46-434e-b653-4d0b65d18caa
CLIENT_ID        (the Contractor Timesheet app-only client id — reused for read/write)
CLIENT_SECRET    (its secret Value)
SP_SITE_ID       jdmclennan.sharepoint.com,7925a1f9-...,7d6a8277-...
LIST_ID          013ec528-5efd-4b78-a86b-6b8c148c2ff5   (Timesheets)
SUP_CLIENT_ID    (the NEW Supervisor Timesheet client id)
```
Redeploy after setting them.

### 5. Test
- Make sure a `Sites` row has your email in **ManagerEmail**, and there's at least
  one **Submitted** timesheet on that site.
- Open the site → **Sign in with Microsoft** → you should see that timesheet as a
  card → **Approve** → confirm the Timesheets row flips to `Approved` with an
  `ApprovedDate`. Reject asks for a reason and sets `RejectionReason`.

## Notes
- **Identity is verified server-side** and approvals are re-checked against the
  supervisor's sites, so nobody can approve outside their patch.
- Writes use the **app-only** connection, so supervisors don't each need SharePoint
  permissions.
- This replaces the old per-item Power Automate approval for these entries — decide
  per your rollout whether to keep both or retire the flow.
