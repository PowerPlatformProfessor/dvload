# Unattended scheduled runs — setup guide

This guide is for **users of dvload** who want a daily import that runs
without anyone signed in. It requires a one-time setup in your Microsoft
Entra tenant and your Dataverse environment. Expect 20–30 minutes and
these permissions:

- Ability to create an app registration in Entra ID (or someone who can
  do it for you)
- System Administrator on the target Dataverse environment (to create an
  Application User)

## Why you need your own app registration

Interactive use (`dvload login`, the Excel add-in) works with no setup —
you just sign in. Unattended runs can't: a scheduled task has nobody
present to complete a sign-in prompt, and cached user tokens expire
(~90 days idle, password changes, conditional access changes). The fix
is **app-only authentication**: your organization registers its own app
with its own secret or certificate, and dvload authenticates as that app.
Credentials are per-organization by design — they cannot ship with the
tool.

## Prerequisites

- Windows 10/11 with Excel desktop installed (the schedule uses Windows
  Task Scheduler; `--refresh` drives Excel via COM)
- The machine must be on (not asleep) at the scheduled time
- dvload installed and a working mapping (`.dvmap.json`) — build one in
  the add-in or see the main README

## Step 1 — Register an app in Microsoft Entra

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com)
   → **Applications → App registrations → New registration**.
2. Name: `dvload (app-only)` — or anything you'll recognize later.
3. Supported account types: **Accounts in this organizational directory
   only** (single tenant).
4. No redirect URI needed. Register.
5. From the **Overview** page, copy the **Application (client) ID** and
   **Directory (tenant) ID** — you'll need both in Step 3.
6. Create a credential, either:
   - **Certificates & secrets → New client secret.** Copy the *value*
     immediately — it's shown only once. Note the expiry date; the
     import will stop working when it lapses, so put a rotation
     reminder in your calendar. Or, preferred:
   - **A certificate.** No secret-expiry surprises. You'll pass the PEM
     (certificate + private key in one file) to dvload in Step 3.

API permissions are not required for this flow — Dataverse checks the
Application User's security roles (Step 2), not the app's permission
grants.

## Step 2 — Create an Application User in Dataverse

The app registration needs an identity inside Dataverse to act as.

1. Open the [Power Platform admin center](https://admin.powerplatform.microsoft.com)
   → **Environments → your environment → Settings → Users + permissions
   → Application users**.
2. **New app user** → pick the app registration from Step 1.
3. Assign a **business unit**.
4. Assign a **security role** with create/update privileges on the
   tables you'll import into. A custom minimal role is best; *System
   Customizer* works if you don't want to build one.

This Application User is the principal that will appear as
`createdby` / `modifiedby` on every imported record.

## Step 3 — Store the credentials

One-time, on the machine that will run the schedule:

```bash
# with a client secret (you'll be prompted for it):
dvload app-login \
  --env https://yourorg.crm.dynamics.com \
  --client-id <application-client-id> \
  --tenant-id <directory-tenant-id>

# or with a certificate:
dvload app-login \
  --env https://yourorg.crm.dynamics.com \
  --client-id <application-client-id> \
  --tenant-id <directory-tenant-id> \
  --cert ./dvload-app.pem
```

The secret is stored DPAPI-protected in `~/.dvload/secrets.dat` — it is
readable only by your Windows account on this machine, and is never
written to the scheduled task definition.

`app-login` immediately tests the credentials by acquiring a token, so a
wrong tenant ID, missing Application User, or expired secret fails here
rather than at 3:30 AM. Verify any time with:

```bash
dvload whoami --env https://yourorg.crm.dynamics.com
```

It should report app-only auth for the environment, with no warnings.

## Step 4 — Schedule the import

```bash
dvload schedule ./contacts.dvmap.json \
  -w "C:\path\to\customers.xlsx" \
  --time 03:30 \
  --name "Daily Dataverse contacts import"
```

This registers a Windows Scheduled Task that runs
`dvload run … --refresh` daily: it refreshes the workbook's Power
Queries headlessly, reads the table, authenticates silently as the
Application User, and loads the rows.

If `schedule` warns that the environment is using delegated (signed-in
user) auth, go back to Step 3 — a schedule on cached user tokens will
break silently within weeks.

## Checking on it

- Failed rows land in `logs/failed_<workbook>_<timestamp>.xlsx` next to
  the workbook, in source-column shape — fix the cells and re-run just
  that file. Suppress with `--no-failed-rows` if the data is sensitive.
- Add `--notify-url <webhook>` to the run for a Teams/Slack summary
  after each run.
- Interrupted runs resume from a checkpoint with `--resume`.

## Rotating or removing credentials

```bash
# rotate: run app-login again with the new secret/cert — it overwrites
dvload app-login --env <url> --client-id <id> --tenant-id <id>

# remove:
dvload app-logout --env <url>
```

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `app-login` fails acquiring a token | Wrong tenant ID, or the secret was pasted with whitespace / already expired |
| Token acquired but runs get 403 | Application User missing, in the wrong environment, or its security role lacks privileges on the target tables |
| Runs worked, then stopped with auth errors | Client secret expired — create a new one and re-run `app-login` |
| Task didn't fire | Machine asleep or off at the scheduled time; check Task Scheduler history |
| `--refresh` hangs or fails | Excel desktop not installed, or the workbook's data source needs credentials Excel can't supply headlessly |
