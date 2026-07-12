# dvload

Map an Excel table (typically a Power Query output) to a Microsoft Dataverse
entity, then load via the OData Web API. Comes in two front ends that share
one engine:

- **Excel add-in** — task pane UI for picking a table, picking a Dataverse
  entity, building a column mapping, and running an import on demand.
- **CLI** — runs a saved `.dvmap.json` against an `.xlsx` file from the
  command line. Pairs with Windows Task Scheduler for daily unattended
  imports.

The mapping is portable: build it in the add-in, save it next to your
workbook, and have the CLI run it nightly.

## Repo layout

```
dvload/
├── packages/
│   ├── core/      # mapping engine, OData client, xlsx reader (shared)
│   ├── cli/       # Node CLI: dvload run|login|validate|schedule
│   └── addin/     # Office.js task pane add-in
├── package.json   # npm workspaces
└── tsconfig.base.json
```

## Prerequisites

- Node.js 20+
- Windows 10/11 (CLI scheduling and Power Query refresh are Windows-only in v1)
- Excel desktop, signed in to the same Microsoft 365 tenant as your Dataverse environment
- An Azure AD app registration (see below)

## Auth model

The CLI supports two flows; the add-in supports the first.

| | Delegated (user signs in) | App-only (client credentials) |
|---|---|---|
| Used by | add-in always; CLI by default | CLI when configured |
| Flow | device-code (CLI) or popup (add-in) | client_id + secret |
| Identity in Dataverse | the signed-in user | an "Application User" you create |
| Refresh expiry | ~90 days idle | none |
| Survives MFA / conditional access changes | sometimes | yes |
| Best for | interactive use | scheduled / unattended runs |

You can have both configured for the same environment. The CLI prefers
app-only when it's set; pass `--user` to force delegated (handy when
you're iterating in a shell where a schedule has stored a secret).

### Quick start (dev mode, no Entra app registration)

For local testing you can skip the Entra ID app registration entirely.
The CLI and add-in default to Microsoft's **PowerApps** public client id
(`2ad88395-b77d-4561-9441-d0e40824f9bc`), which is multi-tenant and has
Dataverse access pre-consented. The consent screen will say "Microsoft
PowerApps" the first time you sign in — that's expected in dev mode.

```bash
dvload login --env https://yourorg.crm.dynamics.com
```

When dev mode is active you'll see warnings on every auth command and a
yellow banner in the add-in. Both are deliberate — see
[PRE-RELEASE-CHECKLIST.md](./PRE-RELEASE-CHECKLIST.md) for everything
that has to change before you share this with anyone outside your own
machine.

### Delegated setup (replaces dev mode; required before shipping)

You need a **public client** app registration.

1. Open **Microsoft Entra admin center** → Applications → App registrations → New registration.
2. Name: `dvload`.
3. Supported account types: *Accounts in any organizational directory (multi-tenant)*.
4. Redirect URIs:
   - Single-page application: `https://localhost:3000/taskpane.html` (dev) and your prod URL.
   - Public client / native: leave the default `http://localhost` so device-code flow works.
5. Authentication → Allow public client flows: **Yes**.
6. API permissions → Add → Dynamics CRM → user_impersonation (delegated).
7. Copy the **Application (client) ID**.

Pin the client id:

```bash
setx DATAVERSE_LOAD_CLIENT_ID "<your-public-client-id>"
```

Or edit `DEFAULT_PUBLIC_CLIENT_ID` in `packages/cli/src/auth.ts` and
`CLIENT_ID` in `packages/addin/src/auth.ts`.

Then sign in:

```bash
dvload login --env https://contoso.crm.dynamics.com
```

### App-only setup (recommended for scheduled runs)

You need a separate **confidential client** app registration *and* an
Application User inside Dataverse so it has somewhere to map to.

In Microsoft Entra:

1. Register another app: name `dvload (app-only)`. Single-tenant
   is fine here — this app exists in your tenant and is used by your
   schedule, not shared with the world.
2. Certificates & secrets → New client secret. **Copy the value
   immediately** — you can't see it again. Note the expiry; rotate it
   before then. (Use a certificate instead of a secret for production.)
3. API permissions → Dynamics CRM → user_impersonation (delegated). For
   client-credentials flow Dataverse actually checks the Application
   User's roles, not API permissions, so this is mostly for hygiene.
4. Copy the **Application (client) ID** and **Directory (tenant) ID**.

In Dataverse (Power Platform admin center → your environment → Settings
→ Users + permissions → Application users):

1. New app user → pick the registration you just created → assign a
   business unit and at least one security role (e.g. *System
   Customizer* if it's writing data, or a custom role with the minimal
   create/update privileges on your target tables).
2. Confirm. The app user shows up as a record in the `systemusers` table
   and is the principal that will appear on every record's *modifiedby*.

Then store the credentials locally:

```bash
dvload app-login \
  --env https://contoso.crm.dynamics.com \
  --client-id <confidential-client-id> \
  --tenant-id <directory-tenant-id>
# prompts for the secret; stored in Windows Credential Manager
```

The command probes the credentials by acquiring a token; if anything is
misconfigured (wrong tenant, missing app user, expired secret) you'll
find out immediately.

To check what's configured for an environment:

```bash
dvload whoami --env https://contoso.crm.dynamics.com
```

To rotate or remove:

```bash
dvload app-login --env <url> --client-id <id> --tenant-id <id>  # overwrites
dvload app-logout --env <url>                                   # clears
```

## First run

```bash
# from the repo root
npm install
npm run build
```

### CLI

```bash
# One-time interactive sign-in (delegated):
dvload login --env https://contoso.crm.dynamics.com

# Or set up app-only auth (recommended for schedules — see "App-only setup"):
dvload app-login --env https://contoso.crm.dynamics.com \
  --client-id <id> --tenant-id <id>

# Confirm what's configured.
dvload whoami --env https://contoso.crm.dynamics.com

# Sanity-check a mapping against Dataverse metadata.
dvload validate ./contacts.dvmap.json

# Dry run — coerce + plan, but don't write anything.
dvload run ./contacts.dvmap.json -w ./customers.xlsx --dry-run

# Real run with Power Query refresh first.
dvload run ./contacts.dvmap.json -w ./customers.xlsx --refresh

# Force the interactive (delegated) flow even if app-only is configured.
dvload run ./contacts.dvmap.json -w ./customers.xlsx --user

# Faster bulk loads: 4 parallel batches, skip plugin/flow execution.
dvload run ./contacts.dvmap.json -w ./customers.xlsx --concurrency 4

# Resume an interrupted run from its checkpoint (same workbook only).
dvload run ./contacts.dvmap.json -w ./customers.xlsx --resume

# Post a summary to a Teams/Slack incoming webhook when done.
dvload run ./contacts.dvmap.json -w ./customers.xlsx --notify-url https://hooks.example/...

# Run several mappings in dependency order (accounts before contacts).
dvload run-all ./nightly.manifest.json
```

Failed rows are written to `logs/failed_<workbook>_<timestamp>.xlsx` in the
same column shape as the source table — fix the cells and re-run just that
file. Suppress with `--no-failed-rows` if the data is sensitive.

### Add-in (development)

**Entra app setup (one-time):** Register a public-client, multi-tenant app in Entra ID.
Add `https://localhost:3000/taskpane.html` as a **Single-page application** redirect URI and grant
`Dynamics CRM → user_impersonation` (delegated). Set the client ID in
`packages/addin/webpack.config.js` (the `DefinePlugin` default) or pass it at build time
via the `DATAVERSE_LOAD_CLIENT_ID` env var.

```powershell
# Terminal 1 — webpack dev server (hot-reload, HTTPS on port 3000)
npm run dev:addin

# Terminal 2 — sideload the manifest into Excel desktop
npm --workspace=@dvload/addin run start
```

If you see "Failed to add loopback exemption", run this **once** in an admin PowerShell
and then re-run the start command:

```powershell
CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
```

To stop:

```powershell
npm --workspace=@dvload/addin run stop
```

## Daily scheduled import

```bash
node packages/cli/dist/index.js schedule ./contacts.dvmap.json \
  -w "C:\Users\you\Documents\customers.xlsx" \
  --time 03:30 \
  --name "Daily Dataverse contacts import"
```

This registers a Windows Scheduled Task that runs daily at 03:30. The task
shells out to `dvload run … --refresh`, which:

1. Opens Excel headlessly via PowerShell + COM
2. Calls `Workbook.RefreshAll()` and waits for queries to finish
3. Saves the workbook
4. Reads the named table
5. Authenticates silently using the cached refresh token
6. Loads rows into Dataverse via batched POST/PATCH

For unattended runs, **use app-only auth** (see "App-only setup" above).
Run `dvload app-login` once, then `dvload schedule …`.
The schedule command warns you if the environment is using delegated
auth — those tokens eventually expire (90 days idle, password change,
conditional access changes), and a scheduled task can't pop a device-code
prompt to recover. App-only credentials don't expire until you rotate
the secret.

## Power Query Template (.pqt) bridge

`.pqt` is the format Dataverse Dataflows and Power Query Online emit — a
plain ZIP containing `MashupDocument.pq` (the M code), `MashupMetadata.json`
(per-query Dataverse field mappings), `Metadata.json` (name/description),
and `[Content_Types].xml`. The CLI can produce and consume it, which gives
you a clean path between local imports and Microsoft-hosted Dataflows.

### Extract from Excel

```bash
# Just the M code, in .pqt form:
dvload extract-pqt ./customers.xlsx
# writes ./customers.pqt

# With your column mapping injected into MashupMetadata.json:
dvload extract-pqt ./customers.xlsx \
  --mapping ./contacts.dvmap.json \
  --name "Daily contacts import" \
  -o ./out/contacts.pqt
```

The injected variant imports into a Dataverse Dataflow with the
column-to-attribute mapping pre-populated — handy if you eventually want
to move from running the CLI locally to running it server-side as a
hosted Dataflow.

Excel stores Power Query in a binary `DataMashup` blob inside
`customXml/`, not as a native .pqt. The codec parses the QDEFF wrapper,
pulls out the inner ZIP containing `Formulas/Section1.m`, and repackages
it. The M round-trips perfectly; the rich Dataflow metadata
(GUIDs, connection overrides, environment ids) is Dataflows-specific and
won't be present in an Excel-sourced .pqt — that gets filled in only on
import into Dataverse, or via `--mapping`.

### Import from a Dataverse Dataflow

```bash
# Pull a .pqt that someone exported from Dataverse Dataflows or PQ Online
dvload import-pqt ./activities.pqt \
  --env https://contoso.crm.dynamics.com \
  --query activitypointer \
  -o ./activities.dvmap.json
# also write the M alongside the mapping:
dvload import-pqt ./activities.pqt --env ... --emit-m
```

`import-pqt` reads `MashupMetadata.json → QueriesMetadata → FieldsMetadata`
and synthesizes a working `.dvmap.json` from it. Each
`SourceColumnName → DestinationFieldType + target field` entry becomes a
column mapping. The migration path is: build it once in Power Query
Online with the proper field mappings, export to .pqt, and from there
either keep running it as a Dataflow or pull the mapping into the CLI
and run locally on your own schedule.

## Mapping JSON shape

See `packages/core/examples/contacts.dvmap.json` for a full example. Key
fields:

| field | what it does |
|---|---|
| `environmentUrl` | Your Dataverse base URL, e.g. `https://contoso.crm.dynamics.com`. |
| `targetEntitySet` | Plural set name, e.g. `contacts`, `accounts`. |
| `sourceTable` | Excel table name (the kind PQ writes to). |
| `conflictMode` | `insert`, `upsert`, or `skip-if-exists`. |
| `upsertKey` | Required for upsert. List of attributes that form a Dataverse alternate key. |
| `batchSize` | Rows per `$batch` changeset (Dataverse caps at 1000). |
| `columns[].kind` | `string`, `integer`, `boolean`, `datetime`, `dateonly`, `lookup`, `choice`, etc. |
| `columns[].bindEntitySet` | For `lookup`: the entity set to bind to. |
| `columns[].lookupResolution` | `guid`, `alternateKey`, or `text` (match any attribute by exact text). |
| `columns[].keyAttribute` | When `lookupResolution=alternateKey/text`: the attribute to match on. |
| `columns[].createIfMissing` | `text` lookups: create the record when no match exists. |
| `columns[].duplicateBehavior` | `text` lookups: `error` (default) or `first` when 2+ records match. |
| `columns[].optionMap` | For `choice`/`multichoice`: human label → integer option value. The add-in fills this automatically from metadata. |
| `syncAction` | `conflictMode=sync`: `deactivate` (default) or `delete` records missing from the source. |
| `concurrency` | Parallel `$batch` requests, 1-8. Default 1. |
| `skipUnchanged` | upsert/sync: pre-read targets and skip rows/attributes with no changes. |
| `bypassCustomLogic` | Send bypass headers so plugins and Power Automate flows don't fire. |
| `impersonateUserId` | systemuser GUID to impersonate (`MSCRMCallerID`). |
| `notifyUrl` | Webhook that receives a `{text}` summary after each run. |

## Known limits in v1

- **Windows-only** for scheduled runs (Excel COM dependency).
- **Power Query refresh** requires Excel to be installed on the machine
  running the schedule. Headless / server scenarios aren't supported.
- **Client secrets** are stored in Windows Credential Manager. For
  long-lived deployments, swap to certificate auth (Azure AD supports it;
  ConfidentialClientApplication accepts a `clientCertificate` instead of
  `clientSecret`).

## License

MIT (suggested — adjust to taste before publishing).
