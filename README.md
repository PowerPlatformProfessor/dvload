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
├── docs/          # architecture, data formats, JSON schemas, auth notes
├── packaging/     # distribution manifests (Scoop, winget notes)
├── tests/         # seeded test-data generators and fixtures
├── package.json   # npm workspaces
└── tsconfig.base.json
```

### Documentation map

| Document | What it covers |
|---|---|
| this file | user-facing reference: install, auth, commands, mapping fields |
| [FEATURES.md](./FEATURES.md) | full feature inventory and planned work |
| [docs/ARCHITECTURE.md](./docs/ARCHITECTURE.md) | module boundaries, the load pipeline, invariants, extension points |
| [docs/DATA-FORMATS.md](./docs/DATA-FORMATS.md) | coercion rules, `$batch` wire format, QDEFF layout, on-disk state, log formats |
| [docs/schema/](./docs/schema/) | JSON Schemas for `.dvmap.json` and `.dvplan.json` |
| [docs/AUTH-NOTES.md](./docs/AUTH-NOTES.md) | what's been verified about Entra/Conditional Access, and what hasn't |
| [docs/SCHEDULED-RUNS.md](./docs/SCHEDULED-RUNS.md) | end-user guide for unattended nightly imports |
| [TEST-PROTOCOL.md](./TEST-PROTOCOL.md) | 22-section manual/E2E protocol, up to 100k rows |
| [TELEMETRY.md](./TELEMETRY.md) | what is collected, and how to turn it off |
| [PRE-RELEASE-CHECKLIST.md](./PRE-RELEASE-CHECKLIST.md) | what must be done before shipping |

## Prerequisites

- Node.js 20+
- Windows 10/11 (CLI scheduling and Power Query refresh are Windows-only in v1)
- Excel desktop, signed in to the same Microsoft 365 tenant as your Dataverse environment

No Entra ID app registration is needed for interactive use — dvload
ships with a registered multi-tenant app. Unattended scheduled runs need
a per-org registration; see [docs/SCHEDULED-RUNS.md](./docs/SCHEDULED-RUNS.md).

## Auth model

The CLI supports two flows; the add-in supports the first.

| | Delegated (user signs in) | App-only (client credentials) |
|---|---|---|
| Used by | add-in always; CLI by default | CLI when configured |
| Flow | browser (CLI, default), device-code (CLI, headless), popup (add-in) | client_id + secret |
| Identity in Dataverse | the signed-in user | an "Application User" you create |
| Refresh expiry | ~90 days idle | none |
| Survives MFA / conditional access changes | sometimes | yes |
| Best for | interactive use | scheduled / unattended runs |

You can have both configured for the same environment. The CLI prefers
app-only when it's set; pass `--user` to force delegated (handy when
you're iterating in a shell where a schedule has stored a secret).

### Quick start (delegated)

No Entra setup needed, and no admin approval:

```bash
dvload login --env https://yourorg.crm.dynamics.com
```

The CLI signs in through the shared Microsoft Dataverse client
(`51f81489-12ee-4a9e-aaae-a2591f45987d`) — the same one XrmToolBox and the
XRM Tooling SDK use. It's a Microsoft-owned app whose Dataverse delegated
permission is already consented in every tenant, so you never hit the
"Need admin approval" wall.

This is not a privilege bypass. You still sign in as yourself, and
Dataverse still enforces your security roles. What you give up is
attribution: the consent screen and your tenant's sign-in logs will say
*Microsoft Dynamics CRM*, not dvload.

If that client is unavailable — some tenants block it via Conditional
Access or Power Platform's *allowed client apps* control — dvload
automatically retries with its own registered multi-tenant public client
(`dataverse-load`, `e6828b0f-9fde-43f8-85d0-602660d498bb`), which may then
need a one-time admin approval. Sign-ins that fail for reasons about *you*
rather than the app (Conditional Access, MFA, declined consent) are not
retried.

To skip the shared client and always sign in as dvload:

```bash
setx DVLOAD_NO_SHARED_CLIENT 1
```

`dvload whoami --env <url>` reports which client id the stored session
actually used.

#### Sign-in flow

By default `dvload login` opens your system browser (authorization code +
PKCE, with a loopback listener on `127.0.0.1`). It switches to device code
automatically when there's no local browser to open — over SSH, or on Linux
with no display server. Override with `--interactive` / `--device-code`, or
`DVLOAD_AUTH_FLOW=interactive|device-code`.

Because the shared Microsoft client is both pre-consented *and* accepts
`http://localhost` redirects, browser sign-in through it needs neither admin
consent nor device code flow. That's the combination that lets tools like
XrmToolBox connect in locked-down tenants.

If you see `AADSTS900971` ("No reply address provided") on the Entra page,
suspect the URL rather than the app registration — a mangled authorize
request arrives missing its `redirect_uri` and produces exactly this error.
dvload prints the sign-in URL before opening it; paste that into your browser
by hand. If the manual paste works, the launcher is at fault. See
[docs/AUTH-NOTES.md](./docs/AUTH-NOTES.md).

Note that Entra validates redirect URIs only *after* authentication and shows
the result as a browser page, so dvload can't observe these failures or retry
them — it times out after three minutes (`DVLOAD_AUTH_TIMEOUT_MS`).

**If you get `AADSTS53003` — "your sign-in was successful but does not meet
the criteria to access this resource"** — Conditional Access refused to
issue the token. The most common cause for a CLI is the **Authentication
Flows** condition blocking device code flow; Microsoft recommends getting
"as close as possible to a unilateral block" on it, because an attacker can
generate a code and phish someone into entering it. Retry with:

```bash
dvload login --env <url> --interactive
```

That isn't a way around the policy — it's the flow the policy steers you
to. A real browser can present device state (primary refresh token,
compliant-device claim) that device code structurally cannot.

If the blocking condition is device compliance or location instead, no
dvload setting will help. Check **Entra admin center → Monitoring →
Sign-in logs → the failed attempt → Conditional Access tab** to see which
policy and condition actually fired. dvload prints this guidance itself
when it detects a CA block.

> **Add-in note:** none of this applies to the Excel add-in, which always
> uses a real app registration. Browser flows need a `spa`-type redirect
> URI on the app — both for the redirect itself and for the token
> endpoint's CORS header — and you cannot add one to a Microsoft-owned
> app. Swapping MSAL.js for another library doesn't change that.

### Using your own app registration (optional)

Orgs that prefer their own registration — for conditional-access
policies, consent branding, or auditability — can override the client id.
Setting it also disables the shared-client default. Register a **public
client** app:

1. Open **Microsoft Entra admin center** → Applications → App registrations → New registration.
2. Supported account types: single-tenant is fine for internal use.
3. Redirect URIs:
   - Single-page application: the URL your copy of the add-in is hosted at (only if you self-host the add-in).
   - Public client / native: leave the default `http://localhost`.
4. Authentication → Allow public client flows: **Yes**.
5. API permissions → Add → Dynamics CRM → user_impersonation (delegated).
6. Copy the **Application (client) ID** and pin it:

```bash
setx DATAVERSE_LOAD_CLIENT_ID "<your-client-id>"
```

> **This disables the shared-client fallback.** Once pinned, dvload uses only
> your app, so it must have `http://localhost` registered under Authentication
> -> Mobile and desktop applications or browser sign-in fails with
> `AADSTS900971`. To go back to the pre-consented shared client, clear it:
> `setx DATAVERSE_LOAD_CLIENT_ID ""` (new shell required). `dvload login`
> prints which client id it's using and what pinned it, so check that first
> when sign-in misbehaves.

The CLI picks up the env var on the next run. For the add-in, the id is
baked in at build time (`DATAVERSE_LOAD_CLIENT_ID` at webpack build), so
overriding it means self-hosting a rebuild.

### App-only setup (recommended for scheduled runs)

> **End users:** there's a standalone, step-by-step version of this in
> [docs/SCHEDULED-RUNS.md](./docs/SCHEDULED-RUNS.md) — share that one.

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
# prompts for the secret; stored DPAPI-protected in ~/.dvload/secrets.dat

# or, better for scheduled runs — a certificate instead of a secret
# (no rotation-policy expiry; PEM must contain cert + private key):
dvload app-login \
  --env https://contoso.crm.dynamics.com \
  --client-id <confidential-client-id> \
  --tenant-id <directory-tenant-id> \
  --cert ./dvload-app.pem
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
dvload run-all ./nightly.dvplan.json
```

Failed rows are written to `logs/failed_<workbook>_<timestamp>.xlsx` in the
same column shape as the source table — fix the cells and re-run just that
file. Suppress with `--no-failed-rows` if the data is sensitive.

### Add-in (development)

Dev builds use the built-in `dataverse-load` app, which already has
`https://localhost:3000/taskpane.html` as an SPA redirect URI — no Entra
setup needed. To build against a different registration, pass
`DATAVERSE_LOAD_CLIENT_ID` at build time (its app needs that same SPA
redirect URI and `Dynamics CRM → user_impersonation` delegated permission).

```powershell
# Terminal 1 — webpack dev server (hot-reload, HTTPS on port 3000)
npm run dev:addin

# Terminal 2 — sideload the manifest into Excel desktop
npm --workspace=@dvload/addin run start

# Same from the CLI (local/dev repo clone)
dvload addin start
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

Full end-user setup guide (Entra app, Application User, credentials,
troubleshooting): [docs/SCHEDULED-RUNS.md](./docs/SCHEDULED-RUNS.md).

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

```bash
# One .dvmap.json per query (for multi-table Dataflows):
dvload import-pqt ./flow.pqt --env https://contoso.crm.dynamics.com --all-queries -o ./mappings/
```

### Load a Dataflow's queries into Excel (experimental)

```bash
dvload pqt-to-xlsx ./flow.pqt -o ./flow.xlsx
```

Builds a fresh workbook with the .pqt's M queries embedded natively in
Power Query (synthesized `DataMashup`/QDEFF part). Open it in Excel →
Data → Queries & Connections: every query is there as connection-only;
use *Load To…* to land one on a sheet. The QDEFF writer is experimental —
if Excel rejects the file, the fallback is `import-pqt --emit-m` and
pasting the M into a Blank Query's Advanced Editor.

### In the add-in

*Import .pqt…* in the task pane reads a Dataflow export directly: it
lists every query with its field-mapping count, *Use mapping* populates
the column grid from the selected query's `FieldsMetadata`, and *Copy M*
puts the M document on the clipboard for pasting into Excel's Advanced
Editor.

The task pane also supports run plans: add/edit/load/save `.dvplan.json`
steps, assign `stage`/`dependsOn`, and capture alternate-key links for
cross-step lookup wiring.

## Mapping JSON shape

Full example: `packages/core/examples/contacts.dvmap.json`. Machine-readable
schema: [docs/schema/dvmap.schema.json](./docs/schema/dvmap.schema.json) —
point your editor at it for completion and inline validation:

```json
{ "$schema": "../docs/schema/dvmap.schema.json", "schemaVersion": 1, "...": "..." }
```

The authoritative validation lives in `parseMapping` / `validateMapping`
(`packages/core/src/mapping.ts`); the schema mirrors it. Key fields:

| field | what it does |
|---|---|
| `environmentUrl` | Your Dataverse base URL, e.g. `https://contoso.crm.dynamics.com`. |
| `targetEntitySet` | Plural set name, e.g. `contacts`, `accounts`. |
| `sourceTable` | Excel table name (the kind PQ writes to). |
| `conflictMode` | `insert`, `upsert`, or `skip-if-exists`. |
| `upsertKey` | Required for upsert. Either the attributes forming a Dataverse alternate key, or a single `uniqueidentifier` column holding the record's own id (see [Upserting on the record id](#upserting-on-the-record-id)). |
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

### Upserting on the record id

`upsertKey` normally names a Dataverse **alternate key** — the usual integration
case, where the source system has no idea what a Dataverse GUID is. But if your
source already holds the record ids (a re-import, a migration between
environments, or ids you minted yourself), you can upsert on the primary key
instead: map it as `uniqueidentifier` and name it as the only `upsertKey`.

```json
{
  "conflictMode": "upsert",
  "upsertKey": ["accountid"],
  "columns": [
    { "source": "accountid", "target": "accountid", "kind": "uniqueidentifier", "treatEmptyAsNull": true },
    { "source": "Name", "target": "name", "kind": "string", "treatEmptyAsNull": true }
  ]
}
```

dvload addresses these rows as `accounts(<guid>)` rather than the
`accounts(accountid='<guid>')` alternate-key form, which is what Dataverse
expects for a primary key — the record is created with that id if it doesn't
exist and updated if it does. No alternate key needs to exist on the table.

Two things to know:

- The id column is dropped from the request body (it's already in the URL).
  Dataverse rejects writes to the primary key on an existing row, so leaving it
  in would fail every update while letting creates through.
- Values must be real GUIDs. A blank or malformed id fails that row rather than
  silently creating a record with a server-generated id.

Compound keys and non-GUID keys are unaffected and still use the alternate-key
form.

## Run plan JSON shape (`.dvplan.json`)

Use a run plan to orchestrate multiple `.dvmap.json` files while keeping each
single-table mapping unchanged. Schema:
[docs/schema/dvplan.schema.json](./docs/schema/dvplan.schema.json).

```json
{
  "schemaVersion": 1,
  "name": "nightly load",
  "stopOnError": true,
  "steps": [
    {
      "id": "accounts",
      "mapping": "./accounts.dvmap.json",
      "workbook": "./data.xlsx",
      "stage": 1
    },
    {
      "id": "contacts",
      "mapping": "./contacts.dvmap.json",
      "workbook": "./data.xlsx",
      "stage": 2,
      "dependsOn": ["accounts"],
      "alternateKeyLinks": [
        {
          "fromStep": "accounts",
          "lookupTarget": "parentcustomerid_account",
          "keyAttribute": "accountnumber"
        }
      ]
    }
  ]
}
```

- Steps in the same `stage` run in parallel.
- `dependsOn` forces explicit order.
- `alternateKeyLinks` enforces lookup-by-alternate-key wiring across steps:
  - upstream step must expose the key in `upsertKey`
  - downstream lookup must use `lookupResolution: "alternateKey"` and matching `keyAttribute`.
- Legacy `run-all` manifests (`runs[]`) still work; they're auto-upgraded in memory.

## Known limits in v1

- **Windows-only** for scheduled runs (Excel COM dependency).
- **Power Query refresh** requires Excel to be installed on the machine
  running the schedule. Headless / server scenarios aren't supported.
- **Secrets and tokens** live in `~/.dvload/` — DPAPI-protected
  (CurrentUser) on Windows, mode 0600 on POSIX. Nothing is stored in Windows
  Credential Manager. For long-lived deployments prefer certificate auth:
  `dvload app-login --cert <pem>`. Full layout:
  [docs/DATA-FORMATS.md](./docs/DATA-FORMATS.md#on-disk-state-dvload).

## Contributing

Read [docs/ARCHITECTURE.md](./docs/ARCHITECTURE.md) first — particularly the
[invariants](./docs/ARCHITECTURE.md#invariants) and
[extension points](./docs/ARCHITECTURE.md#extension-points) sections, which
list the steps a change usually needs to touch. Docs are treated as part of
the change, not a follow-up; see
[.github/copilot-instructions.md](./.github/copilot-instructions.md).

## License

MIT — see [LICENSE](./LICENSE).
