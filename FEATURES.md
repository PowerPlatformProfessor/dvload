# dvload — Feature inventory

## Existing features

### CLI (`dvload` command)

| Command | Description |
|---|---|
| `run` | Load rows from an `.xlsx` table (or `.csv`/`.tsv` file) into Dataverse via OData. Flags: `--dry-run`, `--refresh`, `--user`, `--max-errors`, `--concurrency`, `--resume`, `--notify-url`, `--no-failed-rows`, `--non-interactive`, `--json` (machine-readable result for pipelines). |
| `run-all` | Run several mappings from a run plan (`.dvplan.json`): stage-based parallel execution, dependency ordering, optional per-step overrides, and backward-compatible support for legacy `runs[]` manifests. Honors `stopOnError`. |
| `validate` | Check a `.dvmap.json` against local schema and live Dataverse metadata. `--no-remote` skips the network probe. |
| `login` / `logout` | Delegated auth via system browser (auth code + PKCE), falling back to device code on headless machines; refresh token cached DPAPI-encrypted in `~/.dvload/`. |
| `app-login` / `app-logout` | App-only auth via client secret or certificate (`--cert <pem>`); stored in the DPAPI-protected secure store. |
| `whoami` | Show which auth mode is configured for an environment and probe a live token. |
| `profile add/remove/list` | Named environment shortcuts stored in `~/.dvload/profiles.json`. |
| `extract-pqt` | Extract Power Query M code from an `.xlsx` into a `.pqt` archive, optionally injecting a column mapping into `MashupMetadata.json`. |
| `import-pqt` | Synthesise a `.dvmap.json` from the `MashupMetadata.json` inside an existing `.pqt`. `--all-queries` emits one mapping per query; `--emit-m` writes the M document. |
| `pqt-to-xlsx` | **Experimental:** build an `.xlsx` with the `.pqt`'s queries embedded natively in Power Query (QDEFF/DataMashup writer). Queries arrive connection-only; use "Load To…" in Excel. `--open` launches it in Excel when done (interactive runs ask; `--no-open` for scripts). Also in the pane as *Create workbook*, download-only. |
| `schedule` | Register a Windows Scheduled Task for nightly unattended imports (`dvload run --refresh`). |
| `addin` | Launch/stop add-in local workflow from CLI (`dvload addin start|stop|dev`). |

### Mapping engine (`@dvload/core`)

- **Conflict modes:** `insert`, `upsert`, `skip-if-exists`, `sync` (upsert + deactivate/delete target records missing from the source; `syncAction` picks the removal behavior, default deactivate; refuses to run on an empty source table)
- **Field types:** `string`, `memo`, `integer`, `decimal`, `money`, `double`, `boolean`, `datetime`, `dateonly`, `uniqueidentifier`, `lookup`, `choice`, `multichoice`, `status`, `state`
- **Lookup resolution:** by GUID, by alternate key, or by **text match** on any attribute (`lookupResolution: "text"` + `keyAttribute`), with `createIfMissing` (auto-create unmatched records) and `duplicateBehavior` (`error`/`first`) for ambiguous matches
- **Constant-value columns** (`constant` instead of `source`): apply a fixed value to every record without a source column — e.g. a fixed owner (`ownerid` lookup bound to a user/team GUID), a hardcoded choice value, or a marker string. Coerced through the same pipeline as cell values; works in the CLI and scheduled runs since it's part of the `.dvmap.json` schema
- **Delta detection** (`skipUnchanged`): pre-reads target records, strips attributes whose values already match, and skips no-op rows — keeps the audit log and plugins quiet
- **Parallel batches** (`concurrency: 1-8`): multiple `$batch` requests in flight, still honoring Retry-After throttling
- **Bypass custom logic** (`bypassCustomLogic`): sends `MSCRM.BypassCustomPluginExecution` + `MSCRM.SuppressCallbackRegistrationExpanderJob` per operation (needs the bypass privilege)
- **Impersonation** (`impersonateUserId`): `MSCRMCallerID` header so records are created as a specific user
- **Checkpoint/resume**: `checkpoint` progress events + `startOffset`; the CLI persists checkpoints and `--resume` continues an interrupted run if the workbook is unchanged
- **Cooperative cancellation** (`signal: AbortSignal`): stops scheduling new batches; in-flight batches complete, unattempted rows count as skipped, result flagged `cancelled`; sync-mode removal pass never runs on a cancelled run
- **Webhook notification** (`notifyUrl` or `--notify-url`): POSTs a `{text}` summary (Teams/Slack incoming-webhook compatible) after each run
- **Polymorphic lookup target selection:** for attributes that can point to multiple entity types (e.g. `regardingobjectid`), the add-in fetches the valid target entity sets from Dataverse metadata and lets the user pick one; the chosen `bindEntitySet` is saved in the mapping
- **`optionMap`** for `choice`/`multichoice` columns (human label → integer option value)
- **Batched OData** `$batch` POST/PATCH with configurable `batchSize` (Dataverse cap: 1000)
- **`maxErrors` guard** — stops a run early after N failures
- **Progress event stream** — `start`, `batch`, `row-success`, `row-error`, `sync`, `checkpoint`, `done` with created / updated / skipped / unchanged / removed / failed counts
- **Run log** written as `.jsonl` alongside the workbook
- **Failed-rows re-run file** — failures are written to `logs/failed_<workbook>_<timestamp>.xlsx` in the same column shape as the source table; fix the cells and re-run just that file (suppress with `--no-failed-rows`)

### Resilience

- **Throttle-aware retry** on 429/503/504, honouring `Retry-After` (seconds or HTTP-date) with exponential backoff and a 60s cap
- **Dropped-connection retry**: a request that throws instead of responding (machine sleeps mid-run, VPN reconnects, link blips) is retried with the same backoff rather than failing an entire `$batch` of rows. Failures are classified — a connection that was never established (`ENOTFOUND`, `ECONNREFUSED`, connect timeout) is always safe to replay; one that died mid-flight is only replayed when the operations are idempotent, so plain POST creates can't be duplicated; cancellation is never replayed
- `--max-attempts <n>` on `run` / `run-all` (default 5) to ride out longer outages; retries are counted in the run summary and the `--json` output, and each dropped request is recorded in the run log with status 0 and its error code
- Runs that fail with no HTTP response are called out in the summary as a connectivity problem rather than bad data

### Web UI (Excel task pane, `dvload gui` in a browser, or a Power Platform ToolBox tool)

One bundle. In Excel and the browser it is served from loopback by
`dvload serve`; in the ToolBox it ships as its own tool package
(`packages/pptb`, built from the same sources). The host differences are
confined to `packages/addin/src/host.ts` (plus `src/pptb/` for the ToolBox).

- Sign-in delegated to `dvload serve`, so no Entra app registration and no admin consent — the UI carries no client id and never talks to Entra
- **Three tabs**: *Import* (account, environment, source, target, column mapping, import options, Run import), *Run plan* (multi-step `.dvplan.json` editor, with a step count on the tab), *Dataflows* (read a dataflow out of the environment, plus the `.pqt` import/export tools). Status, progress and the run log sit below the tabs, so a running import keeps reporting whichever tab you're on
- **Account first, environment second**: pick who you are, then where to write. The username is remembered across environment switches, and selecting a saved profile you haven't used yet signs in with that username pre-filled (`login_hint`) rather than starting from an account picker. Profiles you already have a session for are ticked
- `dvload gui` runs the same interface without Excel; everything works except reading the open workbook and extracting its Power Query
- Lists Excel tables in the open workbook, or reads a picked `.xlsx`/`.csv`/`.tsv` file instead
- Fetches Dataverse entities and their attributes from the live environment
- Solution picker: defaults to all entities (Default solution); selecting a solution filters the target-entity list to that solution's tables (via `solutioncomponents`, componenttype 1)
- **Create new table from source** ("＋ Create new table from source…" in the entity picker): infers column types from sample rows (string/memo/integer/decimal/boolean/datetime/dateonly), lets you rename columns, pick the primary name column and an alternate key, then one button creates the table + columns + key via the metadata API and immediately runs the import against it ("Create table and run import")
- Column mapping UI (source column → target attribute + field kind)
- Type-ahead comboboxes on the solution, entity, target-attribute, and lookup entity-set pickers (type first letters to filter)
- Fixed-value rows ("Add fixed value"): set a field to a constant on every record; lookup targets (e.g. owner) get a live record search against the bound entity set (users/teams), storing the picked record's GUID as a chip you can clear. Alternate-key and text resolutions take a plain key value instead, and the field is disabled until the bound entity set is chosen
- Guided lookup config sub-row: labelled **Binds to → Match by → Key field**, in fill-in order. "Match by" offers GUID / alternate key / text match; the key field is a picker populated from the bound entity's metadata (only single-attribute alternate keys for alt-key mode, text attributes for text mode) rather than a typed logical name. Text mode also exposes `createIfMissing` and `duplicateBehavior`
- Inline per-column validation (`validateColumn`) under each mapping row, so a half-configured lookup shows its error while you build it; "Save mapping…" warns before writing a mapping `dvload run` would reject
- Auto-suggest column mappings based on column name similarity
- Import options UI: conflict mode (insert/upsert/skip-if-exists/sync), upsert key, sync action, batch size, parallel batches, bypass plugins/flows, skip unchanged rows, **dry run**
- **Advanced options**: mapping name and description, `maxErrors` (stop after N row errors), `notifyUrl` (Teams/Slack webhook posted after the run), and **Run as user** — a systemuser search that sets `impersonateUserId` / `MSCRMCallerID`
- **Date format override** on datetime/dateonly columns, so ambiguous text dates (`03/04/2025`) can be pinned to `dd/MM/yyyy` rather than guessed
- Option-set labels fetched from metadata: picking a choice/multichoice/status/state target auto-fills `optionMap`, so spreadsheet cells can contain labels instead of integers
- **File source** ("Use a file instead…"): read rows from another `.xlsx`/`.xlsm` or a `.csv`/`.tsv` rather than a table in the open workbook. A picked workbook without a named table falls back to the first sheet's used range
- .pqt round-trip: read a Dataverse Dataflow / PQ Online export, list its queries with field-mapping counts, populate the grid from one query or **download a mapping for every query** (adding a run-plan step each), copy the M code, and **export a .pqt** with the current mapping written into its `FieldsMetadata`
- Run-plan editor: load/save/edit `.dvplan.json` including plan name, description and `stopOnError`; per-step `stage` / `dependsOn` / `overrides` (`maxErrors=N`, `concurrency=N`, `notifyUrl=…`, `dryRun`, `user`, `noFailedRows`); alternate-key links (`fromStep` + lookup target + key attribute) for cross-step dependencies; and inline `validateRunPlan` errors before save
- Cancel button during runs (confirm dialog; in-flight batches finish, summary shows how far it got) and a close-pane warning while an import is running
- **Resume after cancel**: a cancelled run leaves a checkpoint in the workbook's settings, and the pane offers to restart at the row it stopped on (guarded against the source changing underneath)
- **Failed-rows download**: the rows that errored, as an `.xlsx` you can fix and re-import — the pane equivalent of the CLI's `failed_*.xlsx`
- Persists the last mapping per workbook in Office Settings store; fields the pane has no control for (`createdAt`, `logDir`) survive a load → save round-trip instead of being silently dropped
- Saved environment profiles in `localStorage`
- Dev-mode banner — warns when using the fallback Microsoft PowerApps client ID

### Power Platform ToolBox tool (`packages/pptb`)

The same pane packaged as a [PPTB](https://www.powerplatformtoolbox.com/)
tool. PPTB never gives tools an access token, so the engine runs on a
`DataverseGateway` adapter over the ToolBox's `window.dataverseAPI` bridge
(`packages/addin/src/pptb/pptb-client.ts`) instead of the direct OData client.

- Uses the ToolBox's **active connection** — the account/environment steps
  disappear, and a connection switch reloads the tool
- `$batch` and PATCH-upserts are **emulated** through per-record bridge calls
  (per-row accounting semantics preserved; slower, and upserts are
  probe-then-write rather than atomic)
- Dataflow import runs fully in-page (core's conversion against the bridge)
- Saves (`.dvmap.json`, failed rows, run logs, workbooks) go through the
  ToolBox's native save dialog — a sandboxed iframe can't trigger downloads
- Not available there: **bypass plugins**, **run as user** (both are
  per-request headers the bridge can't send) and **alternate-key creation**
  in create-table mode; sync mode's removal listing depends on the bridge
  paging past 5,000 rows (truncation removes fewer, never more)
- Build/debug/publish flow in [packages/pptb/README.md](./packages/pptb/README.md)

---

## Planned / not yet started

| Feature | Notes |
|---|---|
| **macOS / Linux scheduling** | `schedule` uses Windows Task Scheduler. Cross-platform support is v2 scope. |
| **Schedule cadences beyond daily** | `schedule` only supports a fixed daily `--time`. Add `--every` (weekly/weekdays/monthly/hourly, specific days via `schtasks /sc weekly /d MON,WED`, N-hour intervals via `/sc hourly /mo N`). |
| **Headless / server-side Power Query refresh** | Currently requires Excel desktop installed on the machine; server scenarios unsupported. |
| ~~**Certificate auth for app-only**~~ | Done: `dvload app-login --cert <pem>` stores the cert in the secure store. |
| ~~**Single-file executable distribution**~~ | Scaffolded: `npm run bundle` (esbuild) + Node SEA in `.github/workflows/release.yml`. Code signing still to wire up (see `packaging/README.md`). |
| ~~**`winget` package**~~ | Documented in `packaging/README.md`; submit after the first signed release. Scoop manifest template in `packaging/scoop/`. |
| ~~**CI build**~~ | Done: `.github/workflows/ci.yml` (build + test + lint on windows-latest). |
| ~~**Production add-in host**~~ | Dropped, not done — there is nothing to host. `dvload serve` serves the pane from loopback, so dev and production are the same setup and the manifest URL never changes. |
| ~~**Entra app registration for the add-in**~~ | No longer required. The UI gets tokens from `dvload serve`, which uses the CLI's pre-consented client. No admin consent, in either front end. |
| **Auto-start `dvload serve`** | The pane is dead if the sidecar isn't running; today that's a "start dvload" screen. Register a logon-triggered Scheduled Task (the `schedule` command already has the Task Scheduler plumbing) or a tray shim. |
| **AppSource listing** | Poor fit as designed: a listed add-in must load from a public host, and this one loads from localhost and needs the CLI running. Would mean reintroducing a hosted pane, its own app registration, and admin consent. Centralized Deployment or sideloading are the routes that fit. |
| **Public telemetry dashboard (Power BI)** | À la FetchXML Builder ([jonasr.app/xtb-stats](https://jonasr.app/xtb-stats/)): Power BI report over the Application Insights events, published publicly and linked from TELEMETRY.md. Doubles as transparency (users see exactly what granularity exists) and marketing. Requires: create the App Insights resource, bake the connection string into CLI + add-in builds, build the report, publish-to-web. |
| **`dvload init` wizard** | Guided first-run: env URL → profile → auth choice → login. Collapses the Entra-registration onboarding cliff. |
