# dvload — Feature inventory

## Existing features

### CLI (`dvload` command)

| Command | Description |
|---|---|
| `run` | Load rows from an `.xlsx` table (or `.csv`/`.tsv` file) into Dataverse via OData. Flags: `--dry-run`, `--refresh`, `--user`, `--max-errors`, `--concurrency`, `--resume`, `--notify-url`, `--no-failed-rows`, `--non-interactive`, `--json` (machine-readable result for pipelines). |
| `run-all` | Run several mappings in declared order from a manifest `.json` (for lookup dependencies between tables). Honors `stopOnError`. |
| `validate` | Check a `.dvmap.json` against local schema and live Dataverse metadata. `--no-remote` skips the network probe. |
| `login` / `logout` | Delegated (device-code) auth; refresh token cached DPAPI-encrypted in `~/.dvload/`. |
| `app-login` / `app-logout` | App-only auth via client secret or certificate (`--cert <pem>`); stored in the DPAPI-protected secure store. |
| `whoami` | Show which auth mode is configured for an environment and probe a live token. |
| `profile add/remove/list` | Named environment shortcuts stored in `~/.dvload/profiles.json`. |
| `extract-pqt` | Extract Power Query M code from an `.xlsx` into a `.pqt` archive, optionally injecting a column mapping into `MashupMetadata.json`. |
| `import-pqt` | Synthesise a `.dvmap.json` from the `MashupMetadata.json` inside an existing `.pqt`. `--all-queries` emits one mapping per query; `--emit-m` writes the M document. |
| `pqt-to-xlsx` | **Experimental:** build an `.xlsx` with the `.pqt`'s queries embedded natively in Power Query (QDEFF/DataMashup writer). Queries arrive connection-only; use "Load To…" in Excel. |
| `schedule` | Register a Windows Scheduled Task for nightly unattended imports (`dvload run --refresh`). |

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
- **Webhook notification** (`notifyUrl` or `--notify-url`): POSTs a `{text}` summary (Teams/Slack incoming-webhook compatible) after each run
- **Polymorphic lookup target selection:** for attributes that can point to multiple entity types (e.g. `regardingobjectid`), the add-in fetches the valid target entity sets from Dataverse metadata and lets the user pick one; the chosen `bindEntitySet` is saved in the mapping
- **`optionMap`** for `choice`/`multichoice` columns (human label → integer option value)
- **Batched OData** `$batch` POST/PATCH with configurable `batchSize` (Dataverse cap: 1000)
- **`maxErrors` guard** — stops a run early after N failures
- **Progress event stream** — `start`, `batch`, `row-success`, `row-error`, `sync`, `checkpoint`, `done` with created / updated / skipped / unchanged / removed / failed counts
- **Run log** written as `.jsonl` alongside the workbook
- **Failed-rows re-run file** — failures are written to `logs/failed_<workbook>_<timestamp>.xlsx` in the same column shape as the source table; fix the cells and re-run just that file (suppress with `--no-failed-rows`)

### Excel add-in (task pane)

- MSAL popup sign-in (delegated flow)
- Lists Excel tables in the open workbook
- Fetches Dataverse entities and their attributes from the live environment
- Solution picker: defaults to all entities (Default solution); selecting a solution filters the target-entity list to that solution's tables (via `solutioncomponents`, componenttype 1)
- Column mapping UI (source column → target attribute + field kind)
- Type-ahead comboboxes on the solution, entity, target-attribute, and lookup entity-set pickers (type first letters to filter)
- Fixed-value rows ("Add fixed value"): set a field to a constant on every record; lookup targets (e.g. owner) get a live record search against the bound entity set (users/teams), storing the picked record's GUID
- Auto-suggest column mappings based on column name similarity
- Import options UI: conflict mode (insert/upsert/skip-if-exists/sync), upsert key, sync action, batch size, parallel batches, bypass plugins/flows, skip unchanged rows
- Option-set labels fetched from metadata: picking a choice/multichoice/status/state target auto-fills `optionMap`, so spreadsheet cells can contain labels instead of integers
- Import .pqt: read a Dataverse Dataflow / PQ Online export in the task pane, list its queries with field-mapping counts, populate the mapping grid from any query, and copy the M code for pasting into Excel's Advanced Editor
- Persists the last mapping per workbook in Office Settings store
- Saved environment profiles in `localStorage`
- Dev-mode banner — warns when using the fallback Microsoft PowerApps client ID

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
| **Production add-in host** | Replace all `localhost:3000` URLs in `manifest.xml` with a real hosted URL. Production webpack builds now refuse the borrowed dev client id. |
| **AppSource listing** | Finalise manifest (real GUID, icons, metadata), test on Win/Mac/Web/iPad, submit via Partner Center. Centralized Deployment (M365 admin center) is the better route for known orgs. |
| **`dvload init` wizard** | Guided first-run: env URL → profile → auth choice → login. Collapses the Entra-registration onboarding cliff. |
