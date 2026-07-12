# dvload — Feature inventory

## Existing features

### CLI (`dvload` command)

| Command | Description |
|---|---|
| `run` | Load rows from an `.xlsx` table into Dataverse via OData. Flags: `--dry-run`, `--refresh`, `--user`, `--max-errors`, `--concurrency`, `--resume`, `--notify-url`, `--no-failed-rows`. |
| `run-all` | Run several mappings in declared order from a manifest `.json` (for lookup dependencies between tables). Honors `stopOnError`. |
| `validate` | Check a `.dvmap.json` against local schema and live Dataverse metadata. `--no-remote` skips the network probe. |
| `login` / `logout` | Delegated (device-code) auth; caches a refresh token in Windows Credential Manager. |
| `app-login` / `app-logout` | App-only (client credentials) auth; secret stored in Windows Credential Manager. |
| `whoami` | Show which auth mode is configured for an environment and probe a live token. |
| `profile add/remove/list` | Named environment shortcuts stored in `~/.dvload/profiles.json`. |
| `extract-pqt` | Extract Power Query M code from an `.xlsx` into a `.pqt` archive, optionally injecting a column mapping into `MashupMetadata.json`. |
| `import-pqt` | Synthesise a `.dvmap.json` from the `MashupMetadata.json` inside an existing `.pqt`. |
| `schedule` | Register a Windows Scheduled Task for nightly unattended imports (`dvload run --refresh`). |

### Mapping engine (`@dvload/core`)

- **Conflict modes:** `insert`, `upsert`, `skip-if-exists`, `sync` (upsert + deactivate/delete target records missing from the source; `syncAction` picks the removal behavior, default deactivate; refuses to run on an empty source table)
- **Field types:** `string`, `memo`, `integer`, `decimal`, `money`, `double`, `boolean`, `datetime`, `dateonly`, `uniqueidentifier`, `lookup`, `choice`, `multichoice`, `status`, `state`
- **Lookup resolution:** by GUID, by alternate key, or by **text match** on any attribute (`lookupResolution: "text"` + `keyAttribute`), with `createIfMissing` (auto-create unmatched records) and `duplicateBehavior` (`error`/`first`) for ambiguous matches
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
- Column mapping UI (source column → target attribute + field kind)
- Auto-suggest column mappings based on column name similarity
- Import options UI: conflict mode (insert/upsert/skip-if-exists/sync), upsert key, sync action, batch size, parallel batches, bypass plugins/flows, skip unchanged rows
- Option-set labels fetched from metadata: picking a choice/multichoice/status/state target auto-fills `optionMap`, so spreadsheet cells can contain labels instead of integers
- Persists the last mapping per workbook in Office Settings store
- Saved environment profiles in `localStorage`
- Dev-mode banner — warns when using the fallback Microsoft PowerApps client ID

---

## Planned / not yet started

| Feature | Notes |
|---|---|
| **macOS / Linux scheduling** | `schedule` uses Windows Task Scheduler. Cross-platform support is v2 scope. |
| **Headless / server-side Power Query refresh** | Currently requires Excel desktop installed on the machine; server scenarios unsupported. |
| **Certificate auth for app-only** | Replace client secret (which must be rotated) with a certificate. `ConfidentialClientApplication` already supports `clientCertificate`. |
| **Single-file executable distribution** | Bundle CLI with `pkg`/`nexe`, code-sign the `.exe` so Windows SmartScreen passes. |
| **`winget` package** | Publish to the Windows Package Manager for easy install. |
| **CI build** | GitHub Actions workflow: `npm ci && npm run build` to keep the scaffold honest. |
| **Production add-in host** | Replace all `localhost:3000` URLs in `manifest.xml` with a real hosted URL. |
| **AppSource listing** | Finalise manifest (real GUID, icons, metadata), test on Win/Mac/Web/iPad, submit via Partner Center. |
