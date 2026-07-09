# dvload — Feature inventory

## Existing features

### CLI (`dvload` command)

| Command | Description |
|---|---|
| `run` | Load rows from an `.xlsx` table into Dataverse via OData. Flags: `--dry-run`, `--refresh`, `--user`, `--max-errors`. |
| `validate` | Check a `.dvmap.json` against local schema and live Dataverse metadata. `--no-remote` skips the network probe. |
| `login` / `logout` | Delegated (device-code) auth; caches a refresh token in Windows Credential Manager. |
| `app-login` / `app-logout` | App-only (client credentials) auth; secret stored in Windows Credential Manager. |
| `whoami` | Show which auth mode is configured for an environment and probe a live token. |
| `profile add/remove/list` | Named environment shortcuts stored in `~/.dvload/profiles.json`. |
| `extract-pqt` | Extract Power Query M code from an `.xlsx` into a `.pqt` archive, optionally injecting a column mapping into `MashupMetadata.json`. |
| `import-pqt` | Synthesise a `.dvmap.json` from the `MashupMetadata.json` inside an existing `.pqt`. |
| `schedule` | Register a Windows Scheduled Task for nightly unattended imports (`dvload run --refresh`). |

### Mapping engine (`@dvload/core`)

- **Conflict modes:** `insert`, `upsert`, `skip-if-exists`
- **Field types:** `string`, `memo`, `integer`, `decimal`, `money`, `double`, `boolean`, `datetime`, `dateonly`, `uniqueidentifier`, `lookup`, `choice`, `multichoice`, `status`, `state`
- **Lookup resolution:** by GUID or by alternate key (`keyAttribute`)
- **Polymorphic lookup target selection:** for attributes that can point to multiple entity types (e.g. `regardingobjectid`), the add-in fetches the valid target entity sets from Dataverse metadata and lets the user pick one; the chosen `bindEntitySet` is saved in the mapping
- **`optionMap`** for `choice`/`multichoice` columns (human label → integer option value)
- **Batched OData** `$batch` POST/PATCH with configurable `batchSize` (Dataverse cap: 1000)
- **`maxErrors` guard** — stops a run early after N failures
- **Progress event stream** — `start`, `batch`, `row-success`, `row-error`, `done` with created / updated / skipped / failed counts
- **Run log** written as `.jsonl` alongside the workbook

### Excel add-in (task pane)

- MSAL popup sign-in (delegated flow)
- Lists Excel tables in the open workbook
- Fetches Dataverse entities and their attributes from the live environment
- Column mapping UI (source column → target attribute + field kind)
- Auto-suggest column mappings based on column name similarity
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
