# dvload — Copilot instructions

## What this repo is

A monorepo (npm workspaces) with three packages:
- **`packages/core`** — shared mapping engine, OData client, xlsx reader. Compiled with `tsc`.
- **`packages/cli`** — `dvload` CLI (login, run, validate, schedule, pqt). Compiled with `tsc`.
- **`packages/addin`** — Office.js Excel task-pane add-in. Bundled with webpack.

**Read [docs/ARCHITECTURE.md](../docs/ARCHITECTURE.md) before changing the engine.**
It documents the `loadRows` pipeline, the invariants that aren't obvious from the
types (idempotency-gated retries, frontier-based checkpoints, `undefined` vs `null`
in payloads, injection guards on batch bodies), and the multi-file checklist each
kind of change needs. [docs/DATA-FORMATS.md](../docs/DATA-FORMATS.md) covers
coercion rules, the `$batch` wire format, the QDEFF binary layout, and every
on-disk artifact.

`.dvmap.json` is the contract between the add-in and the CLI. Any breaking change
to `packages/core/src/mapping.ts` bumps `SCHEMA_VERSION` **and** updates
`docs/schema/dvmap.schema.json`. Same for `run-plan.ts` and `dvplan.schema.json`.
Verify with `node tests/schema-parity.mjs`, which fails when the schemas and the
hand-rolled validators disagree.

## Build

```powershell
npm install          # from repo root
npm run build        # builds all three workspaces in order
```

Build individual workspace:
```powershell
npm run build --workspace=@dvload/addin
```

## Add-in dev workflow

```powershell
# Terminal 1 — webpack dev server (HTTPS, port 3000, hot-reload)
npm run dev:addin

# Terminal 2 — sideload into Excel desktop
npm --workspace=@dvload/addin run start

# Stop sideloading
npm --workspace=@dvload/addin run stop
```

**First time on a machine:** the `start` script asks to add a loopback exemption for
Edge WebView. It needs admin rights. If it fails, run this once in admin PowerShell:
```powershell
CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
```

## UI auth

- **The UI does not authenticate.** No MSAL, no client id, no Entra traffic.
  `packages/addin/src/auth.ts` POSTs `/api/token` to the local sidecar
  (`dvload serve`) on its own origin, and caches the token until 5 minutes
  before the JWT `exp`.
- Why: a browser flow needs an `spa` redirect URI on the app registration
  (for the redirect *and* the token endpoint's CORS header), which can't be
  added to Microsoft's pre-consented Dataverse client. A Node process has
  neither constraint. Result: no app registration, no admin consent.
- Therefore **`dvload serve` must be running** or the pane shows a "dvload
  isn't running" screen. Port 44321, hardcoded in both manifests because
  `SourceLocation` is a literal URL.
- `window.fetch` is rebound to `window` in `initAuth()` to avoid the
  "Illegal invocation" error the Office WebView throws on a bare `fetch`.
- `packages/addin/src/host.ts` is the only module that knows whether it's
  running in Excel or a browser. Don't reach for `Office.*` outside it.

## Known quirks

| Issue | Fix |
|---|---|
| `webpack serve --https` unknown option | Removed — HTTPS is configured in `webpack.config.js` via `office-addin-dev-certs` |
| `Cannot find name 'process'` in addin build | Browser tsconfig has no Node types by design; the UI needs no build-time config |
| Pane loads but sign-in does nothing | `dvload serve` isn't running, or the webpack dev server (:3000) is being used instead — it serves no `/api` |
| Test files (`*.test.ts`) pulled into webpack | Excluded via `"exclude"` in `packages/addin/tsconfig.json` |
| `HTMLOptionsCollection` not iterable | Added `"DOM.Iterable"` to lib in `packages/addin/tsconfig.json` |
| Module not found `../auth.js` | Added `extensionAlias: { ".js": [".ts", ".js"] }` to webpack resolve config |
| Sign-in "Illegal invocation" (MSAL + WebView) | `window.fetch = window.fetch.bind(window)` before MSAL init in `auth.ts` |

## CLI auth

```powershell
dvload login --env https://yourorg.crm.dynamics.com        # delegated (browser; --device-code for headless)
dvload app-login --env <url> --client-id <id> --tenant-id <id>  # app-only
dvload whoami --env <url>                                   # check what's configured
```

Credentials live in `~/.dvload/` — **not** Windows Credential Manager. `keytar` was
removed deliberately (unmaintained native binary, main `npm install` failure mode on
locked-down machines). The store is `secrets.dat`, DPAPI-protected at CurrentUser
scope via a PowerShell child process, which is the same protection class as
Credential Manager. On POSIX it's plaintext `secrets.json` at mode 0600. Layout and
account-key naming: [docs/DATA-FORMATS.md](../docs/DATA-FORMATS.md#on-disk-state-dvload).

## Documentation sync

When a feature is added, changed, or a known limit is resolved, update all affected
docs **in the same edit**. Docs are part of the change, not a follow-up.

| Document | Update when |
|---|---|
| `README.md` | user-visible behaviour, flags, mapping fields, known limits |
| `FEATURES.md` | any feature added, or a planned item completed |
| `docs/ARCHITECTURE.md` | module boundaries, pipeline phases, invariants, extension points |
| `docs/DATA-FORMATS.md` | coercion rules, wire format, on-disk artifacts, env vars |
| `docs/schema/*.schema.json` | any change to the `.dvmap.json` / `.dvplan.json` shape |
| `docs/AUTH-NOTES.md` | anything learned about Entra, consent, or Conditional Access — record what was *verified* separately from what was assumed |
| `TEST-PROTOCOL.md` | any feature that needs manual/E2E coverage |
| `PRE-RELEASE-CHECKLIST.md` | items that must be done before shipping |

Two claims to keep honest, because they've drifted before: where credentials are
stored, and which platforms scheduling supports.
