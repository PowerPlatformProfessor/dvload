# dvload — Copilot instructions

## What this repo is

A monorepo (npm workspaces) with three packages:
- **`packages/core`** — shared mapping engine, OData client, xlsx reader. Compiled with `tsc`.
- **`packages/cli`** — `dvload` CLI (login, run, validate, schedule, pqt). Compiled with `tsc`.
- **`packages/addin`** — Office.js Excel task-pane add-in. Bundled with webpack.

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

## Add-in auth

- Uses MSAL.js popup flow (`@azure/msal-browser` v3).
- Client ID is injected at **build time** via webpack `DefinePlugin` in
  `packages/addin/webpack.config.js`. Override with env var `DATAVERSE_LOAD_CLIENT_ID`.
- Entra app must have `https://localhost:3000/taskpane.html` as a **Single-page application**
  redirect URI (not Web, not Public client).
- `window.fetch` is rebound to `window` in `initAuth()` to avoid the "Illegal invocation"
  error that MSAL triggers inside the Office WebView.

## Known quirks

| Issue | Fix |
|---|---|
| `webpack serve --https` unknown option | Removed — HTTPS is configured in `webpack.config.js` via `office-addin-dev-certs` |
| `Cannot find name 'process'` in addin build | Browser tsconfig has no Node types; client ID is injected by `DefinePlugin` instead |
| Test files (`*.test.ts`) pulled into webpack | Excluded via `"exclude"` in `packages/addin/tsconfig.json` |
| `HTMLOptionsCollection` not iterable | Added `"DOM.Iterable"` to lib in `packages/addin/tsconfig.json` |
| Module not found `../auth.js` | Added `extensionAlias: { ".js": [".ts", ".js"] }` to webpack resolve config |
| Sign-in "Illegal invocation" (MSAL + WebView) | `window.fetch = window.fetch.bind(window)` before MSAL init in `auth.ts` |

## CLI auth

```powershell
dvload login --env https://yourorg.crm.dynamics.com        # delegated (device-code)
dvload app-login --env <url> --client-id <id> --tenant-id <id>  # app-only
dvload whoami --env <url>                                   # check what's configured
```

Credentials stored in Windows Credential Manager.
