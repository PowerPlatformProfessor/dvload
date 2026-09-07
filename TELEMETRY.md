# Telemetry

dvload can collect **anonymous usage data** to help improve the tool. This
document is the complete reference: every event, every field, how to turn
it off, and what is never collected. The implementation is open source —
audit it in `packages/cli/src/telemetry.ts` and
`packages/addin/src/telemetry.ts`.

## The short version

- **Anonymous, no identifier.** Launches are counted in aggregate. There
  is no install id, cookie, or any other persistent identifier, so events
  can't be linked to you, your machine, your tenant, or your Microsoft
  account — and can't be counted as "unique users" either, by design.
- **No customer data, ever.** See the "Never collected" list below.
- **Announced, and notice-first.** The CLI prints a one-time notice — and
  that first run sends *nothing*, so you can opt out before any data ever
  leaves your machine. Likewise the add-in's first session is silent; the
  checkbox is visible before collection starts.
- **Easy to disable.** `dvload telemetry off`, `DVLOAD_TELEMETRY=0`, the
  cross-tool `DO_NOT_TRACK=1`, or untick the add-in checkbox.
- **Can never slow you down.** Events are fire-and-forget with a 2-second
  cap at exit; offline machines silently drop them.
- **Off in forks.** A build without a connection string baked in cannot
  send anything at all — `dvload telemetry` (status) tells you which kind
  of build you have.

## Never collected

- Dataverse environment URLs (they contain your org name)
- Mapping files or any of their contents (table names, column names,
  targets, constants)
- Cell values or any row data
- Record GUIDs, usernames, email addresses, tenant or user ids
- Error *messages* (they can embed data) — only short error *codes*
- Exact row counts — only coarse buckets (`1-100`, `101-10k`, …)
- Any persistent identifier — no install id, device id, or cookie; and IP
  addresses are not retained (Application Insights stores `0.0.0.0`)

## Events

Every event carries: `toolVersion` and `os` (CLI only, e.g. `win32`).
Nothing that identifies the machine or user — no install id. The count of
events *is* the launch count.

### CLI

| Event | Fields |
|---|---|
| `cli_run` | `mode` (insert/upsert/skip-if-exists/sync), `dryRun`, `refresh`, `resume` (true/false), `concurrency` (1–8), `rows` (bucket), `failed` (bucket), `outcome` (ok/partial/cancelled), `errorCodes` (up to 3 short codes, e.g. `COERCE,400`), `durationSec` |

### Excel add-in

| Event | Fields |
|---|---|
| `addin_open` | — |
| `addin_run` | `mode`, `rows` (bucket), `failed` (bucket), `outcome` (ok/partial/cancelled) |
| `addin_create_table` | `columns` (bucket), `hasKey` (true/false) |

Row-count buckets: `0`, `1-100`, `101-10k`, `10k-100k`, `>100k`.

## Turning it off

**CLI**

```bash
dvload telemetry          # status: on/off, config path, example event
dvload telemetry off      # persist opt-out (~/.dvload/telemetry.json)
dvload telemetry on       # re-enable
```

Environment variables (override the config; useful for CI):

```bash
set DVLOAD_TELEMETRY=0    # dvload-specific
set DO_NOT_TRACK=1        # honored per the console DNT convention
```

**Add-in:** untick "Share anonymous usage data" at the bottom of the task
pane. The choice is stored per machine in the browser storage of the pane.

## Where the data goes

Events are sent to an Azure Application Insights instance operated by the
author, over HTTPS, and used solely to understand which features are used
and where imports fail. No data is sold or shared.

## For maintainers / forks

The connection string is injected at build time from the
`DVLOAD_AI_CONNECTION_STRING` env var (in CI, the repo *variable* of the
same name — it is not a secret, since it ships inside the bundles anyway):

- CLI: baked by `scripts/bundle.mjs` at bundle time. The env var at
  runtime still overrides the baked value (set it empty to silence a
  release build).
- Add-in / browser UI: baked by webpack's DefinePlugin at build time.
- Power Platform ToolBox: **always off.** The ToolBox's CSP has no
  exception for the ingestion endpoint, so events could never leave the
  tool iframe; the PPTB build bakes an empty string deliberately
  (`packages/pptb/webpack.config.js`).

Leave it empty and the telemetry code paths are inert. If you add an
event, document it here in the same commit — this file is the contract.
