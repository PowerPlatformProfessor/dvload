# Changelog

All notable changes to dvload are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and dvload adheres
to [Semantic Versioning](https://semver.org/).

## [0.1.0] - 2026-09-03

First public release.

### Added

- **Core engine** (`@dvload/core`): map an Excel table or CSV — typically a
  Power Query output — to a Microsoft Dataverse table and load it through the
  OData Web API. Column mapping with suggestions, lookup resolution by GUID,
  alternate key, or text (including customer and owner polymorphic lookups),
  choice columns, insert / upsert / skip / sync modes, parent-then-child
  ordering, batching with configurable parallelism, dry runs, and
  checkpoint/resume for interrupted runs.
- **Portable mappings**: mappings are saved as `.dvmap.json` files, so a
  mapping built in a UI can be run unattended by the CLI on a schedule.
  Multi-file plans (`.dvplan.json`) validate alternate-key links and execute
  staged dependencies via `dvload run-all`.
- **CLI** (`dvload`): `login`, `whoami`, `run`, `run-all`, `serve`, `gui`,
  and `addin` commands. Distributed as a self-contained signed exe via
  winget and Scoop; Node.js is only needed for npm installs or source builds.
- **Excel task pane add-in**: served from the local `dvload serve` sidecar on
  loopback, so the install needs no admin rights, no hosted pane, and no
  Entra app registration.
- **Browser UI** (`dvload gui`): the same pane in a regular browser tab.
- **Power Platform ToolBox tool** (`@dvload/pptb`): the pane packaged for
  powerplatformtoolbox.com. Runs against PPTB's bridge (`window.dataverseAPI`)
  through a `DataverseGateway` adapter — no token ever reaches the tool.
  Batch, upsert, and conditional-write semantics are emulated per record;
  per-request headers (`bypassCustomLogic`, impersonation) and alternate-key
  creation are not available in this front end.
- **Auth**: delegated sign-in by default using a shared Microsoft client with
  dvload's own app registration as fallback (`DVLOAD_NO_SHARED_CLIENT=1` to
  opt out); app-only auth with a client secret or certificate for scheduled
  runs. Tokens are brokered by the local CLI process — no client id ships in
  any UI bundle.
- **Dataflow import**: import a Power Query dataflow's output to Excel and map
  it onward to Dataverse.

[0.1.0]: https://github.com/PowerPlatformProfessor/dvload/releases/tag/v0.1.0
