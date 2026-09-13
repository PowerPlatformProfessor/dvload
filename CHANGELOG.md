# Changelog

All notable changes to dvload are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and dvload adheres
to [Semantic Versioning](https://semver.org/).

## [0.2.0] - 2026-09-13

### Added

- **Run a `.dvplan.json` from the UI.** The Run plan tab's *Run all steps*
  executes a plan in the pane itself — Excel, `dvload gui`, and the Power
  Platform ToolBox, which has no CLI beside it — using the same stage and
  dependency ordering as `dvload run-all` (now shared from `@dvload/core`).
  Steps run sequentially in the pane; the CLI still runs a stage in parallel.
- **Several mappings open at once** on the Import tab: a chip per mapping,
  *+ New mapping*, double-click a chip to rename it, and click to switch the
  whole editor — source, target, columns, options — between them. The set
  survives closing the pane.
- ***Save all → Run plan*** writes every open mapping to its `.dvmap.json`
  and turns the set into run-plan steps, updating a step in place when its
  mapping is re-saved so stages and dependencies tuned on the plan tab
  survive.
- The plan editor is file-aware and visual: *Add current mapping as step*
  saves the mapping file the step points at, cards are grouped under **Stage
  N** headings in execution order, and steps can be dragged to reorder
  (dropping onto another stage's card adopts that stage).
- *Run import* states which mapping it is about to run when more than one is
  open, and points at the Run plan tab for running them all.

### Fixed

- Create-table mode showed an empty column grid for file sources — it read
  sample rows from the open workbook only, which no browser or ToolBox
  session has.
- The new table's schema name is now editable, with a live `prefix_` preview,
  instead of being derived silently from the display name.
- Switching account in the UI signs the previous user out of every
  environment, so which account an import runs as can no longer depend on
  which environment is selected.
- A run-plan step with no mapping or workbook is reported by name instead of
  sending the pane looking for a file called `""`.
- Telemetry is baked into CLI release builds. It was read only at runtime, so
  the shipped executable could never send anything.

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

[0.2.0]: https://github.com/PowerPlatformProfessor/dvload/releases/tag/v0.2.0
[0.1.0]: https://github.com/PowerPlatformProfessor/dvload/releases/tag/v0.1.0
