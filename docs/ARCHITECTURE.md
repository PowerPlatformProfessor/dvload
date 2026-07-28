# dvload — architecture

How the code is organised and why. This is the document to read before
changing the engine, and the document to rebuild from if the source were
lost. Behaviour-level reference lives in [README.md](../README.md) and
[FEATURES.md](../FEATURES.md); on-the-wire and on-disk detail lives in
[DATA-FORMATS.md](./DATA-FORMATS.md).

## Contents

- [The one-paragraph version](#the-one-paragraph-version)
- [Package boundaries](#package-boundaries)
- [Data flow](#data-flow)
- [The load pipeline](#the-load-pipeline-loadrows)
- [DataverseClient](#dataverseclient)
- [Coercion layer](#coercion-layer)
- [Source readers](#source-readers)
- [Run plans](#run-plans)
- [Table generation](#table-generation)
- [.pqt codec](#pqt-codec)
- [CLI structure](#cli-structure)
- [Add-in structure](#add-in-structure)
- [Auth](#auth)
- [Invariants](#invariants)
- [Extension points](#extension-points)
- [Testing strategy](#testing-strategy)

## The one-paragraph version

`.dvmap.json` is the contract. The add-in writes it, the CLI reads it, and
`@dvload/core` executes it. Everything else is a front end. The engine takes
(mapping, rows, client) and returns a `LoadResult`; it never touches the
filesystem, never prompts, and never knows whether it's running in Node or
in an Office WebView. That constraint is what lets one implementation of
upsert/sync/lookup-resolution serve both a task pane and a scheduled task.

## Package boundaries

```
packages/core     @dvload/core   — mapping schema, coercion, OData client, load engine, .pqt codec
packages/cli      dvload         — argv → filesystem → core → stdout; auth, scheduling, PQ refresh
packages/addin    @dvload/addin  — Office.js task pane; MSAL popup auth; DOM UI over core
```

Dependency direction is strictly `cli → core` and `addin → core`. There is
no `core → cli` or `addin ↔ cli` edge, and nothing in `core` may import a
Node built-in at module scope.

**The hard rule in `core`:** it must load in a browser. Node-only APIs are
either avoided or reached through a dynamic `await import()` inside the
function that needs them (see `readTableFromFile`'s CSV branch, which
imports `node:fs/promises` lazily). Every reader that needs bytes has a
`*FromBuffer` twin so the add-in — which has a `Blob`, not a path — can use
the same code. Violating this doesn't fail the `tsc` build; it fails at
runtime inside Excel, which is much later and much more annoying.

`core` has three runtime dependencies: `exceljs` (xlsx read/write), `jszip`
(.pqt and QDEFF containers), and nothing else. Validation is hand-rolled
rather than using a schema library — see [Coercion layer](#coercion-layer).

### Module map (`packages/core/src`)

| Module | Role | Depends on |
|---|---|---|
| `types.ts` | Runtime-agnostic types: `SourceRow`, `LoadResult`, `ProgressEvent`, field kinds. Dependency-free by rule. | — |
| `mapping.ts` | The `.dvmap.json` schema, `parseMapping`, `validateMapping`, `validateColumn`, `mappingWarnings`, `sourceValue`. | `types` |
| `run-plan.ts` | The `.dvplan.json` schema, `parseRunPlan`, `validateRunPlan`, legacy `runs[]` upgrade. | — |
| `coerce.ts` | Cell value → Dataverse wire value. `coerceValue`, `coerceRow`, `parseWithFormat`. | `mapping` |
| `dataverse.ts` | OData Web API client: retry, `$batch`, metadata reads, lookup resolution, input validation. | — |
| `load.ts` | The orchestrator. `loadRows` and its helpers. | all of the above |
| `xlsx-reader.ts` | ExcelJS wrapper: read a named table or used range; write the failed-rows workbook. | `types` |
| `csv-reader.ts` | RFC 4180 CSV/TSV parser, same output shape as the xlsx reader. | `types` |
| `tablegen.ts` | Type inference + Dataverse metadata payloads for "create table from source". Pure functions. | — |
| `pqt.ts` | Power Query Template codec, plus the QDEFF/DataMashup reader and writer. | `mapping` |
| `index.ts` | Public surface. Re-exports everything except the Node-only reader entry points, which are named explicitly. | all |

## Data flow

```
                 ┌──────────────────────────────┐
   Excel table   │  add-in task pane            │  builds + saves
   or .xlsx      │  (Office.js, MSAL popup)      │──────────────┐
   or .csv/.tsv  └──────────────────────────────┘              │
        │                                                       ▼
        │                                              ┌──────────────────┐
        │                                              │  .dvmap.json     │
        │                                              │  (the contract)  │
        │                                              └──────────────────┘
        ▼                                                       │
 ┌─────────────────┐   headers + rows    ┌──────────────────┐   │ parseMapping
 │ xlsx-reader     │────────────────────▶│                  │◀──┘
 │ csv-reader      │                     │   loadRows()     │
 │ Office.js excel │                     │   (core/load.ts) │
 └─────────────────┘                     └──────────────────┘
                                            │          ▲
                            BatchOperation[] │          │ BatchResultItem[]
                                             ▼          │
                                       ┌──────────────────────┐
                                       │  DataverseClient     │──▶ /api/data/v9.2/$batch
                                       │  (retry, $batch,     │──▶ EntityDefinitions
                                       │   metadata, lookups) │
                                       └──────────────────────┘
                                             │
                            ProgressEvent    │  getToken()
                                 stream      ▼
                    ┌────────────────────────────────────┐
                    │ CLI: progress bar, .jsonl log,     │
                    │ checkpoint file, failed-rows .xlsx │
                    │ Add-in: progress UI, checkpoint in │
                    │ Office Settings, failed-rows Blob  │
                    └────────────────────────────────────┘
```

Both front ends do the same three things: get rows from somewhere, get a
token from somewhere, then subscribe to `onProgress` and render it. All
divergence is in those three edges.

## The load pipeline (`loadRows`)

`packages/core/src/load.ts`. One exported entry point, called by both
`dvload run` and the task pane's *Run import*.

```
loadRows({ mapping, rows, client, onProgress, dryRun, startOffset, signal })
```

Phases, in order:

**1. Offset clamp.** `startOffset` is rounded *down* to a multiple of
`batchSize` (`clampOffset`). A resume must not begin mid-batch, because
checkpoints are only ever written at batch boundaries. Rows before the
offset are counted as `skipped` immediately.

**2. Lookup pre-resolution** (`buildLookupCache`). Before any write, every
non-GUID lookup column is resolved once per *distinct* source value, not
once per row:

- `lookupResolution: "guid"` — skipped here; validated per-row at build time.
- `"alternateKey"` — one `resolveByKey` GET per distinct value.
- `"text"` — `resolveManyByText` batches distinct values into `or`-filter
  queries, 15 per request, to avoid the N+1 that one-GET-per-value produces
  on a 100k-row column. The response map is keyed by **lowercased** value
  because Dataverse text comparison is case-insensitive.

The cache is `Map<target, Map<sourceValue, { guid? , failure? }>>`.
Resolution problems are recorded as a `failure` string rather than thrown,
so a single bad value fails its own rows with a specific message instead of
aborting the run. Ambiguity (2+ matches) is a failure unless
`duplicateBehavior: "first"`; a miss is a failure unless `createIfMissing`,
which creates the record with only `keyAttribute` set.

In `dryRun` no resolution happens at all — the cache is left empty and
lookup binding is skipped.

**3. Batch slicing.** Rows are pre-sliced into `{ offset, slice }` batches of
`batchSize` so a worker pool can pull from a flat list.

**4. Per-batch operation building** (`buildOperation`, per row):

- `coerceRow` produces the payload for all non-lookup columns.
- Lookups are bound as `target@odata.bind = "/entitySet(guid)"`.
- Clearing a lookup uses the plain navigation property set to `null`, never
  a null `@odata.bind` (the OData deserializer rejects that), and only when
  `conflictMode` can update an existing record — a plain create has nothing
  to clear.
- `conflictMode` decides the verb: `insert` → `POST entitySet`; everything
  else → `PATCH entitySet(<keyExpr>)`, with `If-None-Match: *` added for
  `skip-if-exists` so the server refuses to update.
- A row that throws here (coercion failure, unresolved lookup) is recorded
  as a `RowError` and simply not added to the batch. It never reaches HTTP.

**5. Delta detection** (`applySkipUnchanged`, only when `skipUnchanged` and
`conflictMode` is `upsert`/`sync`). For each PATCH op, GET the target record
with a `$select` of just the non-lookup attributes in the body, in chunks of
10 parallel reads. Compare with `valuesEqual` (kind-aware: dates by epoch,
multichoice by sorted set, numbers numerically, GUIDs case-insensitively).
Attributes that match are dropped from the body; an op with nothing left is
dropped entirely and counted `skipped` + `unchanged`. Lookup binds are always
kept — comparing navigation values would cost another metadata round-trip
per column. A failed read falls back to sending everything, which is the
safe direction.

**6. Execution.** `client.batch(ops, { idempotent: conflictMode !== "insert" })`.
The `idempotent` flag is the whole reason POST creates can't be silently
duplicated by a retry — see [DataverseClient](#dataverseclient).

**7. Result accounting.** Per operation: 201 → `created`, 200/204 →
`updated`, 412 under `skip-if-exists` → `skipped` (the requested outcome,
not an error), anything else → `RowError`. Operations the server returned
*no* response for are turned into explicit errors, so the totals always add
up rather than silently under-reporting.

**8. Worker pool.** `poolSize = clamp(mapping.concurrency, 1, batches.length)`
identical workers pull the next batch index until exhausted. JS is
single-threaded, so shared counter mutation needs no locking — only the HTTP
requests overlap. Each worker checks two stop conditions before taking work:
`signal.aborted` and `maxErrors` reached.

**9. Checkpointing.** With concurrency > 1, batch 7 can finish before batch
5, so the last *completed* batch is not a safe resume point. The pool tracks
a `completedThrough[]` array and advances a **contiguous-completion
frontier**; the emitted `checkpoint` offset is the offset of the first
not-yet-completed batch. Resuming from it may redo work that already
succeeded, but never skips work that didn't.

**10. Unattempted rows.** If the run was cancelled or tripped `maxErrors`,
rows in batches never started are added to `skipped` so `total` still equals
the sum of the outcomes.

**11. Sync removal** (`syncRemoveMissing`, `conflictMode: "sync"` only, and
never on a cancelled or dry run). Build the set of key tuples present in the
source, `queryAll` the target entity set selecting the key attributes,
and `PATCH statecode=1` (deactivate) or `DELETE` everything whose tuple is
absent. Three safety properties worth preserving:

- **Refuses to run on 0 source rows.** An empty source would otherwise
  deactivate the entire table. This throws rather than warning.
- **Tuple parts are joined with NUL (`\u0000`).** A printable separator lets
  `["a b","c"]` and `["a","b c"]` collide, which would remove the wrong
  records.
- **GUID key parts are lowercased** (`normalizeKeyPart`). Dataverse returns
  GUIDs lowercase; a spreadsheet holding uppercase ids would look entirely
  absent from the source and sync would remove every row it was meant to
  keep.
- **Never runs after cancellation** — removal decisions from a partially
  loaded source would delete records that simply weren't reached yet.

### Cancellation contract

`signal` is cooperative and deliberately weak: in-flight batches are allowed
to complete, because their rows were already sent and will exist in
Dataverse whether or not we wait for the response. Aborting the fetch would
make the result *less* accurate, not more. Unattempted rows count as
`skipped`, the result carries `cancelled: true`, and the sync pass is
suppressed.

### Progress event stream

`start` → (`row-success` | `row-error`)* interleaved with `batch` and
`checkpoint` → `sync`* → `done`. Consumers should treat it as append-only
and not assume ordering between `row-*` and `batch` events. `checkpoint`
carries the resume offset described above; `done` carries the final
`LoadResult`, which is also the function's return value.

## DataverseClient

`packages/core/src/dataverse.ts`. A thin, dependency-free OData client that
takes a `getToken: () => Promise<string>` thunk so it pins no MSAL flow and
works identically under Node and in a WebView.

Base URL is `${environmentUrl}/api/data/${apiVersion}` (default `v9.2`).

### Retry classification

`fetchWithRetry` handles two failure shapes that are easy to conflate:

- **A response with a retryable status** (429/503/504 by default). Honour
  `Retry-After` — integer seconds *or* HTTP-date, per RFC 7231 — else
  exponential backoff `baseDelay * 2^(attempt-1)`, capped at 60s. The body
  is drained before sleeping so the connection isn't leaked.
- **A thrown request** (no response at all: machine slept, VPN reconnected,
  DNS died). `classifyNetworkError` walks the `cause` chain — undici hides
  the real code under a generic "fetch failed" — and returns:
  - `abort` — caller cancelled. Never replayed.
  - `not-executed` — `ENOTFOUND`, `EAI_AGAIN`, `ECONNREFUSED`,
    `UND_ERR_CONNECT_TIMEOUT`, `ERR_SOCKET_CONNECTION_TIMEOUT`. The server
    provably never saw it, so replay is safe **even for POST**.
  - `ambiguous` — anything else. The request may have executed. Replayed
    only when the caller passes `idempotent`.

`create()` passes `idempotent: false`, as does `batch()` in `insert` mode,
which also narrows retryable statuses to `[429]` only — a 429 is guaranteed
not-executed, whereas a 503/504 can arrive after the server already
committed. This is the mechanism that makes "retry aggressively" safe
without duplicating records.

Every attempt, including retried and thrown ones, is reported to
`onRequest` with `status: 0` for throws. The CLI turns that into the
`.jsonl` run log.

### $batch: one changeset per operation

`batch()` wraps **each** operation in its own changeset. A single shared
changeset is atomic: Dataverse aborts it on the first failure and rolls
everything back, which silently invalidated per-row accounting and made
`skip-if-exists` abort a whole batch on the first existing record.

Two things make per-operation status reporting work:

- `Prefer: odata.continue-on-error` on the outer request. Without it,
  Dataverse stops at the first failed changeset and returns the rest
  unexecuted under an outer 400 — even though earlier changesets committed.
- Parsing the multipart body **even when the outer status is not 2xx**,
  provided it looks like a batch response (`^--batchresponse`). Some
  failures and some proxies surface per-operation outcomes under a non-2xx
  envelope; treating that as a whole-batch failure would report
  successfully-created rows as failed, and a naive re-run would duplicate
  them.

Response parsing splits on every boundary line (`^--…`) rather than
extracting one boundary, because each changeset response carries its own
server-generated boundary. Bodies are single-line JSON, so no body line can
start with `--`. Results are correlated back to operations by `Content-ID`
and re-sorted into the original order.

Created-record ids come from the `OData-EntityId` response header, falling
back to `@odata.id` in the representation body. Never from de-pluralising
the entity-set name — that breaks on `opportunities`, `addresses`, and
`people`/`systemuserid`.

### Injection defence

Batch bodies are built by string concatenation, with embedded HTTP request
lines. Every value that reaches one comes from a mapping file or a
spreadsheet cell, i.e. untrusted. Four guards, all in `dataverse.ts`:

| Guard | Applies to |
|---|---|
| `assertLogicalName` (`^[A-Za-z_][A-Za-z0-9_]*$`) | entity sets, attributes, `$select` members, upsert-key attribute names |
| `assertGuid` | solution ids and other GUIDs interpolated into URLs |
| `formatKeyLiteral` | key values: doubles `'`, percent-encodes, and additionally escapes `( ) ' ! * ~` that `encodeURIComponent` leaves raw |
| `assertSafeBatchUrl` + a CRLF check on every custom header | last line of defence before concatenation |

A cell containing CRLF would otherwise inject headers or entire extra
operations into a changeset. Do not remove these; do not build a URL by
concatenation anywhere else without them.

### Metadata reads

`getEntitySetInfo` (entity set → logical name + primary id/name attribute)
is cached per client instance. Everything else is fetched on demand:
`listEntities`, `getEntityDefinition`, `listSolutions`,
`getSolutionEntityIds` (componenttype 1), `getLookupTargets`,
`getLookupNavigationProperty`, `getOptionSetLabels`.

Two subtleties the add-in depends on:

- **`getLookupNavigationProperty`** exists because the writable single-valued
  navigation property is *not* always the attribute logical name.
  `parentcustomerid` → `parentcustomerid_account`, and some lookup-typed
  attributes (`contact.accountid`) have no writable navigation property at
  all. Binding the wrong name fails the entire payload with "undeclared
  property".
- **`getOptionSetLabels`** tries four metadata casts in turn (Picklist,
  MultiSelectPicklist, Status, State) because the caller only knows the
  field kind, not which cast the attribute answers to.

`queryAll` follows `@odata.nextLink` to exhaustion with
`Prefer: odata.maxpagesize=5000`.

## Coercion layer

`packages/core/src/coerce.ts`. `coerceRow` walks the columns, skips lookups
(they need async resolution, handled in `load.ts`), and returns
`{ payload, lookups }`. `undefined` from `coerceValue` means *omit the
attribute*; `null` means *send null and clear the field* — that distinction
is what `treatEmptyAsNull` controls, and conflating them either clobbers
data or fails to clear it.

The per-kind rules are exhaustively specified in
[DATA-FORMATS.md § Coercion](./DATA-FORMATS.md#coercion-rules). Read that
before changing anything here: several rules look arbitrary and are not
(the `dateonly` ISO short-circuit exists because routing through `new Date()`
shifts a day for strings JS parses as local time).

**Why validation is hand-rolled.** `parseMapping` / `parseRunPlan` are
several hundred lines of explicit type checks with no schema library. This
was deliberate: a runtime schema dependency (zod) caused repeated install
failures on locked-down machines, and `core` ships into a browser bundle
where every dependency is weight. The cost is verbosity; the benefit is
zero runtime dependencies for validation and error messages that name the
exact JSON path. `docs/schema/*.schema.json` mirrors these checks for
editor tooling — it is generated documentation, not the enforcement point.

**Three-layer validation**, by design:

| Layer | Function | When | On failure |
|---|---|---|---|
| Parse | `parseMapping` | loading JSON | throws `MappingParseError` |
| Structure | `validateMapping` / `validateColumn` | before a run; live in the add-in UI | returns `string[]`, blocks the run |
| Advisory | `mappingWarnings` | before a run | returns `string[]`, prints and continues |

`validateColumn` is split out of `validateMapping` specifically so the task
pane can render an error under the offending row instead of one concatenated
banner. Keep it column-local; whole-mapping checks (duplicate targets,
`upsertKey` coverage) belong in `validateMapping`.

## Source readers

Three sources, one output shape: `{ headers: string[], rows: SourceRow[] }`.

- **`xlsx-reader.ts`** (ExcelJS). Resolves a table by name, else the first
  table on the sheet, else — only when no `tableName` was demanded — the
  used range of the first non-empty sheet. `unwrap()` flattens ExcelJS cell
  objects: `.text`, `.result` (formula), `richText` runs joined, hyperlink
  text. Dates stay `Date` objects, which is why `coerce.ts` can use UTC
  getters safely. Blank rows are dropped; blank headers become `col_<n>`.
  ExcelJS is consumed through `any` — its shipped types don't resolve
  reliably under classic module resolution.
- **`csv-reader.ts`**. Hand-written RFC 4180 parser: quoted fields, `""`
  escapes, CRLF or LF, UTF-8 BOM stripped. Values stay strings (empty →
  `null`); no type sniffing, because the coercion layer already does that
  job and doing it twice produces disagreements.
- **Office.js** (`packages/addin/src/excel.ts`). Reads the live workbook via
  `Excel.run`, using `valueTypes` to reconstruct dates from Excel serials.

The failed-rows writer (`writeRowsToFile` / `writeRowsToBuffer`) emits a
workbook whose table has exactly the source headers, so the output is a
valid input to the same mapping. Two entry points for one builder because
the CLI writes a path and the add-in can only hand the user a Blob.

## Run plans

`packages/core/src/run-plan.ts` owns the schema; `packages/cli/src/commands/run-all.ts`
owns execution. A plan orchestrates several single-table mappings without
changing the `Mapping` contract.

`buildExecutionBatches` turns steps into stages:

1. Collect explicit dependencies: `dependsOn` **plus** every
   `alternateKeyLinks[].fromStep` (a key link is a data dependency, so it
   implies ordering without the user restating it).
2. If *any* step declares a `stage`, add an implicit dependency from every
   step to every step with a lower stage number. Stage numbers are therefore
   barriers, not just labels — mixing stages and `dependsOn` is well-defined.
3. Repeatedly emit the ready set (all dependencies satisfied), restricted to
   the lowest stage present in it, preserving file order within a batch.
   Steps in one emitted batch run in parallel via `Promise.all`.
4. An empty ready set with work remaining throws — belt-and-braces behind
   `validateRunPlan`'s DFS cycle detection.

`validatePlanMappings` runs *before* any execution and parses every
referenced mapping, so a typo in step 5 fails the plan instead of failing
after steps 1–4 have written to Dataverse. It also enforces the
`alternateKeyLinks` contract: the upstream step must expose `keyAttribute`
in its `upsertKey`, and the downstream column must be a lookup with
`lookupResolution: "alternateKey"` and a matching `keyAttribute`.

Override precedence is `step.overrides.x ?? commandLine.x`, i.e. the plan
wins over the flag. Legacy `runs[]` manifests are upgraded in memory to
one step per entry, `id: "run-<n>"`, `stage: n` — sequential, as they were.

## Table generation

`packages/core/src/tablegen.ts` is pure: it infers kinds and builds metadata
payloads, and the add-in POSTs them via `DataverseClient`. Kept separate
because type inference is the part worth unit-testing and the HTTP calls are
the part that isn't.

`inferColumnKind` tries, in order: boolean (`true/yes/y` / `false/no/n`) →
number (integer if all integral and within int32, else decimal) → date
(ISO-shaped only; `datetime` if any sample has a time component, else
`dateonly`) → string, escalating to `memo` past 400 characters. String
`MaxLength` is `clamp(maxObserved * 2, 100, 4000)` — headroom for values
longer than the sample.

Creation order matters and is fixed by Dataverse: `EntityDefinitions` POST
carries **only** the primary name attribute, remaining columns follow as
individual `Attributes` POSTs, then `Keys`. The alternate key's backing
index activates asynchronously, so an upsert immediately after creation may
not find it — plain inserts are unaffected.

## .pqt codec

`packages/core/src/pqt.ts` does two independent jobs.

**Container level (easy).** `.pqt` is a plain ZIP: `MashupDocument.pq`,
`MashupMetadata.json`, `Metadata.json`, `[Content_Types].xml`.
`readPqt`/`writePqt` round-trip it. `mappingFromPqt` synthesises a mapping
from `QueriesMetadata[q].FieldsMetadata`; `injectMappingIntoPqt` writes the
reverse.

Two semantic decisions in that translation, both load-bearing:

- `DeleteExistingDataOnLoad` (Dataflow truncate-and-reload) has **no** dvload
  equivalent and is *not* upsert. Import always emits `conflictMode: insert`
  and appends a note to the description; export always writes `false`.
  Mapping it to `upsert` also produced structurally invalid mappings, since
  a `.pqt` cannot supply an `upsertKey`.
- Constant columns are skipped on export — they have no Power Query source
  column to reference.

**Binary level (hard).** Excel doesn't store `.pqt`; it stores a base64
`<DataMashup>` blob in `customXml/item*.xml`, wrapping an inner OPC ZIP that
holds `Formulas/Section1.m`. The reader (`parseQdeffPackage`) and the
experimental writer (`writeDataMashup`, `buildWorkbookWithQueries`) implement
MS-QDEFF. The byte layout, the XML payloads Excel demands, and the minimal
`.xlsx` scaffolding are specified in
[DATA-FORMATS.md § QDEFF](./DATA-FORMATS.md#qdeff--datamashup-binary-layout) —
that section exists because it is the one part of this repo that cannot be
re-derived from behaviour, only from the wire format.

The reader is reliable; the writer is experimental. Excel is strict about
these parts and the documented fallback is `import-pqt --emit-m` plus a
paste into the Advanced Editor.

## CLI structure

`packages/cli/src/index.ts` is Commander wiring only — one `.command()` per
file in `commands/`. Adding a subcommand means adding a file and one block
here.

`run.ts` exports **`executeRun`** alongside `runCommand`; `run-all.ts` calls
`executeRun` directly rather than shelling out, so a plan shares one process
and one token provider. `runCommand` is a thin wrapper that maps a failed
result onto `process.exitCode = 1`.

The run pipeline, in order: parse mapping → `validateMapping` (exit 2 on
error) → print warnings → apply `--max-errors`/`--concurrency` overrides →
optional resume (checkpoint read, which also *disables* `--refresh`, since
refreshing would change the data under a resume) → optional PQ refresh →
read source → build token provider → construct client → `loadRows` →
write `.jsonl` log, delete the checkpoint, write failed-rows workbook →
notify webhook → print summary.

`--non-interactive` defaults to `!process.stdout.isTTY`, so scheduled and
piped runs fail fast instead of hanging on a device-code prompt written into
a log nobody reads.

Checkpoint writes are serialised through a promise chain
(`checkpointChain`), because concurrent workers emit `checkpoint` events
concurrently and interleaved writes to one file would corrupt it. The chain
is awaited before the file is deleted.

Platform-specific work is isolated in two modules, both shelling out to
PowerShell because there is no clean COM or DPAPI story from Node:

- **`refresh.ts`** — `New-Object -ComObject Excel.Application`, force
  `BackgroundQuery = $false` on Power Query connections (types 1 and 7) to
  make `RefreshAll()` synchronous, spin on `CalculationState`, save, close.
  **The timeout lives inside the PowerShell script**, bounded by a
  Stopwatch, with `finally` blocks that always `Quit()`. An earlier design
  killed the process from Node, which skipped `finally` and left orphaned
  headless `EXCEL.EXE` on a machine that runs nightly imports. Node keeps a
  kill timer only as a last resort, 30s past the script's own deadline.
- **`schedule.ts`** — writes a `.cmd` wrapper into `~/.dvload/tasks/` and
  points `schtasks /TR` at the file rather than embedding a command line
  that `cmd.exe` would re-parse. The wrapper pins the absolute node
  executable and CLI entry script (or the SEA exe when there's no script on
  disk) so the task survives a PATH that doesn't include `dvload`. Literal
  `%` is doubled for batch expansion; quotes and newlines in paths are
  rejected outright.

## Add-in structure

`packages/addin/src/`:

| File | Role |
|---|---|
| `taskpane/taskpane.ts` | The pane. Single `AppState` object, imperative re-render functions. |
| `taskpane/taskpane.html` | Static markup; every control has a stable `id` that `el<T>(id)` looks up. |
| `auth.ts` | MSAL.js popup flow, dev-client detection. |
| `excel.ts` | Office.js table listing and reading. |
| `suggest.ts` | Column-name similarity → suggested mappings. The one unit-tested module here. |
| `combobox.ts` | Type-ahead wrapper (`enhanceSelect`) over a plain `<select>`. |
| `telemetry.ts` | Opt-in event reporting. |
| `commands/` | Ribbon command surface (near-empty; the pane is the product). |

`taskpane.ts` is ~3,000 lines and deliberately framework-free — a bundled
framework is dead weight inside an Office WebView, and the pane is one
screen. The pattern throughout is: mutate `state`, then call the matching
`rerender*` / `render*` function. There is no diffing and no reactivity;
if a control doesn't update, the fix is a missing re-render call.

`AppState` holds the whole pane: account, environment, tables, file source,
selected entity + its attributes, the entity list, `mappings`
(`ColumnMapping[]`), `planSteps`, `planMeta`, `checkpoint`, and
**`mappingExtras`**. That last field is the round-trip guarantee:
mapping-level fields the pane has no control for (`createdAt`, `logDir`,
anything added to the schema later) are parked there and spread back in by
`buildMapping()`, so loading a hand-written or CLI-written mapping and
re-saving it doesn't silently drop them. Any new schema field that the pane
doesn't render must survive through this path.

Three metadata caches keyed by name — `lookupTargetsCache`,
`optionLabelsCache`, `solutionEntityIdsCache` — sit next to `state` and are
not cleared on entity change, only on reload.

Persistence, all documented in
[DATA-FORMATS.md § Add-in state](./DATA-FORMATS.md#add-in-persisted-state):
`Office.context.document.settings` (per workbook) holds
`dvload:lastMapping`, `dvload:lastRunPlan`, `dvload:runCheckpoint`;
`localStorage` (per user/host) holds `dvload:profiles` and MSAL's token
cache.

Only writable attributes are offered as targets: `IsValidForCreate ||
IsValidForUpdate`, which filters out `fullname`, `contact.accountid`, and
friends that would fail at run time.

### Build

Webpack, not `tsc`, because the output is a browser bundle:

- Client id is injected at build time via `DefinePlugin` as
  `ADDIN_CLIENT_ID`, from `DATAVERSE_LOAD_CLIENT_ID`. Production builds
  refuse the borrowed dev client id.
- HTTPS on :3000 via `office-addin-dev-certs`, configured in
  `webpack.config.js` (not a `--https` CLI flag).
- `resolve.extensionAlias { ".js": [".ts", ".js"] }` so `core`'s
  ESM-style `./foo.js` imports resolve to TypeScript sources.
- `*.test.ts` excluded in `packages/addin/tsconfig.json` or webpack pulls
  test files into the bundle.
- `lib` includes `DOM.Iterable`; the browser tsconfig has no Node types,
  which is why the client id can't come from `process.env` at runtime.
- `window.fetch = window.fetch.bind(window)` before MSAL init, or MSAL
  throws "Illegal invocation" inside the Office WebView.

## Auth

Two independent implementations by necessity — see
[docs/AUTH-NOTES.md](./AUTH-NOTES.md) for the empirical findings behind the
CLI's design, including what was actually tested versus assumed.

**CLI** (`packages/cli/src/auth.ts`, msal-node). Delegated *or* app-only,
app-only preferred when configured, `--user` forces delegated.

The delegated flow tries a **chain** of client ids:
`resolveClientIdChain()` returns `[shared Microsoft Dataverse client, dvload's
own client]`, or a single pinned id when `--client-id` /
`DATAVERSE_LOAD_CLIENT_ID` is set, or dvload-only under
`DVLOAD_NO_SHARED_CLIENT=1`. Falling through to the next client happens only
when the failure is *about the app* (`shouldTryNextClient`); failures about
the user — Conditional Access, MFA, declined consent — are surfaced, not
retried, because retrying can't help and hides the real cause.
`isConditionalAccessBlock` / `conditionalAccessHint` turn `AADSTS53003` into
actionable advice (retry with `--interactive`; device-code flow is what CA's
Authentication Flows condition usually blocks).

The successful client id is **stored** (`delegatedClient:<host>`) because
MSAL keys cache entries by client id: silent acquisition with a different id
finds nothing and drops the user back into an interactive prompt.

Flow selection: browser auth-code + PKCE with a loopback listener by
default, device code when there's no browser to open (`defaultLoginFlow`
inspects SSH/display-server env), overridable via flags or
`DVLOAD_AUTH_FLOW`.

Scope is always `<environment origin>/.default`.

**Add-in** (`packages/addin/src/auth.ts`, msal-browser). MSAL popup only,
against a real app registration. It *cannot* borrow the shared Microsoft
client, and this is not a library choice: a browser auth-code flow needs a
registered redirect URI, Entra only returns
`Access-Control-Allow-Origin` from the token endpoint when the caller's
origin matches an `spa`-type redirect URI, you cannot add a redirect URI to
a Microsoft-owned app, and Entra rejects `spa` URIs in non-SPA flows.
`redirectUri` is computed as `origin + pathname` so the same bundle works at
`localhost:3000/taskpane.html` and under a hosted subpath.

Credential storage is described in
[DATA-FORMATS.md § On-disk state](./DATA-FORMATS.md#on-disk-state-dvload).
Note that nothing is stored in Windows Credential Manager — the store is
`~/.dvload/secrets.dat`, DPAPI-protected at CurrentUser scope, which is the
same protection class. `keytar` was removed deliberately: an unmaintained
native `.node` binary was both a supply-chain risk and the main
`npm install` failure mode on locked-down machines.

## Invariants

Break these and something fails in a way that's hard to trace back:

1. **`core` must load in a browser.** No Node built-ins at module scope.
2. **`.dvmap.json` is a versioned contract.** Any breaking change to
   `mapping.ts` bumps `SCHEMA_VERSION` and must be reflected in
   `docs/schema/dvmap.schema.json`.
3. **`undefined` ≠ `null` in a payload.** Omit vs. clear.
4. **Never interpolate an unvalidated value into a URL or batch body.**
   Use `assertLogicalName` / `assertGuid` / `formatKeyLiteral`.
5. **Never replay a possibly-executed POST.** The `idempotent` flag is the
   only thing standing between an aggressive retry policy and duplicate
   records.
6. **Checkpoints are batch-aligned and frontier-based**, never "last batch
   that finished".
7. **Sync mode never runs on an empty source or a cancelled run.**
8. **Counters must reconcile:** `total == succeeded + failed + skipped` for
   every exit path, including cancellation and `maxErrors`. `created`,
   `updated` are subsets of `succeeded`; `unchanged` is a subset of
   `skipped`.
9. **The add-in must round-trip unknown mapping fields** via
   `state.mappingExtras`.
10. **Tests and docs move with the code.** See
    [.github/copilot-instructions.md](../.github/copilot-instructions.md).

## Extension points

**A new field kind.** Add to `DataverseFieldKind` in `types.ts` → add a
`case` in `coerceValue` (the `never` exhaustiveness check will fail the
build until you do) → add a comparison branch in `valuesEqual` if equality
isn't string equality → add to `FIELD_KINDS` in `mapping.ts` → add to the
add-in's `matchesKind` filter → update `docs/schema/dvmap.schema.json`,
FEATURES.md, and DATA-FORMATS.md's coercion table.

**A new conflict mode.** `ConflictMode` in `types.ts` → `CONFLICT_MODES` in
`mapping.ts` → a `case` in `buildOperation` → the `validateMapping` rules
about `upsertKey` / `skipUnchanged` → decide its `idempotent` value in
`loadRows`' call to `client.batch` (this is the easy thing to forget) →
the add-in's conflict-mode `<select>`.

**A new source format.** Write a reader returning
`{ headers, rows: SourceRow[] }` with a `*FromBuffer` variant, export it
explicitly from `index.ts` (not `export *`, if it touches Node), and add a
branch to `readTableFromFile`'s extension dispatch.

**A new CLI subcommand.** One file in `commands/`, one `.command()` block in
`index.ts`. If it needs a token, use `getTokenProvider`; don't construct
MSAL directly.

**A new mapping field.** `Mapping` interface → `parseMapping` (with a
default) → `validateMapping` if it constrains other fields →
`docs/schema/dvmap.schema.json` → README's field table. If the add-in
doesn't render it, verify it survives via `mappingExtras`.

## Testing strategy

Unit tests sit beside their subjects as `*.test.ts` and run on Node's
built-in test runner (`npm test`, per workspace). The load engine is tested
by injecting a fake `fetch` into `DataverseClient` — that's the seam, and
it's why `DataverseClientOptions.fetch` exists.

Coverage is deliberately concentrated where behaviour is subtle rather than
spread evenly: `load.test.ts` and `dataverse.test.ts` (batch parsing,
accounting), `network-retry.test.ts` (the classification matrix),
`cancel.test.ts` / `batch-abort.test.ts` (cancellation semantics),
`upsert-primary-key.test.ts`, `lookup-clear.test.ts`,
`csv-and-format.test.ts`, `constants.test.ts`, `run-plan.test.ts`,
`tablegen.test.ts`, `pqt.test.ts`, and the CLI's auth/login/run/validate
suites.

**Schema parity.** `node tests/schema-parity.mjs` (needs `ajv`, and a built
`core`) runs ~70 documents through both `parseMapping`/`parseRunPlan` and
`docs/schema/*.schema.json` and fails if they disagree. The asymmetry it
enforces matters: a schema that rejects something the parser accepts would
red-underline valid mapping files in the user's editor, so that direction is
always a failure. The reverse is allowed only for referential rules draft-07
cannot express — cross-references between array items, uniqueness of a nested
property, cycle detection — and each exemption must be named in the script's
`KNOWN_SCHEMA_GAPS`, which also fails if an entry goes stale. Run it after any
change to `mapping.ts` or `run-plan.ts`.

What unit tests can't cover — Excel COM, Office.js, real Entra flows, real
Dataverse throttling at 100k rows — is covered by
[TEST-PROTOCOL.md](../TEST-PROTOCOL.md), a 22-section manual protocol with
seeded data generators in `tests/dummy-data/make-load-data.js`. Treat the
two as one suite: the protocol is the integration half, and it is the
closest thing to an executable specification this repo has.
