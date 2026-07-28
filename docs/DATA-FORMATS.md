# dvload — data formats

Exact specification of every format dvload reads, writes, or puts on a wire.
This is the reference for behaviour that is invisible from the CLI surface:
coercion rules, the `$batch` body, the QDEFF binary layout, and everything
persisted to disk.

Structure and rationale live in [ARCHITECTURE.md](./ARCHITECTURE.md);
user-facing usage lives in [README.md](../README.md).

## Contents

- [Source reading](#source-reading)
- [Coercion rules](#coercion-rules)
- [Explicit date formats](#explicit-date-formats)
- [Change comparison (`skipUnchanged`)](#change-comparison-skipunchanged)
- [OData wire format](#odata-wire-format)
- [Retry semantics](#retry-semantics)
- [On-disk state (`~/.dvload`)](#on-disk-state-dvload)
- [Run outputs](#run-outputs)
- [Add-in persisted state](#add-in-persisted-state)
- [Environment variables](#environment-variables)
- [.pqt container format](#pqt-container-format)
- [QDEFF / DataMashup binary layout](#qdeff--datamashup-binary-layout)
- [Generated-table metadata payloads](#generated-table-metadata-payloads)

## Source reading

All three readers produce the same shape:

```ts
{ headers: string[], rows: Array<Record<string, unknown>> }
```

Row objects are keyed by header text. **Header matching is exact** — a
mapping's `columns[].source` must equal the header string as read, including
case and internal spaces.

### Rules common to every reader

- Rows where every cell is blank are dropped, so trailing empty rows in a
  worksheet don't become failed rows.
- A blank header becomes `col_<n>`, where `n` is the 1-based column position
  (the sheet column index for xlsx, the field position for CSV).
- Row indices in `RowError.rowIndex` and `RowSuccess.rowIndex` are **0-based
  over the retained rows**, not spreadsheet row numbers. A file with a header
  row plus one dropped blank row will not line up with what Excel shows.

### .xlsx / .xlsm (ExcelJS)

Table resolution order:

1. The table named `sourceTable`, on `sourceSheet` if given, else searched
   across all worksheets.
2. If no `tableName` was requested: the first table found on any sheet.
3. If still nothing **and no `tableName` was requested**: the used range of
   the first non-empty worksheet, treated as header row plus data. Requires
   at least 2 rows.

Requesting a table by name that doesn't exist is always an error — it never
falls back to a sheet scan, because silently loading the wrong range is worse
than failing.

Used range is computed as `A1:<lastColumnLetter><actualRowCount>` from
ExcelJS's `actualRowCount` / `actualColumnCount`.

Cell values are unwrapped in this order:

| Cell content | Value produced |
|---|---|
| Empty | `null` |
| `Date` | the `Date` object, unchanged |
| Object with `.text` (hyperlink, shared string) | `.text` |
| Object with `.result` (formula) | `.result` — the cached result, not the formula |
| Object with `.richText[]` | the runs' `.text` concatenated |
| Anything else | as-is (number, boolean, string) |

Dates surviving as `Date` objects matters: ExcelJS constructs them as UTC, so
the coercion layer's UTC getters are correct. A reader that stringified them
first would reintroduce the timezone ambiguity the `format` option exists to
resolve.

### .csv / .tsv

Hand-written RFC 4180 parser. Delimiter is `,` for `.csv` and `\t` for
`.tsv`, selected from the file extension.

- A field is quoted only if `"` is its **first** character; `""` inside a
  quoted field is a literal quote.
- Row terminators: `\r\n`, `\n`, or bare `\r`.
- A leading UTF-8 BOM (Excel writes one) is stripped.
- A final row without a trailing newline is still emitted.
- Headers are trimmed; data values are **not** trimmed.
- Empty string → `null`. No type sniffing at all: every value stays a string
  and the coercion layer converts it. Sniffing here would produce a second,
  disagreeing set of type rules.

### Live workbook (add-in, Office.js)

Reads via `Excel.run` using the table's header range and data body range, and
uses `valueTypes` to rebuild `Date` values from Excel serial numbers. Same
blank-row and blank-header rules.

## Coercion rules

`coerceValue(value, column)` in `packages/core/src/coerce.ts`. One pass per
cell, no locale awareness, no implicit fallbacks beyond what's listed.

### Blank handling comes first

A value is blank when it is `null`, `undefined`, or a string that is empty
after trimming.

| `treatEmptyAsNull` | Result | Effect on the request |
|---|---|---|
| `true` (default) | `null` | Attribute sent as `null` — **clears** the field |
| `false` | `undefined` | Attribute **omitted** from the payload entirely |

The `undefined` / `null` distinction is load-bearing throughout the engine:
`coerceRow` only copies a value into the payload when it is not `undefined`.
Conflating the two either clobbers existing data or makes clearing a field
impossible.

Lookups are the exception — see [Lookup binding](#lookup-binding).

### Per-kind conversion

| `kind` | Accepts | Produces | Notes |
|---|---|---|---|
| `string`, `memo` | anything | `String(value)` | No trimming, no length check. Dataverse enforces `MaxLength`. |
| `integer`, `status`, `state` | number, or a string that `Number()` parses | `Math.trunc(n)` | Truncates toward zero — `3.7` → `3`, `-3.7` → `-3`. Non-finite throws `CoerceError("not an integer")`. |
| `decimal`, `money`, `double` | number, or numeric string | the number | Non-finite throws `CoerceError("not a number")`. No rounding to the attribute's precision. |
| `boolean` | `true`/`false`, or (case-insensitive, trimmed) `true`/`yes`/`y`/`1` and `false`/`no`/`n`/`0` | boolean | Anything else throws. Note `1`/`0` are accepted as strings *and* as numbers via `String()`. |
| `datetime` | `Date`, or string | ISO 8601 UTC (`toISOString()`) | With `format`: parsed parts are treated as **UTC**, deliberately, so the same file loads identically on any machine. Without `format`: `new Date(String(value))`, i.e. JS parsing rules. Invalid throws. |
| `dateonly` | `Date`, or string | `yyyy-MM-dd` | See below — this one has a deliberate short-circuit. |
| `uniqueidentifier` | anything | `String(value).trim().toLowerCase()` | Not validated here. Validation happens where it matters: as a lookup GUID, or as a primary-id upsert key. |
| `choice` | number, `optionMap` label, or numeric string | integer | Label lookup uses `hasOwnProperty`, so a label named `toString` can't collide with a prototype member. Unmapped non-numeric throws `CoerceError("unknown option")`. |
| `multichoice` | array, or a string split on `,` or `;` | comma-joined integer string, e.g. `"1,3,7"` | Each part is trimmed; blanks dropped; each resolved via `optionMap` then numerically. |
| `lookup` | anything | passed through unchanged | Real handling is in `load.ts`, which needs async resolution and a navigation-property bind. |

**`dateonly` string handling**, in order:

1. With `format`: parse per the format, emit `yyyy-MM-dd` from the parsed
   parts.
2. Without `format`, if the string matches `^(\d{4})-(\d{2})-(\d{2})(?:$|[T ])`:
   take the date part **verbatim**, without constructing a `Date`. This is not
   an optimisation — `new Date("2025-01-01 00:00")` is parsed as *local* time
   by JS, and reading UTC getters back off it shifts the day for anyone west
   of UTC.
3. Otherwise: `new Date(...)` and format from UTC getters.

### Constant columns

`constant` replaces `source` (exactly one must be present). `sourceValue()`
is the single choke point that returns either the cell or the constant, so
constants flow through identical coercion, identical lookup resolution, and
identical validation. They are, however, skipped when exporting to a `.pqt`,
having no Power Query source column to point at.

## Explicit date formats

`column.format` applies to `datetime` and `dateonly` only, and only to string
input. Tokens are matched longest-first at each position:

| Token | Matches | Field |
|---|---|---|
| `yyyy` | 4 digits | year |
| `MM` | exactly 2 digits | month |
| `M` | 1–2 digits | month |
| `dd` | exactly 2 digits | day |
| `d` | 1–2 digits | day |
| `HH` | exactly 2 digits | hour (24h) |
| `H` | 1–2 digits | hour (24h) |
| `mm` | exactly 2 digits | minute |
| `ss` | exactly 2 digits | second |

Every other character is a literal and is regex-escaped. The compiled pattern
is anchored (`^…$`), so a trailing timezone or stray text fails rather than
being ignored. Compiled formats are cached per format string.

There is **no 12-hour, month-name, timezone-offset, or fractional-second
token**. Unmatched fields default to year 1970, month 1, day 1, 00:00:00.

Impossible dates are rejected: the parsed parts are round-tripped through
`Date.UTC` and compared field by field, because `Date.UTC` silently rolls
Feb 30 into Mar 2. A mismatch returns `null`, which surfaces as
`CoerceError('does not match format "…"')`.

Examples: `dd/MM/yyyy`, `yyyy-MM-dd HH:mm`, `d.M.yyyy`.

## Change comparison (`skipUnchanged`)

`valuesEqual(kind, ours, theirs)` decides whether an attribute is dropped
from a PATCH body. Deliberately loose, because "what we're about to send"
and "what Dataverse returned" are differently typed representations of the
same value.

| `kind` | Comparison |
|---|---|
| either side null/undefined | equal only if **both** are null-ish |
| `datetime` | `Date.parse()` on both, compared as epoch ms; unparseable ours ⇒ not equal |
| `multichoice` | split on `,`, trim, drop blanks, **sort**, rejoin, compare — order-insensitive |
| `integer`, `decimal`, `money`, `double`, `choice`, `status`, `state` | `Number(a) === Number(b)` |
| `boolean` | `Boolean(a) === Boolean(b)` |
| `uniqueidentifier` | case-insensitive string compare |
| anything else | `String(a) === String(b)` |

Not compared, and therefore always sent:

- Keys containing `@` — i.e. `@odata.bind` lookup binds. Comparing a
  navigation-property value against a bind path would need another metadata
  round-trip per column.
- Every attribute of a row whose pre-read GET failed. The fallback is to send
  everything, which is the safe direction.

## OData wire format

Base URL: `<environmentUrl>/api/data/v9.2` (version overridable via
`DataverseClientOptions.apiVersion`).

### Standard headers

On every non-batch request:

```http
OData-MaxVersion: 4.0
OData-Version: 4.0
Accept: application/json
Content-Type: application/json; charset=utf-8
Prefer: return=representation,odata.include-annotations="*"
Authorization: Bearer <token>
```

`queryAll` replaces `Prefer` with
`odata.maxpagesize=5000,odata.include-annotations="*"` and follows
`@odata.nextLink` until absent.

### Per-operation headers

Built from the mapping, applied to every operation in a batch:

| Header | Source | Meaning |
|---|---|---|
| `MSCRM.BypassCustomPluginExecution: true` | `bypassCustomLogic` | Skip synchronous plugins |
| `MSCRM.SuppressCallbackRegistrationExpanderJob: true` | `bypassCustomLogic` | Skip Power Automate triggers |
| `MSCRMCallerID: <guid>` | `impersonateUserId` | Act as another systemuser |
| `If-None-Match: *` | `conflictMode: skip-if-exists` | Create only; 412 if the record exists |

### Verb and URL by conflict mode

| `conflictMode` | Request |
|---|---|
| `insert` | `POST <entitySet>` |
| `upsert`, `sync` | `PATCH <entitySet>(<keyExpr>)` |
| `skip-if-exists` | `PATCH <entitySet>(<keyExpr>)` + `If-None-Match: *` |
| `sync` removal pass | `PATCH <entitySet>(<guid>)` with `{"statecode":1}`, or `DELETE <entitySet>(<guid>)` |

### Key expressions

Two forms, chosen by `isPrimaryIdKey` — a single `upsertKey` entry whose
column has `kind: "uniqueidentifier"`:

**Primary id.** `accounts(3f2504e0-4f89-11d3-9a0c-0305e82c3301)` — bare GUID,
unquoted, lowercased. An `Edm.Guid` key rejects a quoted string literal.
`PATCH` on this form creates the record with that id if absent and updates it
if present; no alternate key needs to exist. Two consequences:

- The id column is **deleted from the request body**, since it's already in
  the URL and Dataverse rejects a write to the primary key on an existing
  row — leaving it in would fail every update while letting creates through.
- A blank or malformed id fails that row rather than creating a record with a
  server-generated id.

**Alternate key.** `contacts(emailaddress1='a%40b.com')`, comma-joined for
compound keys: `attr='v1',attr2='v2'`.

Value encoding (`formatKeyLiteral`):

| Input | Output |
|---|---|
| finite number | bare, e.g. `42`. Non-finite throws. |
| boolean | `true` / `false` |
| anything else | `'` + percent-encoded string + `'`, after doubling any `'` per OData rules |

Percent-encoding additionally covers `( ) ' ! * ~`, which
`encodeURIComponent` leaves raw. This is not cosmetic: batch bodies are built
by string concatenation around embedded HTTP request lines, so an unencoded
character with structural meaning is a request-smuggling vector.

A blank `upsertKey` value in a source row fails that row with
`upsertKey attribute "<attr>" is blank in source row` — it is never
substituted or defaulted.

### Lookup binding

Resolved lookups are written as an OData bind annotation:

```json
{ "primarycontactid@odata.bind": "/contacts(3f2504e0-4f89-11d3-9a0c-0305e82c3301)" }
```

The property name is the **writable single-valued navigation property**,
which is not always the attribute logical name (`parentcustomerid` →
`parentcustomerid_account`).

Clearing a lookup does **not** use a null bind. A null `@odata.bind`
annotation is rejected by the OData deserializer ("undeclared property …
only has property annotations but no property value"). Instead the plain
navigation property is set to `null`:

```json
{ "primarycontactid": null }
```

and only when `conflictMode` is `upsert` or `sync` — a plain create has
nothing to clear, so the attribute is omitted.

### $batch body

`POST $batch` with:

```http
Content-Type: multipart/mixed; boundary=batch_<uuid>
Prefer: odata.continue-on-error
```

**Each operation gets its own changeset**, so operations execute
independently and per-row accounting is accurate. Line terminator is CRLF
throughout.

```
--batch_3f8a…
Content-Type: multipart/mixed; boundary=changeset_1_9b1c…

--changeset_1_9b1c…
Content-Type: application/http
Content-Transfer-Encoding: binary
Content-ID: 1

PATCH https://contoso.crm.dynamics.com/api/data/v9.2/contacts(emailaddress1='a%40b.com') HTTP/1.1
Content-Type: application/json; charset=utf-8
OData-MaxVersion: 4.0
OData-Version: 4.0
Prefer: return=representation,odata.include-annotations="*"
MSCRMCallerID: 11111111-2222-3333-4444-555555555555

{"firstname":"Ada","lastname":"Lovelace"}
--changeset_1_9b1c…--
--batch_3f8a…--
```

Notes:

- Request URLs inside the batch are **absolute**, built from the client's
  base URL.
- `Content-ID` is the operation's index within the batch, 1-based, and is
  the only correlation key back to a source row.
- `DELETE` operations have no body; the body line is emitted empty.
- Any custom header containing CR or LF aborts body construction, as does
  any operation URL containing whitespace or CRLF.

### $batch response parsing

The response contains one `changesetresponse` per operation, each with its
own server-generated boundary. Parsing therefore splits on **every** line
starting with `--` and keeps parts containing `HTTP/1.1`. Response bodies are
single-line JSON, so no body line can begin with `--`.

Per part: `Content-ID` → operation, `HTTP/1.1 <nnn>` → status, body after the
first blank line. Results are re-sorted by `Content-ID`.

Status interpretation:

| Status | Meaning |
|---|---|
| 201 | Created — POST, or a PATCH-upsert that inserted |
| 200, 204 | Updated — PATCH found an existing record |
| 412 under `skip-if-exists` | Record existed; counted as `skipped`, not an error |
| other non-2xx | `RowError`, message from `error.message` in the body, else `HTTP <nnn>` |
| no response for an operation | synthesised `RowError` ("no response returned for this operation in the `$batch` reply") so totals reconcile |

Created ids come from the `OData-EntityId` response header, falling back to
`@odata.id` in the representation body. Never from de-pluralising the entity
set name.

An outer non-2xx status is only treated as a whole-batch failure when the
body does **not** match `^--batchresponse`. Some failures and some proxies
report per-operation outcomes under a non-2xx envelope, including operations
that succeeded; throwing there would report created rows as failed and a
naive re-run would duplicate them.

### Text lookup resolution

`resolveManyByText` batches distinct values into `or`-filter queries, **15
per request**:

```
GET <entitySet>?$select=<primaryIdAttribute>,<attribute>
    &$filter=<attribute> eq 'v1' or <attribute> eq 'v2' or …
```

Results are keyed by the **lowercased** matched value, because Dataverse text
comparison is case-insensitive and the response casing may differ from the
spreadsheet's. `resolveByText` (single value, `$top=2`) exists for the
ambiguity check.

`resolveByKey` reads `<entitySet>(<keyAttribute>='<value>')?$select=createdon`
and extracts the GUID from `@odata.id`, rather than guessing the primary-id
attribute name — de-pluralisation fails on `people` → `systemuserid`. A 404
returns `null`.

## Retry semantics

Defaults: `maxAttempts: 5` (including the first), `baseDelayMs: 500`,
`maxDelayMs: 60000`, `retryableStatuses: [429, 503, 504]`.

**Responses.** Delay comes from `Retry-After` per RFC 7231 — integer (or
decimal) seconds, or an HTTP-date, converted to a delay from now, floored at
0. Absent or unparseable, delay is `baseDelay * 2^(attempt-1)` capped at
`maxDelayMs`. There is no jitter. The response body is drained before
sleeping so the connection isn't leaked. `RetryInfo.source` records which
path was taken: `retry-after-seconds`, `retry-after-date`, or `backoff`.

**Thrown requests** (no response at all). `classifyNetworkError` walks up to
5 levels of the `cause` chain, because undici hides the real code beneath a
generic "fetch failed", and collects both `.code` and `.name` at each level:

| Class | Triggers | Replayed? |
|---|---|---|
| `abort` | `AbortError`, `ABORT_ERR` | Never |
| `not-executed` | `ENOTFOUND`, `EAI_AGAIN`, `ECONNREFUSED`, `UND_ERR_CONNECT_TIMEOUT`, `ERR_SOCKET_CONNECTION_TIMEOUT` | Always — the server provably never saw it, even for POST |
| `ambiguous` | anything else | Only when the caller declares the operation idempotent |

Backoff for thrown requests is always exponential (there's no header to
read), reported with `status: 0` and the underlying code in
`RetryInfo.error`.

**Idempotency by caller:**

| Caller | `idempotent` | Retryable statuses |
|---|---|---|
| `batch()` in `insert` mode | `false` | `[429]` only — a 429 is guaranteed not-executed; a 503/504 can arrive after the server committed |
| `batch()` in any other mode | `true` | default |
| `create()` | `false` | default |
| all metadata/query reads | `true` | default |

Every attempt is reported to `onRequest`, including retried and thrown ones.

## On-disk state (`~/.dvload`)

Nothing is stored in Windows Credential Manager. `keytar` was removed
deliberately — an unmaintained native `.node` binary was both a supply-chain
risk and the main `npm install` failure mode on locked-down machines.

```
~/.dvload/
├── secrets.dat                 Windows: DPAPI-protected secret map
├── secrets.json                POSIX:   plaintext secret map, mode 0600
├── msal-cache-<host>.bin       MSAL token cache (refresh tokens), protected
├── msal-cache-<host>.json      legacy plaintext cache; migrated then deleted
├── profiles.json               name → environment URL
└── tasks/
    └── <task name>.cmd         scheduled-task wrapper scripts
```

`<host>` is the environment URL's host, e.g. `contoso.crm.dynamics.com`.

### Protection

On Windows, bytes are round-tripped through
`[Security.Cryptography.ProtectedData]::Protect/Unprotect` at **CurrentUser**
scope via a short-lived PowerShell child process, base64 over stdin/stdout to
avoid quoting problems. Only the same user on the same machine can decrypt —
the same protection class as Credential Manager. Roughly 200ms per call, so
it happens once per process on load and once per mutation, never in a hot
path.

On POSIX there is no DPAPI equivalent in use: files are written plaintext
with mode `0600`. For long-lived deployments prefer certificate auth.

### `secrets.dat` / `secrets.json`

A flat JSON object, `account → secret`, loaded once per process and rewritten
on mutation.

| Account key | Value |
|---|---|
| `delegated:<host>` | MSAL `homeAccountId` of the signed-in user |
| `delegatedTenant:<host>` | tenant id used at sign-in |
| `delegatedClient:<host>` | client id that succeeded — MSAL keys its cache by client id, so silent acquisition must reuse the same one or it finds nothing and reprompts |
| `appOnly:<host>` | JSON `AppOnlyCredentials`: `{ clientId, tenantId, secret? , certificatePem? }` |

A read failure other than `ENOENT` throws rather than degrading to "no
credentials", so a corrupt store is reported instead of silently forcing a
re-login. `ENOENT` yields an empty map.

`detectAuthMode(env)` returns `appOnly` if app-only credentials exist, else
`delegated` if `delegated:<host>` exists, else `none`.

### `profiles.json`

`{ "<name>": "<environmentUrl>" }`, 2-space indented. A JSON parse failure
throws with the file path — a stray comma must not silently become "no
profiles", which would make every `--profile` lookup fail with a misleading
message.

### `tasks/<name>.cmd`

Written by `dvload schedule`; `schtasks /TR` points at the file rather than
an inline command line, so `cmd.exe` never re-parses user paths.

```bat
@echo off
"C:\Program Files\nodejs\node.exe" "C:\...\dist\index.js" run "C:\...\contacts.dvmap.json" -w "C:\...\customers.xlsx" --refresh --non-interactive
```

The filename is the task name with every run of non-`[\w.-]` characters
replaced by `_`. Literal `%` in any path is doubled (`%%`) for batch
expansion. Paths or task names containing `"` or a newline are rejected
outright. In a single-file (SEA) build there is no script on disk, so the
launcher is just the executable path.

Task registration: `schtasks /Create /TN <name> /TR "<script>" /SC DAILY /ST
<HH:MM> /F /RL LIMITED` — user scope, no admin rights, overwriting any
existing task of the same name.

## Run outputs

All three land in `logDir`, resolved relative to the **mapping file**, not
the workbook.

### Run log — `<workbook-stem>_<yyyy-MM-dd>_<HH-mm-ss>.jsonl`

One JSON object per line, in this order:

1. Exactly one summary line: `{"event":"summary", …LoadResult}` with `errors`
   omitted (they get their own lines).
2. One line per HTTP request — a `RequestLogEntry`, no `event` field:
   ```json
   {"method":"POST","url":"https://…/$batch","status":200,"ok":true,
    "startedAt":"2026-07-28T02:30:00.000Z","finishedAt":"2026-07-28T02:30:01.412Z"}
   ```
   Non-ok responses carry `errorBody` (parsed JSON if possible, else raw
   text). **A dropped request appears with `status: 0`** and its error code
   string in `errorBody`.
3. One line per successful row: `{"event":"success","rowIndex":…,"sourceRow":{…},"status":201,"id":"…"}`.
4. One line per failed row: `{"event":"error","rowIndex":…,"sourceRow":{…},"message":"…","code":"…","httpStatus":…}`.

`code` is `COERCE` for coercion failures, the Dataverse error code for
`DataverseError`, and absent otherwise. `rowIndex` is `-1` for errors from
the sync removal pass, whose `sourceRow` is `{ <primaryIdAttribute>: "<guid>" }`.

Timestamps are UTC. The file is written once, at the end of the run — it is
not a streaming log, and a hard kill loses it.

### Checkpoint — `<mapping-stem>.checkpoint.json`

```json
{ "offset": 4000, "workbookSize": 4815873, "workbookMtimeMs": 1753080000000 }
```

- `offset` is the resume row index: the offset of the first batch not yet
  contiguously completed (see ARCHITECTURE's frontier discussion). Always a
  multiple of `batchSize`.
- `workbookSize` / `workbookMtimeMs` are captured once at run start. `--resume`
  refuses the checkpoint if either differs, since resuming against changed
  data would skip rows silently.
- Written on every `checkpoint` event through a serialised promise chain, so
  concurrent workers can't interleave writes.
- Deleted when the run completes, **even if rows failed** — failures are
  handled by the failed-rows file, not by resume. Not written at all on a dry
  run.
- `--resume` also suppresses `--refresh`; refreshing would change the data
  under the resume.

### Failed rows — `failed_<workbook-stem>_<yyyy-MM-dd>_<HH-mm-ss>.xlsx`

A workbook with one worksheet ("Failed rows") holding a table whose name is
the mapping's `sourceTable` and whose columns are exactly the source
headers — so the file is a valid input to the same mapping.

Written when there is at least one failure with `rowIndex >= 0` (sync-pass
errors are excluded, having no source row), `--no-failed-rows` was not
passed, and it isn't a dry run. A write failure warns and continues; it
never fails the run.

Stem sanitisation for all three names: basename without extension, a
trailing `.dvmap` stripped, then every non-`[\w.-]` character replaced by
`_`.

### `--json` output

With `--json`, progress rendering is suppressed and the `LoadResult` is
printed to stdout as a single JSON object. Exit code is 1 when
`result.failed > 0`, 2 when the mapping failed validation.

### Webhook payload

`POST` to `notifyUrl` (or `--notify-url`), `Content-Type: application/json`:

```json
{ "text": "dvload: \"Daily contacts\" → contacts ✓ succeeded. 120 created, 45 updated, 3 skipped (3 unchanged), 0 failed of 168 rows in 12.4s." }
```

Status prefix is `✓ succeeded` or `⚠ completed with errors`. A non-2xx
response or a thrown request warns and is otherwise ignored — the run has
already happened.

## Add-in persisted state

**`Office.context.document.settings`** — scoped to the workbook, travels with
the file:

| Key | Contents |
|---|---|
| `dvload:lastMapping` | the full `Mapping` as JSON, written by `buildMapping()` |
| `dvload:lastRunPlan` | the full `RunPlan` as JSON |
| `dvload:runCheckpoint` | `{ offset, total, sourceTable, savedAt }` — a cancelled run's resume point |

The add-in checkpoint differs from the CLI's: it validates against
`sourceTable` and `total` rather than file size and mtime, since it has no
file to stat. `savedAt` is ISO 8601.

Mapping fields the pane has no control for (`createdAt`, `logDir`, anything
added later) round-trip through `state.mappingExtras` rather than being
dropped on save.

**`localStorage`** — scoped to the user and Office host:

| Key | Contents |
|---|---|
| `dvload:profiles` | `Array<{ name, url }>` |
| MSAL keys | token cache (`cacheLocation: "localStorage"`, `storeAuthStateInCookie: false`) |

## Environment variables

| Variable | Scope | Effect |
|---|---|---|
| `DATAVERSE_LOAD_CLIENT_ID` | CLI runtime; add-in **build** time | Pin the public-client app id. Setting it also disables the shared-client fallback, so the app must register `http://localhost`. For the add-in it's injected by webpack `DefinePlugin` as `ADDIN_CLIENT_ID`, so overriding means rebuilding. |
| `DVLOAD_NO_SHARED_CLIENT=1` | CLI | Skip the shared Microsoft client; use dvload's own registration only. |
| `DVLOAD_AUTH_FLOW` | CLI | `interactive` or `device-code` (`devicecode` also accepted), overriding auto-detection. |
| `DVLOAD_AUTH_TIMEOUT_MS` | CLI | Browser sign-in wait before giving up. Default 3 minutes. |
| `DVLOAD_LOOPBACK_PORT` | CLI | Pin the loopback redirect port instead of taking an ephemeral one. |
| `DATAVERSE_LOAD_DEBUG` | CLI | Verbose auth diagnostics. |
| `DVLOAD_TELEMETRY` | CLI | `0`/`off` disables telemetry. See [TELEMETRY.md](../TELEMETRY.md). |
| `DO_NOT_TRACK` | CLI | Honoured as a telemetry opt-out. |
| `DVLOAD_AI_CONNECTION_STRING` | CLI build/run | Application Insights connection string. Absent ⇒ telemetry is inert. |
| `DVLOAD_WB`, `DVLOAD_TIMEOUT_MS` | internal | Passed to the PowerShell refresh script so paths never go through PowerShell quoting. Not user-facing. |

Flow auto-detection (`defaultLoginFlow`) picks device code when
`SSH_CONNECTION` or `SSH_TTY` is set, or on a non-Windows/non-macOS platform
with neither `DISPLAY` nor `WAYLAND_DISPLAY`; otherwise interactive.

## .pqt container format

A plain ZIP (DEFLATE) with exactly four entries:

| Entry | Contents |
|---|---|
| `MashupDocument.pq` | The M code — identical to Excel's `Formulas/Section1.m` |
| `MashupMetadata.json` | Query metadata and per-query Dataverse field mappings |
| `Metadata.json` | `{ "Name", "Description", "Version" }`; version defaults to `1.0.0.0` |
| `[Content_Types].xml` | Standard OOXML content types; preserved verbatim on round-trip |

`MashupMetadata.json` shape (only the modelled fields; unknown keys are
preserved):

```json
{
  "QueryGroups": [],
  "DocumentLocale": "en-US",
  "FastCombine": false,
  "AllowNativeQueries": false,
  "QueriesMetadata": {
    "Contacts": {
      "QueryId": "<uuid>",
      "QueryName": "Contacts",
      "QueryGroupId": null,
      "EntityName": "contacts",
      "LoadEnabled": true,
      "DeleteExistingDataOnLoad": false,
      "FieldsMetadata": {
        "lastname": {
          "SourceColumnName": "Surname",
          "DestinationFieldType": "String",
          "AutoNumberSettings": null
        }
      }
    }
  }
}
```

`FieldsMetadata` is keyed by **Dataverse attribute** and points back at the
source column — the inverse of a `.dvmap.json` column, which is keyed by
source. Type translation:

| dvload `kind` | Dataflow `DestinationFieldType` |
|---|---|
| `string` | `String` |
| `memo` | `Memo` |
| `integer`, `status`, `state` | `Integer` |
| `decimal` | `Decimal` |
| `double` | `Double` |
| `money` | `Money` |
| `boolean` | `Boolean` |
| `datetime` | `DateAndTime` |
| `dateonly` | `DateOnly` |
| `uniqueidentifier` | `UniqueIdentifier` |
| `lookup` | `Lookup` |
| `choice`, `multichoice` | `Choice` |

The reverse map is 1:1 where possible; `status`, `state`, and `multichoice`
are **not** recoverable from a `.pqt` (they collapse into `Integer` and
`Choice`), and any unrecognised type falls back to `string`.

`DeleteExistingDataOnLoad` means truncate-and-reload in Dataflows and has no
dvload equivalent. Import always emits `conflictMode: "insert"` and appends
a note to the mapping description; export always writes `false`. Mapping it
to `upsert` would both misrepresent the semantics and produce an invalid
mapping, since a `.pqt` can't supply an `upsertKey`.

Query names are discovered from the M document with
`/^[ \t]*shared[ \t]+(?:#"([^"]+)"|([A-Za-z_][A-Za-z0-9_]*))[ \t]*=/gm`,
covering both quoted (`#"My Query"`) and bare identifiers.

## QDEFF / DataMashup binary layout

Excel does not store `.pqt`. It stores a base64 blob in
`customXml/item<n>.xml`:

```xml
<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">UEsDBBQ…</DataMashup>
```

Any `item<n>.xml` may hold it, so all are scanned. The tag may carry
attributes; whitespace inside the base64 is stripped.

Decoded, the blob is MS-QDEFF: a sequence of little-endian `uint32`
length-prefixed sections.

```
offset  size  field
0       4     version (uint32 LE) — normally 0
4       4     packageSize
8       n     package         — an OPC ZIP; this is the part that matters
+4            permissionsSize
        n     permissions     — XML
+4            metadataSize
        n     metadata        — see below
+4            bindingsSize
        n     permissionBindings — XML (written as zero-length)
```

`metadata` is itself length-prefixed:

```
0   4   version (uint32 LE) — written as 0
4   4   xmlLength
8   n   xml     — LocalPackageMetadataFile
+4      contentLength
    n   content — a ZIP (written as an empty ZIP)
```

The reader (`parseQdeffPackage`) needs only the first two fields: it slices
`[8, 8 + packageSize)` and opens it as a ZIP, then reads
`Formulas/Section1.m` (or `Formulas/Section1.pq`). A non-zero version warns
rather than throwing — newer versions still place the package immediately
after the header. A `packageSize` exceeding the blob throws "DataMashup
truncated".

### Writer (experimental)

The inner package contains exactly three entries:

| Entry | Contents |
|---|---|
| `[Content_Types].xml` | defaults for `xml` (`text/xml`) and `m` (`application/x-ms-m`) |
| `Config/Package.xml` | `<Package xmlns="http://schemas.microsoft.com/DataMashup">` with `<Version>2.72.5556.181</Version><MinVersion>2.21.0.0</MinVersion><Culture>en-US</Culture><SafeCombine>true</SafeCombine>` |
| `Formulas/Section1.m` | the M document |

`permissions` is a `<PermissionList>` with
`CanEvaluateFuturePackages=false`, `FirewallEnabled=true`, and
`<WorkbookGroupType xsi:nil="true" />`.

`metadata.xml` is a `<LocalPackageMetadataFile>` whose `<Items>` contains one
`AllFormulas` item with an empty `<ItemPath />`, then one `Formula` item per
query:

```xml
<Item>
  <ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/My%20Query</ItemPath></ItemLocation>
  <StableEntries>
    <Entry Type="IsPrivate" Value="l0" />
    <Entry Type="FillEnabled" Value="l0" />
    <Entry Type="ResultType" Value="sTable" />
  </StableEntries>
</Item>
```

The item path is `Section1/` + `encodeURIComponent(queryName)`, then
XML-escaped. The `Value` prefixes are QDEFF's typed-literal encoding — `l`
for a logical (`l0` = false) and `s` for a string (`sTable`) — inferred from
observed workbooks rather than read off a spec, so treat them as
empirically-derived constants and change them only against a real Excel file.

Query names come from `QueriesMetadata` when present, else from parsing the M
document.

### Minimal .xlsx scaffolding

`buildWorkbookWithQueries` emits the smallest workbook Excel will open with
a Power Query payload attached — seven entries:

```
[Content_Types].xml
_rels/.rels                       → xl/workbook.xml
xl/workbook.xml                   one sheet, "Sheet1"
xl/_rels/workbook.xml.rels        → worksheets/sheet1.xml, ../customXml/item1.xml
xl/worksheets/sheet1.xml          empty <sheetData/>
customXml/item1.xml               <DataMashup>base64</DataMashup>
customXml/itemProps1.xml          datastoreItem with a fresh {UPPERCASE-GUID} itemID
customXml/_rels/item1.xml.rels    → itemProps1.xml
```

`itemProps1.xml` must declare
`<ds:schemaRef ds:uri="http://schemas.microsoft.com/DataMashup" />` or Excel
won't recognise the part.

Queries arrive **connection-only** — the user picks *Load To…* per query.
Excel is strict about these parts; treat the writer as experimental and fall
back to `import-pqt --emit-m` plus a paste into the Advanced Editor.

## Generated-table metadata payloads

`tablegen.ts` builds Web API metadata payloads for "create new table from
source". Labels use `LanguageCode: 1033` throughout.

### Type inference

`inferColumnKind(values)` filters blanks, then tests in order:

1. **boolean** — every sample is a boolean or (trimmed, lowercased) one of
   `true`/`yes`/`y`/`false`/`no`/`n`.
2. **integer / decimal** — every sample parses finite. `integer` when all are
   integral and within ±2147483647, else `decimal`.
3. **datetime / dateonly** — every sample is a `Date`, or a string matching
   `^\d{4}-\d{2}-\d{2}([T ].+)?$` that `Date.parse` accepts. `datetime` if
   any sample has a time component (a `Date` with non-zero h/m/s, or a
   string longer than 10 characters), else `dateonly`.
4. **memo** if the longest sample exceeds 400 characters, else **string**
   with `MaxLength = clamp(maxObserved * 2, 100, 4000)`.

An all-blank column yields `string` with `MaxLength: 100`.

Note the date test accepts **ISO-shaped strings only** — `03/04/2025` infers
as a string, deliberately, because guessing day-vs-month order during table
creation would bake an ambiguity into the schema.

### Naming

`sanitizeSchemaSuffix` splits on non-alphanumerics, PascalCases the words,
strips any leading non-letters, and falls back to `Column`. Collisions get
`1` appended until unique (case-insensitively). Then:

- `SchemaName` = `<prefix>_<suffix>`
- logical name = the same, lowercased
- alternate key `SchemaName` = `<prefix>_key_<entitySuffix lowercased>`

### Payloads

| Kind | `@odata.type` | Key properties |
|---|---|---|
| `string` | `StringAttributeMetadata` | `FormatName: {Value:"Text"}`, `MaxLength`, `IsPrimaryName` when applicable |
| `memo` | `MemoAttributeMetadata` | `MaxLength: 100000` |
| `integer` | `IntegerAttributeMetadata` | `Format:"None"`, min/max int32 |
| `decimal` | `DecimalAttributeMetadata` | `Precision: 2`, ±100000000000 |
| `boolean` | `BooleanAttributeMetadata` | `BooleanOptionSetMetadata` with True=1 "Yes", False=0 "No" |
| `datetime` | `DateTimeAttributeMetadata` | `Format: "DateAndTime"` |
| `dateonly` | `DateTimeAttributeMetadata` | `Format: "DateOnly"` |

All attributes are created with `RequiredLevel: { Value: "None" }`.

Entities are created `OwnershipType: "UserOwned"`, `HasNotes: false`,
`HasActivities: false`, with **only** the primary name attribute in
`Attributes` — Dataverse limits attribute creation inside the entity POST to
the primary name. Creation order is therefore fixed:

1. `POST EntityDefinitions` (entity + primary name attribute)
2. `POST EntityDefinitions(LogicalName='…')/Attributes` per remaining column
3. `POST EntityDefinitions(LogicalName='…')/Keys` for the alternate key

Only `string`, `integer`, and `decimal` columns may participate in an
alternate key (`KEYABLE_KINDS`). The key's backing index activates
**asynchronously**, so an upsert immediately after creation may not find it;
plain inserts are unaffected.
