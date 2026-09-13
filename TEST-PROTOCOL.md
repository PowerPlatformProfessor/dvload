# dvload — Pre-release test protocol

Manual/E2E protocol covering every feature in [FEATURES.md](./FEATURES.md), with load tests up to 100k rows. Run against a **dev/sandbox Dataverse environment only** — never production. Complements (does not replace) [PRE-RELEASE-CHECKLIST.md](./PRE-RELEASE-CHECKLIST.md).

Record each test as **Pass / Fail / Blocked / N/A** in the Result column. A release candidate requires all non-experimental sections passing.

---

## 1. Test environment setup (once)

| # | Step |
|---|---|
| 1.1 | Sandbox environment URL: `https://________.crm.dynamics.com`. Confirm it is not production. |
| 1.2 | Create an **alternate key** on `contact` → `emailaddress1` (Power Apps → Tables → Contact → Keys → New, name e.g. `dvload_email_key`). Wait until key status = Active. Required for upsert/sync/skip-if-exists tests. |
| 1.3 | Create a **choice column** on `contact` (e.g. `dvlt_tier`, options: Bronze=1, Silver=2, Gold=3) and a **multichoice column** (e.g. `dvlt_tags`). Required for §9 option-set tests. |
| 1.4 | Test user has sysadmin (or at least create/write/delete on contact+account, plus `prvBypassCustomBusinessLogic` for §17). |
| 1.5 | Optional for §17: register a simple synchronous plugin or Power Automate flow on contact create, to observe bypass behavior. |
| 1.6 | Build workspace: `npm install && npm run build` from repo root — completes with no errors. |

All generated test records carry `lastname = DVLT-Load` or `DVLT-Edge` — cleanup is a bulk delete on that filter (§22).

## 2. Test data generation

From `tests/dummy-data/`:

```bash
node make-load-data.js --edge                        # edge-cases.xlsx (17 tricky rows)
node make-load-data.js --rows 1000                   # contacts_1k.xlsx
node make-load-data.js --rows 10000                  # contacts_10k.xlsx
node make-load-data.js --rows 100000                 # contacts_100k.xlsx (~5 MB)
node make-load-data.js --rows 1000 --mutate 10       # contacts_1k_mutated.xlsx (delta tests)
node make-load-data.js --rows 1000 --errors 25       # 25 injected bad rows (error tests)
node make-load-data.js --rows 1000 --csv             # CSV variant
```

Data is deterministic (seeded): re-running a command reproduces identical files; `--mutate` differs from the base file in exactly the mutated cells. Mapping files `contacts-load-insert.dvmap.json` and `contacts-load-upsert.dvmap.json` sit alongside — **edit `environmentUrl` in both** before starting.

| # | Test | Expected | Result |
|---|---|---|---|
| 2.1 | Generate all files above | Each prints `Wrote … (N rows)`; files open in Excel with a proper table (`tblContacts` / `tblEdgeCases`) | |
| 2.2 | Re-run `--rows 1000` and binary-compare table contents | Identical data (determinism) | |

## 3. Automated unit tests

| # | Test | Expected | Result |
|---|---|---|---|
| 3.1 | `npm test -w @dvload/core` | All pass | |
| 3.2 | `npm test` in `packages/cli` (auth, login, run, validate suites) | All pass | |
| 3.3 | Add-in suggest tests (`suggest.test.ts`) | All pass | |

## 4. Auth & profiles

> Rows 4.2, 4.5, 4.7 and 4.8, plus sections 5 and 6, are automated by
> `tests\unattended-smoke.ps1`. It is dry-run by default and writes nothing
> without `-Write`:
>
> ```powershell
> $env:DVLOAD_SECRET = "<secret>"
> .\tests\unattended-smoke.ps1 -ClientId <app-id> -TenantId <tenant-id>
> ```
>
> Run it before working through the table by hand; it fails fast on the
> credential problems that otherwise show up halfway down as confusing
> data errors.

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 4.1 | Delegated login (browser) | `dvload login --env <url>` | System browser opens; success page; token cached in `~/.dvload/` (DPAPI-encrypted, not plaintext — open the file and confirm) | |
| 4.1b | Delegated login (device code) | `dvload login --env <url> --device-code` | Device-code prompt; succeeds, or fails with the Conditional Access explanation if the tenant blocks the flow | |
| 4.1c | Client fallback | `DVLOAD_NO_SHARED_CLIENT=1 dvload login --env <url>` | Signs in as dvload's own app (may need one-time admin consent); `whoami` reports that client id | |
| 4.2 | whoami (delegated) | `dvload whoami --env <url>` | Shows delegated mode, correct user, which client id was used, live token probe succeeds | |
| 4.3 | Token reuse | Re-run a command without logging in again | No new prompt of any kind — confirms the stored client id matched the MSAL cache | |
| 4.4 | Logout | `dvload logout` then `whoami` | Cache cleared; whoami reports not signed in | |
| 4.5 | App-only (secret) | `dvload app-login --env <url>` with client id + secret | Stored in secure store; `whoami` shows app-only; run works as the Application User | |
| 4.6 | App-only (certificate) | `dvload app-login --cert <pem>` | Same as 4.5 via cert | |
| 4.7 | Precedence | With both configured, run `dvload run`; then again with `--user` | Default prefers app-only; `--user` forces delegated (check identity via createdby on new records) | |
| 4.8 | App-logout | `dvload app-logout` | Secret removed; CLI falls back to delegated | |
| 4.9 | Profiles | `dvload profile add/list/remove`; use profile name instead of `--env` | Stored in `~/.dvload/profiles.json`; commands resolve the URL | |
| 4.10 | Dev-mode warning | With no custom client id set | `[dev mode]` warning on `login`/`whoami` (expected to **disappear** after PRE-RELEASE-CHECKLIST §1 is done — retest then) | |
| 4.11 | Bad credentials | `app-login` with wrong secret | Clear error, nothing stored | |

## 5. Validation

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 5.1 | Valid mapping | `dvload validate contacts-load-insert.dvmap.json` | Passes local + remote checks | |
| 5.2 | Offline | Same with `--no-remote` | Local checks only, no network | |
| 5.3 | Schema errors | Break the file (unknown `conflictMode`, `upsert` without `upsertKey`, lookup with `lookupResolution: "text"` but no `keyAttribute`, `skipUnchanged` with `insert`) — one at a time | Each rejected with a specific, understandable message | |
| 5.4 | Remote mismatch | Set `target` to a nonexistent attribute; set `targetEntitySet` to a nonexistent entity | Remote validation flags both | |

## 6. Core load — insert & dry-run

Use `contacts-load-insert.dvmap.json` + `contacts_1k.xlsx`.

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 6.1 | Dry run | `dvload run … --dry-run` | No records created in Dataverse; row/column preview & counts printed | |
| 6.2 | Insert 1k | `dvload run …` | `created: 1000, failed: 0`; spot-check 5 records in Dataverse (all 8 fields populated, types correct: birthdate date-only, creditlimit money, donotemail boolean) | |
| 6.3 | Run log | Inspect `logs/*.jsonl` | One entry per event; final `done` entry matches console counts | |
| 6.4 | `--json` | Re-run with `--json` on a fresh 1k file | Machine-readable result on stdout, parseable, counts correct; suitable for pipelines | |
| 6.5 | `--non-interactive` | Run with no cached token and `--non-interactive` | Fails fast with clear error instead of prompting | |

## 7. Conflict modes

Requires alternate key from §1.2. Use `contacts-load-upsert.dvmap.json`.

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 7.1 | Upsert = update | Load `contacts_1k.xlsx` (already inserted in 6.2) with upsert mapping | `updated ≈ 1000, created: 0` (or `unchanged` if skipUnchanged active — see §10) | |
| 7.2 | Upsert = mixed | Load `contacts_1k_mutated.xlsx` after deleting ~100 of the target records | Deleted ones re-created, rest updated/unchanged; totals add up | |
| 7.3 | skip-if-exists | Switch mapping to `skip-if-exists`, re-run 1k | `skipped: 1000, created: 0` | |
| 7.4 | sync (deactivate) | Set `conflictMode: "sync"`; delete 50 rows from a copy of the workbook; run | Missing 50 target records deactivated (statecode inactive); `removed: 50` | |
| 7.5 | sync (delete) | Same with `syncAction: "delete"` | Records physically deleted | |
| 7.6 | sync empty-source guard | Run sync against a workbook whose table has 0 data rows | Refuses to run (protects against wiping the target) | |

## 8. Field types & coercion (edge cases)

Load `edge-cases.xlsx` with an insert mapping (copy `contacts-load-insert.dvmap.json`, change `sourceTable` to `tblEdgeCases`, `maxErrors` ≥ 10). The Notes column states each row's intent.

| # | Test | Expected | Result |
|---|---|---|---|
| 8.1 | Unicode rows (diacritics, CJK, emoji) | Stored intact, verified in Dataverse UI | |
| 8.2 | 4000-char memo | Stored complete in `description` | |
| 8.3 | Empty cells + `treatEmptyAsNull` | Attributes null, not `""`/0 | |
| 8.4 | Whitespace padding, quotes, embedded newline | Preserved or sanely handled; no crash | |
| 8.5 | Numbers/booleans as text (`"12345.67"`, `"1"`, `"Yes"`, `"No"`) | Coerced correctly | |
| 8.6 | Negative money, int32 max | Accepted | |
| 8.7 | Bad rows (`not-a-date`, `12,34abc`, date below Dataverse min) | Row-level errors with row index + reason; run continues; good rows unaffected | |
| 8.8 | Duplicate upsert key (two rows, same email) | Re-run edge file in upsert mode: deterministic, documented behavior (last-write or error — record which) | |
| 8.9 | Choice/multichoice/status/state | Extend the mapping with `dvlt_tier` (`choice` + `optionMap` label→value) and `dvlt_tags` (`multichoice`); add label values in the sheet | Labels resolved to option values; bad label → row error | |
| 8.10 | datetime & uniqueidentifier kinds | Add temp columns mapped as `datetime` / `uniqueidentifier` (e.g. GUID into a custom field) | ISO UTC stored correctly; invalid GUID → row error | |

## 9. Lookups

`contact.parentcustomerid` (polymorphic account/contact) is ideal. Create 5 accounts named `DVLT-Acct-1…5` first; add an `Account` column to a small generated workbook.

| # | Test | Config | Expected | Result |
|---|---|---|---|---|
| 9.1 | By GUID | `lookupResolution: "guid"`, cells contain account GUIDs | Linked | |
| 9.2 | By alternate key | Alternate key on account (e.g. accountnumber), `lookupResolution: "alternateKey"` + `keyAttribute` | Linked | |
| 9.3 | By text | `lookupResolution: "text"`, `keyAttribute: "name"`, cells contain `DVLT-Acct-3` etc. | Resolved and linked | |
| 9.4 | Text, no match | Cell `DVLT-Acct-nope`, `createIfMissing: false` | Row error | |
| 9.5 | createIfMissing | Same with `createIfMissing: true` | Account auto-created, then linked | |
| 9.6 | Ambiguous match | Two accounts with identical names; `duplicateBehavior: "error"` then `"first"` | Error / first match respectively | |
| 9.7 | Polymorphic bindEntitySet | In the add-in, map a column to `parentcustomerid` | UI offers valid target entity sets; chosen `bindEntitySet` saved in the exported mapping and honored by CLI | |

## 10. Delta detection (`skipUnchanged`)

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 10.1 | All unchanged | Upsert `contacts_1k.xlsx` twice in a row | 2nd run: `unchanged: 1000`, 0 PATCHes (verify `modifiedon` did not move) | |
| 10.2 | Partial change | Then upsert `contacts_1k_mutated.xlsx` | Exactly ~100 updated, ~900 unchanged | |
| 10.3 | Attribute stripping | Enable auditing on contact for one run | Updated rows only write the changed attribute(s), not all 8 | |

## 11. Error handling & failed-rows workflow

Use `contacts_1k_err25.xlsx`-style file (`--errors 25`) with the insert mapping.

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 11.1 | Row errors don't kill the run | `maxErrors: 100`, run | `failed: 25`, 975 created; each error names row + cause | |
| 11.2 | maxErrors guard | Set `maxErrors: 10`, fresh run | Run stops early after 10 failures; remaining rows counted as `skipped` | |
| 11.3 | Failed-rows file | After 11.1, check `logs/failed_*.xlsx` | Contains exactly the 25 failed rows, same column shape as source | |
| 11.4 | Fix & re-run | Correct the bad cells in the failed-rows file, run it as source | 25 created, 0 failed — loop closes | |
| 11.5 | `--no-failed-rows` | Re-run with flag | No failed-rows file written | |
| 11.6 | `--max-errors` CLI flag | Overrides mapping value | Flag wins | |

## 12. Checkpoint / resume

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 12.1 | Interrupt | Start `contacts_10k.xlsx` insert; Ctrl+C around halfway | Checkpoint persisted; console says how to resume | |
| 12.2 | Resume | `dvload run … --resume` | Continues from checkpoint; final total = 10,000 with **no duplicates** (verify count in Dataverse) | |
| 12.3 | Changed workbook guard | Interrupt, edit the workbook, `--resume` | Refuses (workbook changed) | |
| 12.4 | Kill -9 | Repeat 12.1 killing the process hard | Same recovery — checkpoint survives ungraceful death | |

## 13. Load & performance matrix

Insert mapping, fresh target each run (bulk-delete `DVLT-Load` between runs, §22). Record wall time, rows/sec, failures, and any Retry-After/429 messages. Watch CLI memory (Task Manager) on 100k runs.

| Run | File | batchSize | concurrency | Duration | Rows/s | Failed | 429s seen | Result |
|---|---|---|---|---|---|---|---|---|
| P1 | contacts_1k | 100 | 1 | | | | | |
| P2 | contacts_1k | 500 | 4 | | | | | |
| P3 | contacts_10k | 500 | 1 | | | | | |
| P4 | contacts_10k | 500 | 4 | | | | | |
| P5 | contacts_10k | 1000 | 8 | | | | | |
| P6 | contacts_100k | 500 | 4 | | | | | |
| P7 | contacts_100k | 1000 | 8 | | | | | |
| P8 (stress) | contacts_100k | 10 | 8 | | | | | |

| # | Check | Expected | Result |
|---|---|---|---|
| 13.1 | Scaling sanity | P2 ≥ P1, P4 > P3 throughput; concurrency 8 doesn't corrupt counts (created sums correct every run) | |
| 13.2 | Throttling recovery | P8 deliberately provokes Dataverse service protection (many small batches). Run must honor Retry-After: slows down, **completes with 0 lost rows**, no crash | |
| 13.3 | Memory | 100k runs stay flat-ish (no unbounded growth / OOM); note peak RSS | |
| 13.4 | Log size | 100k-run `.jsonl` is written incrementally and openable | |
| 13.5 | batchSize cap | Set `batchSize: 1500` | Rejected or clamped to Dataverse cap 1000 (record which) | |
| 13.6 | Progress events | During P6, progress output updates smoothly (start/batch/checkpoint visible) | |

## 14. run-all (manifest)

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 14.1 | Ordered run | Manifest: accounts mapping first, then contacts with text-lookup to those accounts | Runs in declared order; lookups resolve because dependency loaded first | |
| 14.2 | stopOnError: true | Make mapping 1 fail (bad entity set) | Mapping 2 not attempted | |
| 14.3 | stopOnError: false | Same failure | Mapping 2 still runs; summary reports both | |

## 15. Webhook notification

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 15.1 | `--notify-url` | Point at a webhook.site URL (or Teams incoming webhook); run 1k | One POST with `{text}` summary; counts match the run | |
| 15.2 | Mapping `notifyUrl` | Same via mapping property | Works; CLI flag overrides mapping | |
| 15.3 | Dead webhook | Unreachable URL | Run itself still succeeds; warning logged | |

## 16. Bypass custom logic & impersonation

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 16.1 | bypassCustomLogic | With the §1.5 plugin/flow active, run with `bypassCustomLogic: true` | Plugin/flow does **not** fire; without the flag it does | |
| 16.2 | Missing privilege | Run 16.1 as a user without the bypass privilege | Clear Dataverse error, not silent ignore | |
| 16.3 | impersonateUserId | Set to another user's systemuserid | New records' `createdby` = impersonated user | |

## 17. Power Query / .pqt round-trip

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 17.1 | extract-pqt | Extract from a workbook with 2+ PQ queries | Valid `.pqt`; M code intact; optional mapping injected into `MashupMetadata.json` | |
| 17.2 | import-pqt | Synthesize mapping from that `.pqt` | Plausible `.dvmap.json`; `--all-queries` emits one per query; `--emit-m` writes M document | |
| 17.3 | pqt-to-xlsx (experimental) | Build `.xlsx` from `.pqt` | Excel opens it; queries appear connection-only in PQ editor; "Load To…" works. Failures here are non-blocking but must be documented | |
| 17.4 | Dataflow export | Import a real Dataverse Dataflow / PQ Online export via add-in pane | Queries listed with field-mapping counts; grid populated from a query; M code copyable | |

## 18. `--refresh` & scheduling (Windows)

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 18.1 | `--refresh` | Workbook with a PQ query feeding the table; run with `--refresh` | PQ refresh executes before load; new source data reflected | |
| 18.2 | schedule | `dvload schedule` for a nightly run | Task appears in Windows Task Scheduler with correct action/trigger | |
| 18.3 | Unattended run | Trigger the task manually while logged out of the CLI session, app-only auth configured | Run completes without any interactive prompt; log + webhook confirm | |
| 18.4 | Unattended failure path | Same but with auth broken | Task fails visibly (exit code, log), not hanging on a prompt | |

## 19. Excel add-in (task pane)

Sideload per `packages/addin` dev instructions. Test in Excel desktop.

| # | Test | Expected | Result |
|---|---|---|---|
| 19.0 | No sidecar | Close `dvload serve`, open the pane: "dvload isn't running" with a Retry button, not a blank pane or a stack trace | |
| 19.1 | Sign-in | Browser opens from the *sidecar* process (no Office popup); pane shows the account after it completes. Info banner names the shared Microsoft client | |
| 19.2 | Browser UI | `dvload gui` opens the same pane; table picker disabled with "Not available outside Excel", file picker works, mapping round-trips via localStorage | |
| 19.3 | Sidecar isolation | With `serve` running, open any other site and `fetch('https://localhost:44321/api/token', {method:'POST'})` from its console — blocked by CORS, no token returned | |
| 19.2 | Table list | All tables in workbook listed; refreshes when a table is added | |
| 19.3 | Entity + solution picker | Entities load; selecting a solution filters to its tables; Default solution = all | |
| 19.4 | Mapping UI + auto-suggest | Suggestions match on name similarity (verify "First Name"→firstname); manual override works | |
| 19.5 | Import options UI | All options settable: conflict mode ×4, upsert key, sync action, batch size, concurrency, bypass, skipUnchanged | |
| 19.6 | Option-set labels | Picking choice/status/state target auto-fills `optionMap` from metadata | |
| 19.7 | Run import | 1k import from the pane completes with progress + correct counts | |
| 19.8 | Persistence | Mapping survives close/reopen of workbook (Office Settings); profiles survive Excel restart (localStorage) | |
| 19.9 | Save mapping → CLI | Mapping saved from add-in runs unmodified via `dvload run` (portability promise) | |
| 19.10 | Large table | Open workbook with `contacts_100k` table | Pane stays responsive; import works or fails gracefully with guidance (record behavior) | |
| 19.11 | Errors surfaced | Import with known-bad rows | Row errors visible in pane, not just console | |
| 19.12 | Tabs | Three tabs: Import / Run plan / Dataflows. Only one panel visible at a time; ←/→ move between them when a tab has focus; the Run plan tab shows a step count once the plan has steps | |
| 19.13 | Account before environment | With no session at all, step 1 says "Sign in once…"; clicking **Sign in** with no environment set asks for one rather than failing. After signing in, step 1 shows the username and step 2 ticks the profiles that account already has a session for | |
| 19.14 | Environment switch keeps the user | Signed in to A, pick profile B from the list: the pane signs in to B automatically with the username pre-filled (usually no account picker in the browser), and the connection bar ends on B with the same username. Typing a URL instead must NOT open a browser — it prompts to click Sign in | |
| 19.15 | Switch user | Choose "Sign in with another account…": Entra shows the account picker (no hint), and the new username replaces the old one in step 1 and in the connection bar. The previous user is signed out everywhere — step 1's picker offers only the new username, and `dvload whoami` in a terminal confirms the old sessions are gone. Cancelling the Entra page instead leaves the old account signed in | |
| 19.16 | Cross-tab handoff | On the Dataflows tab, "Use mapping" (or a dataflow import with *Mapping files* only) lands on the Import tab with the mapping loaded; "Use all" lands on the Run plan tab with the new steps. A dataflow imported *with* a workbook stays put so the Download button is still reachable | |
| 19.17 | Status is shared | Start an import, switch to the Run plan tab while it runs | Progress and the final summary stay visible from every tab | |
| 19.18 | Create table from a file source | With an **added file** (not a workbook table) selected — in Excel and again in `dvload gui`, where files are the only source kind — choose "＋ Create new table from source…": the panel lists every source column with suggested names/types. Regression guard: this used to read rows off the live workbook only, leaving the grid empty for file sources | |
| 19.19 | Table names are editable | In the create-table panel: the **Schema name** field follows the display name as you type (e.g. "Order Lines" → `new_OrderLines`), stops following once edited by hand, and the created table carries exactly the shown `prefix_SchemaName`. Changing the prefix updates the `prefix_` echo | |
| 19.20 | Run plan in the pane | Run plan tab → **Run all steps** with a two-step plan (parent + child via alternate-key link, workbooks added on the Import tab): offers to save the plan first, prompts once for the `.dvmap.json` files, shows one confirm listing stages, executes parent before child, per-step counts land in the status log, and `stopOnError` halts on a failing step. A step whose workbook name isn't among the added files is named in the error before anything writes. The same `.dvplan.json` runs unmodified via `dvload run-all` | |
| 19.23 | Several mappings open at once | Import tab: name a mapping, **+ New mapping**, build a second against a different source. Chips show both, the active one is highlighted, and the name field renames the active chip as you type. Clicking a chip restores that mapping whole — source table, target entity, columns, conflict mode, upsert key, batch size. Closing and reopening the pane brings back every mapping *and* leaves the editor showing the one the active chip names. Closing a chip (✕) leaves the others intact; the last one can't be closed. Reset clears them back to one | |
| 19.25 | Run import runs one mapping | With two or more mappings open, **Run import**: the confirm leads with "This runs ONLY the open mapping “<name>” — the other N … are not included", pointing at Run all steps on the Run plan tab, and the running/finished status lines are prefixed "“<name>” only —". With a single mapping open none of that prefixing appears. A dry run (which skips the confirm when there are no warnings) still says it on the status line | |
| 19.26 | Rename a mapping from its chip | Double-click a mapping chip: it becomes an editable field with the name selected. Enter commits, Escape reverts, clicking away commits. The name also lands in Import options → Advanced → Mapping name, and survives closing and reopening the pane. Double-clicking an *inactive* chip switches to it and renames it in one go — the field must not disappear while that switch happens | |
| 19.24 | Save all → Run plan | With two mappings open, **Save all → Run plan**: one save per mapping (download / ToolBox dialog), the pane lands on the Run plan tab with a step per mapping — mapping path matching the saved file name, workbook path matching the added source file, sequential stages. Two mappings sharing a name are refused up front, naming the file they'd collide on, before anything is written. Editing a mapping and saving all again updates its existing step in place, keeping any stage/dependsOn/overrides set on the plan tab, and **Run all steps** afterwards doesn't re-ask for those mapping files | |
| 19.21 | Plan editor: step ↔ file | **Add current mapping as step** saves the `.dvmap.json` (download / ToolBox save dialog) and the new step points at that exact file name, with the workbook set to the added file the mapping reads. Cancelling the ToolBox save dialog adds no step. Running the plan straight afterwards doesn't re-ask for that mapping | |
| 19.22 | Plan editor: reorder & stages | Cards are grouped under **Stage N** headings in execution order. Editing a step's stage number moves its card under the new heading. Dragging a card's grip onto another card moves it there and adopts that card's stage; dragging works in both directions and leaves no highlight behind. A half-filled step (no mapping/workbook) shows an inline error, is refused by **Run all steps** by name, and still survives closing and reopening the pane | |

## 19b. Power Platform ToolBox tool

Build `packages/pptb` (`npm run build --workspace=@dvload/pptb`), then load
`packages/pptb/dist` via PPTB → Settings → Show Debug Menu → Debug → Load
Local Tool. Needs a ToolBox connection to the sandbox environment.

| # | Test | Expected | Result |
|---|---|---|---|
| 19b.1 | Bootstrap | Tool opens with NO account/environment steps; connection bar shows the ToolBox connection's name + environment host. No requests to localhost:44321 | |
| 19b.2 | No connection | With no active connection, the tool shows an actionable "pick a connection" message, not a blank page | |
| 19b.3 | Entities + metadata | Entity list, solution filter, attribute pickers, lookup targets and option-set labels all load through the bridge | |
| 19b.4 | Insert run | 1k insert from an added `.xlsx` completes with correct counts (expect slower than the add-in — per-record bridge calls, no $batch) | |
| 19b.5 | Upsert emulation | Upsert on an alternate key: existing rows count `updated`, missing rows count `created`; re-run counts all `updated` (or `unchanged` with skipUnchanged) | |
| 19b.6 | skip-if-exists | Existing rows count `skipped` (emulated 412), not failed | |
| 19b.7 | Refused options | With "Bypass plugins" or "Run as user" set (including via a loaded mapping), the run refuses up front with a message naming the add-in/CLI, and the controls stay editable so the option can be cleared in place | |
| 19b.8 | Create table | "Create new table from source" creates table + columns; key columns produce the "alternate key NOT created" note instead of a key | |
| 19b.9 | Saves via dialog | Save mapping / failed rows / run log / dataflow workbook each open the ToolBox save dialog and write the file; cancelling the dialog keeps pending downloads on offer | |
| 19b.10 | Dataflow import | Dataflows list and import (mapping + workbook) with no sidecar running | |
| 19b.11 | Persistence | Mapping and run plan survive closing and reopening the tool (ToolBox settings store) | |
| 19b.12 | Connection switch | Switching the ToolBox connection reloads the tool against the new environment; no stale entity/metadata from the old one | |
| 19b.13 | Portability | A `.dvmap.json` saved here runs unmodified via `dvload run`, and vice versa | |
| 19b.14 | Sync paging | (Record behavior) sync mode against a >5k-row target: verify the removal pass sees all rows, or note the bridge's paging limit in the result | |
| 19b.15 | Run plan in the ToolBox | Same as 19.20, against the ToolBox connection: Run all steps executes parent-then-child through the bridge; a plan step carrying bypass/impersonation (or `user`) is refused/flagged up front naming the step, before any write | |

## 20. CSV / TSV input

| # | Test | Steps | Expected | Result |
|---|---|---|---|---|
| 20.1 | CSV | Run insert mapping against `contacts_500.csv` (set `sourceTable` appropriately) | 500 created; quoted/comma/newline cells parsed correctly | |
| 20.2 | TSV | Convert and repeat | Works | |

## 21. Cross-cutting regression passes

| # | Check | Result |
|---|---|---|
| 21.1 | Every command's `--help` is accurate (run `dvload <cmd> --help` for all 12 commands) | |
| 21.2 | No secrets in any log file, `--json` output, or error message (grep logs for the client secret) | |
| 21.3 | All error messages seen during this protocol were actionable (no raw stack traces for expected failures) | |
| 21.4 | Re-run §3 unit tests after any code fix made during the protocol | |

## 22. Cleanup

Bulk-delete test data in the sandbox: Advanced Find → Contacts where `Last Name` = `DVLT-Load` or `DVLT-Edge` → Bulk Delete; same for accounts named `DVLT-Acct-*`. Delete the scheduled task (18.2), `logs/`, and generated `contacts_*.xlsx`/`.csv` (git-ignore or remove before release).

## 23. Sign-off

| Section | Pass/Fail | Notes |
|---|---|---|
| 2–3 Data & unit tests | | |
| 4 Auth & profiles | | |
| 5 Validation | | |
| 6–7 Load & conflict modes | | |
| 8–9 Field types & lookups | | |
| 10–12 Delta, errors, resume | | |
| 13 Load/performance matrix | | |
| 14–16 run-all, webhook, bypass | | |
| 17–18 PQT & scheduling | | |
| 19 Add-in | | |
| 20–21 CSV & regression | | |

Tester: ____________  Date: ____________  Build/commit: ____________

**Release gate:** all sections Pass (17.3 may carry documented failures as experimental), plus every item in PRE-RELEASE-CHECKLIST.md closed.
