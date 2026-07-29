# Testing dvload

How this project is tested, and why it's tested that way.

The short version: dvload writes to production Dataverse environments,
unattended, on a schedule. The failure mode that matters isn't a crash — it's
a load that reports `created: 1000` while having created the wrong thousand
records. The suite is built around catching that.

---

## The layers

| Layer | Location | Runs in | Gates PRs |
|---|---|---|---|
| Unit | `packages/*/src/**/*.test.ts` | ms | ✅ |
| Integration | `packages/core/test/integration/` | ~200 ms | ✅ |
| Property | `packages/core/test/property/` | ~1 s | ✅ |
| Docs & schema parity | `tests/*.mjs` | ms | ✅ |
| Mutation | `stryker.config.json` | minutes | ❌ nightly |
| Live E2E | `packages/core/test/e2e/` | minutes | ❌ nightly + tags |
| Manual protocol | [TEST-PROTOCOL.md](./TEST-PROTOCOL.md) | hours | ❌ pre-release |

```bash
npm test                  # everything that gates a PR
npm run test:unit
npm run test:integration
npm run test:property
npm run test:coverage     # + per-area coverage floors
npm run test:mutation     # slow; normally CI-only
npm run test:e2e          # needs a sandbox; self-skips without one
npm run verify            # exactly what CI runs
```

---

## Integration tests and the fake Dataverse

This is the part worth understanding before writing tests here.

The obvious way to test a loader is to stub the client:

```ts
const client = { batch: async (ops) => ops.map(() => ({ ok: true, status: 201 })) };
```

That's how the original suite worked, and it has a specific blind spot: it
tests `load.ts` against *your idea* of how `DataverseClient` behaves. Batch
body construction, multipart response parsing, retry classification, upsert
status accounting — the most dangerous code in the repo — are never executed.
A hand-written fake will happily agree with a buggy serializer.

So `packages/core/test/support/fake-dataverse.ts` implements the **wire
protocol** instead: `$batch` multipart parsing, alternate-key URLs,
`If-None-Match`, `@odata.nextLink` paging, `Retry-After` throttling,
`@odata.bind` flattening. Tests drive the real `DataverseClient` against it.

```ts
const fake = createFakeDataverse({ entities: contactsFixture() });

const result = await loadRows({
  mapping: contactMapping({ conflictMode: "upsert", upsertKey: ["emailaddress1"] }),
  rows: contactRows(120),
  client: clientFor(fake),
});

expect(result.created).toBe(120);
expect(fake.records("contacts")).toHaveLength(120);   // the real assertion
```

### What it gives you

**Fault injection.**

```ts
fake.throttleNext(2, { retryAfter: 3 });            // 429 with Retry-After
fake.failNextNetwork(1, "ECONNRESET");              // ambiguous mid-flight drop
fake.failNextNetwork(1, "ECONNREFUSED");            // provably never executed
fake.setOperationFault((op) =>                      // per-row server error
  op.seq === 3 ? { status: 400, message: "plugin rejected row" } : null);
```

The two network codes matter: `ECONNRESET` is ambiguous (the server may have
committed), so a plain insert must **not** be replayed, while `ECONNREFUSED`
provably never reached Dataverse and is safe to replay. That distinction is
the difference between a retry and a thousand duplicate records, and it's
tested in `resilience.test.ts`.

**Wire-level assertions.** `fake.requests` records every request with parsed
batch operations, so you can assert on things the result object can't show —
that each row got its own changeset, that `Prefer: odata.continue-on-error`
was sent, that no injected header made it through.

**Real sockets when you need them.** `await fake.listen()` starts an actual
`node:http` server, so undici is in the path (header casing, chunked bodies,
content-length on multipart replies).

### Extending it

If a test needs behaviour the fake doesn't have, **add it to the fake** rather
than dropping to a stubbed client. That's the whole point. The documented
simplifications are at the top of the file: `$filter` supports only the
`attr eq literal` grammar the client actually builds, there's no `$expand` on
records, and no FetchXML.

---

## Property tests

Example-based tests pin the cases somebody thought of. Property tests state
rules that must hold for *every* input — which is what you want from a layer
that turns arbitrary spreadsheet cells into a production write.

The two areas covered:

**`coerce.property.test.ts`** — coercion invariants. `coerceValue` throws only
`CoerceError` whatever the cell contains (a stray `TypeError` would crash a
run mid-load); output is always JSON-serialisable; `dateonly` never shifts the
calendar day; `parseWithFormat` rejects Feb 30 rather than rolling it to Mar 2.

**`odata-encoding.property.test.ts`** — `formatKeyLiteral` is the most
security-sensitive function in the repo: it puts an arbitrary cell inside an
HTTP request line in a multipart body. The properties say it never emits a
structural character, is injective (two different keys can never address the
same record), round-trips through an independently written decoder, and
survives a URL parse without producing a query string or fragment.

### Reproducing a failure

Runs are seeded, so a red build stays red. To replay a reported counterexample:

```bash
FC_SEED=1685482596 npm run test:property
```

Set `FC_SEED=$RANDOM` locally to explore new ground.

### Constraining generators honestly

If a property fails on an input the system can never see, constrain the
generator **and say why in a comment**. `fc.date()` will happily hand you the
year −159498; Dataverse's minimum is 1753-01-01, so bounding it is correct.
Constraining a generator because the failure is inconvenient is not.

---

## Coverage

Measured with V8, reported per area rather than as one number — a single
global figure would be too soft for the engine and too harsh for the CLI
shell, and would hide both.

| Area | Floor (lines) | Why |
|---|---|---|
| `packages/core/src` | 78% | Every bug writes wrong data somewhere |
| `packages/cli/src` | 45% | Much of it only reachable through a spawned process |
| `packages/addin/src` | 34% | Only the pure suggestion logic is in scope |

On top of the floors sits a **ratchet** (`scripts/coverage-gate.mjs`). It
compares against `.github/coverage-baseline.json` and fails on a drop of more
than 0.5 points. Coverage can rise freely and the baseline is raised
automatically; lowering it takes an explicit, reviewable commit.

Why a ratchet rather than a fixed threshold: set the number low and it never
catches anything; set it high and the first PR touching a hard-to-test file
gets blocked until someone lowers it "temporarily". Both end with a number
nobody believes.

### Known gaps

**CLI command modules report 0% despite being tested.** Several CLI tests
spawn the compiled binary as a child process to assert on real exit codes and
stderr — which is the right way to test a CLI, but V8 in-process coverage
can't see it. Those files are excluded from measurement with a comment saying
so, rather than counted as untested (which would pressure contributors into
writing fake in-process tests that assert nothing). Closing this properly
means merging `NODE_V8_COVERAGE` output from the child processes.

**The add-in task pane is not unit-tested.** 1,700 lines of Office.js host
integration. Covered by the manual protocol, §§12-16.

---

## Mutation testing

Coverage says a line ran. Mutation testing says the tests would have
**noticed** if that line were wrong.

Stryker changes one operator or literal at a time and reruns the suite. A
surviving mutant means some line is executed but never actually asserted on.
Scope is `packages/core` only — mutating the CLI shell mostly produces mutants
nobody should care about, at several times the runtime.

It runs nightly, incrementally, and never gates a PR: a full run is minutes,
and a surviving mutant is a conversation ("is this worth asserting?"), not a
merge blocker.

```bash
npm run test:mutation
npx stryker run --mutate "packages/core/src/coerce.ts"   # one file
```

> **First run:** the score will be whatever it is. Read the report, then raise
> `thresholds.break` in `stryker.config.json` to just under it so the score can
> only go up. A break threshold nobody can meet gets the job disabled, and a
> disabled gate is worse than a modest one.

---

## Live E2E

`packages/core/test/e2e/` runs against a real Dataverse sandbox: real
throttling with real `Retry-After` values, real alternate keys, real plugins.

It is **not** a PR gate, and that's deliberate. Fork PRs have no access to
secrets, so requiring it would permanently block outside contributions; and
using `pull_request_target` to hand secrets to fork code would let any PR
exfiltrate them. The mock-server integration suite is what gates PRs instead.

Safety measures, since these tests write to a live environment:

- Every record carries `lastname = "DVLT-E2E"` plus a random per-run marker,
  so cleanup can never touch anything else.
- Cleanup runs in `afterAll` and is idempotent.
- The suite **refuses to start** unless the hostname looks like a sandbox
  (`dev`/`test`/`sandbox`/`uat`/`sit`/`qa`), overridable only by an explicit
  `DVLOAD_E2E_I_KNOW_WHAT_IM_DOING=1`.
- No E2E test uses `conflictMode: "sync"`. A mistake there is unrecoverable.

Setup is documented at the top of `.github/workflows/e2e-live.yml`. Until it's
configured, the job runs and self-skips with a clear message.

---

## Writing a good test here

**Assert on the outcome, not the mechanism.** `expect(fake.records("contacts"))
.toHaveLength(120)` survives a refactor of how batches are chunked;
`expect(fake.requests).toHaveLength(6)` does not, and fails for reasons that
have nothing to do with correctness.

**Make the counters add up.** Anything touching `load.ts` should keep
`succeeded + failed + skipped === total` true. Scheduled runs alert on those
numbers.

**Name the bug, not the function.** `"does not replay a plain insert batch
after an ambiguous drop"` tells a future reader what breaks if it goes red.
`"tests retry logic"` doesn't.

**Leave the reason in the test.** Especially for regressions:

```ts
it("caps a server-supplied Retry-After at maxDelayMs", async () => {
  // Regression: Retry-After used to bypass maxDelayMs entirely, so a
  // Retry-After of 3600 slept for an hour with no output — indistinguishable
  // from a hung process.
```

Six months from now that comment is the difference between fixing the test and
deleting it.

---

## Relationship to the manual protocol

[TEST-PROTOCOL.md](./TEST-PROTOCOL.md) is a ~20-section manual protocol
covering every feature against a real environment. It is still the pre-release
gate, and it still covers things no automated suite here reaches: the add-in
UI, Excel COM, Windows Task Scheduler, conditional-access login flows.

What's changed is that its most tedious and most regression-prone sections —
§3 (unit tests), §6 (core load), §7 (conflict modes), and much of §8 (field
coercion) — now have automated equivalents that run on every PR. The fake
server is deliberately built around the same fixture the protocol describes
(contact with an `emailaddress1` alternate key), so the two are describing the
same environment.
