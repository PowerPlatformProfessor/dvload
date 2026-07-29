# Contributing to dvload

Thanks for wanting to help.

Before the details, the one thing worth internalising: **dvload writes to
production Dataverse environments, frequently unattended on a schedule.** A
bug here doesn't usually surface as a crash. It surfaces as a thousand
duplicated contacts, or a silently truncated date, discovered three weeks
later by someone who has no idea a tool was involved.

That's why the review bar leans on tests rather than on description, and why
the checks below are blocking rather than advisory.

---

## Getting set up

```bash
git clone https://github.com/PowerPlatformProfessor/dvload.git
cd dvload
npm ci
npm run build     # the CLI and add-in type-check against core's output
npm test
```

Node 20 or newer. Windows is the primary platform for the CLI (DPAPI token
storage, Task Scheduler, Excel COM), but the engine and its tests run
identically on macOS and Linux, and CI covers both.

You do **not** need a Dataverse environment to develop or to run the test
suite. The integration tests run against an in-process fake of the Dataverse
Web API — see [TESTING.md](./TESTING.md).

---

## The workflow

1. **Open an issue first** for anything beyond a bug fix or a docs change.
   It's much cheaper to disagree about an approach in an issue than in a
   review of finished code.
2. Fork, branch from `main`.
3. Make the change, with tests.
4. Run `npm run verify` — this is exactly what CI runs.
5. Open the PR. Fill in the template; the "Risk" section is the most useful
   part for a reviewer.

### Before you push

```bash
npm run verify
```

Which is:

| Step | What it catches |
|---|---|
| `npm run lint` | Floating promises, swallowed errors, `any`, tests with no assertions |
| `npm run typecheck` | Type errors, including in test files |
| `npm run test:coverage` | The suite, plus per-area coverage floors |
| `npm run coverage:gate` | Coverage dropping below the recorded baseline |
| `npm run test:docs` | JSON schemas and README recipes drifting from the parser |

Faster loops while you work:

```bash
npm run test:watch                 # everything, on change
npx vitest --project core          # just the engine's unit tests
npx vitest --project core-int      # just the integration tests
npx vitest run -t "sync"           # tests whose name matches
```

---

## What gets a PR merged

### Every behaviour change needs a test that fails without it

Not "a test that covers the area" — a test that goes red if you revert your
production change. If you can't write one, that's worth saying in the PR;
sometimes it means the design needs to change to be testable.

### Test at the right layer

| Layer | Where | Use it for |
|---|---|---|
| Unit | `packages/*/src/**/*.test.ts` | Pure logic: coercion, mapping validation, key building |
| Integration | `packages/core/test/integration/` | Anything that reaches Dataverse. Runs the real client against a fake server |
| Property | `packages/core/test/property/` | Rules that must hold for *every* input, especially encoding and coercion |
| Live E2E | `packages/core/test/e2e/` | Behaviour only a real environment shows: throttling, alternate keys, plugins |

If you're changing how requests are built or how results are counted, an
integration test is worth more than three unit tests with a stubbed client. A
hand-written fake client will happily agree with a buggy serializer.

### Coverage can go up, not down

The gate compares against `.github/coverage-baseline.json` and fails on a
drop. If you legitimately can't avoid a decrease, run
`npm run coverage:accept` and explain why in the PR description — the baseline
change shows up in the diff and gets reviewed like any other change.

### Conventional Commits

The PR title becomes the squash-merge commit message and feeds the changelog:

```
fix(core): cap Retry-After at maxDelayMs
feat(cli): add --dry-run to run-all
docs: explain alternate-key setup for upsert
```

Types: `feat`, `fix`, `perf`, `refactor`, `docs`, `test`, `build`, `ci`,
`chore`, `revert`. Optional scope: `core`, `cli`, `addin`, `docs`, `deps`,
`ci`, `schema`.

### Never commit

Environment URLs, tenant ids, client ids or secrets, certificates, tokens, or
anything exported from a real environment. See [SECURITY.md](./SECURITY.md).

---

## Areas that need extra care

These files have code owners and get a closer read. Not to discourage changes
— to set expectations about review depth.

**`packages/core/src/coerce.ts`** — every cell value passes through here. A
change to date handling can shift a day; a change to number handling can
truncate a currency. Add property tests, not just examples.

**`packages/core/src/dataverse.ts`** — batch bodies are built by string
concatenation from untrusted spreadsheet content. Anything touching URL or
header construction needs a corresponding case in
`test/integration/request-smuggling.test.ts`.

**`packages/core/src/load.ts`** — the counters (`created`, `updated`,
`skipped`, `unchanged`, `failed`) are what users see and what scheduled runs
alert on. They must always add up to `total`; there's a test asserting exactly
that, and it should stay true for any new code path.

**The sync path** — `conflictMode: "sync"` deactivates or deletes records
absent from the source. There is a guard that refuses to run against an empty
source. Do not weaken it without a very good reason and a very loud test.

**`packages/cli/src/auth.ts` and `secure-store.ts`** — credentials on disk.

---

## Repository layout

```
packages/core     the engine: mapping, coercion, OData client, load orchestration
packages/cli      the `dvload` command
packages/addin    the Excel task-pane add-in
docs/schema       JSON Schema for .dvmap.json and .dvplan.json
tests/            cross-cutting doc/schema parity checks and fixture generators
```

## Branch protection (for maintainers)

Settings → Branches → `main`:

- Require a pull request before merging, with 1 approval
- Require review from Code Owners
- Require status checks to pass: **`CI passed`** (the single aggregate job in
  `ci.yml` — pointing at this one means adding a matrix leg never requires
  editing protection rules, and a skipped job can't be mistaken for a passing
  one)
- Require branches to be up to date before merging
- Require conversation resolution before merging
- Do not allow bypassing the above settings

## Releasing

Tag `vX.Y.Z` on `main`. `release.yml` builds the single-file Windows
executable and attaches it to the GitHub release; `e2e-live.yml` also runs on
tags. Before the first public release, work through
[PRE-RELEASE-CHECKLIST.md](./PRE-RELEASE-CHECKLIST.md) — in particular the
client id and code-signing items.
