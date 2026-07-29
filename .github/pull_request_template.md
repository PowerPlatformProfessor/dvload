<!--
Thanks for contributing to dvload.

This tool writes to production Dataverse environments, often unattended on a
schedule. A bug here doesn't show up as a crash — it shows up as wrong data
that somebody discovers weeks later. That's why the checklist below leans on
tests rather than description.
-->

## What this changes

<!-- One or two sentences. What behaviour is different after this PR? -->

## Why

<!-- The problem, not the solution. Link the issue if there is one. -->

Fixes #

## How it was verified

<!-- Delete what doesn't apply. -->

- [ ] `npm run verify` passes locally (lint, types, tests, coverage, docs)
- [ ] New or changed behaviour has a test that **fails without this change**
- [ ] Tested by hand against a sandbox Dataverse environment
- [ ] N/A — docs / comments / tooling only

## Risk

<!-- Be honest here; it's the most useful section for the reviewer. -->

- [ ] Changes how data is written to Dataverse (coercion, batching, conflict modes)
- [ ] Changes authentication or credential storage
- [ ] Changes the mapping schema or file format
- [ ] Touches the sync / delete path
- [ ] None of the above

<details>
<summary>If any risk box is ticked, describe the blast radius</summary>

What happens to an existing user who upgrades and reruns their current
mapping unchanged? Is any existing mapping file invalidated?

</details>

## Checklist

- [ ] Commits (or the PR title) follow Conventional Commits — `feat:`, `fix:`, `docs:`, `chore:`, `refactor:`, `test:`, `perf:`, `ci:`
- [ ] No secrets, environment URLs, tenant ids or real customer data in code, tests or fixtures
- [ ] Public behaviour changes are reflected in `README.md` / `FEATURES.md`
- [ ] Mapping-schema changes are mirrored in `docs/schema/*.json` (`npm run test:docs` checks this)
