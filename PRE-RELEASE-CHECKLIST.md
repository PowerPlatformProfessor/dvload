# Pre-release checklist

Things that must change before this tool is shared with anyone outside your
own machine. Track these like blockers — at minimum, the auth and manifest
items are non-negotiable.

## 1. Entra app registration — no longer required

An earlier version of this checklist opened with an app-registration
blocker, on the reasoning that a browser auth-code flow needs an `spa`
redirect URI that a Microsoft-owned app can't have, so the pane had to
ship its own registration. It ended by noting the one alternative: *stop
authenticating in the browser entirely and broker tokens through a local
dvload process.*

That is now what happens. `dvload serve` holds the auth and the UI fetches
tokens from it same-origin, so there is no client id in the UI bundle and
nothing to register. Both front ends use the CLI's client-id chain: shared
Microsoft client first, dvload's own app as fallback. See
[docs/AUTH-NOTES.md](./docs/AUTH-NOTES.md).

**Nothing here is a release blocker any more.** Optional, if you want your
own identity on the consent screen for branding, Conditional Access
targeting, or auditability:

- Register a **public client** (Entra admin center → App registrations →
  New registration), multi-tenant, redirect URI `http://localhost` under
  *Public client / native*. No `spa` redirect URI is needed — nothing
  authenticates from a browser.
- Authentication → Allow public client flows: **Yes**
- API permissions → Add → Dynamics CRM → user_impersonation (delegated)
- Branding & properties → publisher name, support URL, privacy statement,
  terms of service. These show on the consent screen.
- Point dvload at it with `DATAVERSE_LOAD_CLIENT_ID`, or change
  `DVLOAD_CLIENT_ID` in `packages/cli/src/auth.ts` to change only the
  fallback. Note that pinning a client id **disables the shared-client
  chain**, so your app must have `http://localhost` registered or sign-in
  fails with `AADSTS900971`.

**Confidential client (app-only flow) — optional.** Only needed if you
want `app-login` to work without users having to register their own
Entra app first. Most tools omit this from the default install and have
power users register their own; that's fine.

## How to know you're not ready yet

- `dvload serve` fails with "could not find the built add-in UI" — the
  release bundle is missing `build/web`. Run the add-in build before
  `npm run bundle`; the bundle step warns loudly about this.
- `npm --workspace=@dvload/addin run validate` rejects the manifest.
- The pane shows an informational note naming *Microsoft Dynamics CRM*.
  That one is **not** a blocker — it is the shared-client default doing
  its job, and it's shown so the attribution isn't a surprise.

## 2. Office add-in manifest

`packages/addin/manifest.xml` is currently a dev manifest. Before
sideloading anywhere outside your machine, and definitely before
AppSource submission:

- Replace `<Id>00000000-0000-0000-0000-000000000000</Id>` with a real
  GUID. Generate with `[guid]::NewGuid()` in PowerShell or `uuidgen`.
  (`manifest.prod.xml` already has one — ship that.)
- Leave the `https://localhost:44321/...` URLs alone. The pane is served
  by `dvload serve` on loopback, so there is no production host to point
  at; that's deliberate, and it's what keeps the install admin-free. If
  you change the port, change it in both manifests and pass `--port` to
  `serve`.
- Update `<ProviderName>`, `<DisplayName>`, `<Description>`,
  `<SupportUrl>`, and the resources at the bottom (`GetStarted.Title`,
  `GetStarted.Description`, `GetStarted.LearnMoreUrl`, etc.).
- Add real 16x16, 32x32, and 80x80 icons under `packages/addin/assets/`
  and verify they load.
- Validate the manifest: `npm --workspace=@dvload/addin run validate`.

## 3. AppSource submission (only if listing publicly)

> **Read this before starting.** AppSource is a poor fit for the current
> design: a listed add-in must load from a public HTTPS host, and this one
> loads from the user's own machine and requires a companion CLI to be
> running. Centralized Deployment for known orgs, or plain sideloading,
> both work fine. Listing publicly would mean reintroducing a hosted pane
> and therefore its own app registration and admin consent — the thing the
> sidecar exists to avoid. Excel for the web and iPad are out for the same
> reason: no loopback.

- Add a privacy policy and terms of service hosted somewhere durable.
- Test the add-in in Excel desktop (Win + Mac).
- Submit through Partner Center → Office Add-ins.
- Expect 5–10 business days for review; reviewers do test data ops, so
  make sure a tester Dataverse environment is reachable.

## 4. CLI distribution

Mechanics are automated — `.github/workflows/release.yml` verifies, builds
the SEA exe with the UI embedded, signs it, smoke-tests it, and attaches it
plus a SHA-256 to the release. Full detail in
[packaging/README.md](./packaging/README.md). What still needs a human:

- **Set up code signing before the first public tag.** The `Sign` step is
  wired for Azure Trusted Signing but skips itself until the `AZURE_*`
  secrets and variables exist, and an unsigned exe means SmartScreen for
  every user. Check the workflow log said `Signature status: Valid`.
- Push the tag (`v<x.y.z>`) — nothing releases without one.
- After the release exists: `npm run stamp-release -- <x.y.z>`, then submit
  the winget PR and update the Scoop bucket. Neither hash can be filled in
  beforehand.
- `npm publish --workspace=dvload` if the npm path is being kept in sync.
- Confirm the README's install instructions still match reality — they say
  `winget install dvload`, which is only true once the manifest lands.

## 5. Reproducible builds

- Pin `engines.node` in `package.json` (already done).
- Commit a `package-lock.json`. Don't .gitignore it.
- Add a CI build (GitHub Actions: `npm ci && npm run build`) to keep the
  scaffold honest as you change things.

## 6. Documentation

- Check that README's "Quick start (delegated)" section still matches
  reality: shared Microsoft client by default, own-app fallback,
  `DVLOAD_NO_SHARED_CLIENT=1` to opt out.
- Add a CHANGELOG.md with at least the v0.1.0 entry.
- Add a LICENSE file. MIT is the path of least resistance.

## 7. Sanity checks before tagging v1

- `dvload whoami --env <prod-env>` names a client id you recognise.
- `DVLOAD_NO_SHARED_CLIENT=1 dvload login` succeeds against a real tenant,
  proving the own-app fallback path actually works and isn't just dead code.
- Add-in does NOT show the yellow `Dev mode` banner.
- Browser console on add-in load does NOT log `[dvload dev mode]`.
- Manifest GUID is not all zeros.
- Manifest URLs do not contain `localhost`.
- `dvload run-all <plan>.dvplan.json` validates alternate-key links and executes staged dependencies as expected.
- `dvload addin start` successfully launches the local add-in workflow in a dev clone.

If all of the above are true, you're shippable.
