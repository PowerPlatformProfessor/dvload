# Pre-release checklist

Things that must change before this tool is shared with anyone outside your
own machine. Track these like blockers — at minimum, the auth and manifest
items are non-negotiable.

## How to know you're not ready yet

- Add-in shows a yellow banner at the top: `Dev mode — signing in as
  "Microsoft Dynamics CRM"`.
- Browser dev console logs `[dvload dev mode]` on add-in start.

Either one means the add-in build is pointed at a borrowed Microsoft
client id, which cannot work in a browser (see below). Fix that first.

**The CLI is a different case and is not a blocker.** By design it signs
in through the shared Microsoft Dataverse client
(`51f81489-12ee-4a9e-aaae-a2591f45987d`) so users never hit an admin
consent wall and falls back to dvload's own
app if that client is blocked. `dvload login` prints a one-line note
saying which identity the consent screen will show. That's informational,
not a warning.

## 1. Register your own Entra ID app(s)

You need at least one app registration, possibly two depending on whether
the CLI will support unattended (app-only) auth out of the box for users.

**Public client (delegated flow) — required for the add-in, fallback for
the CLI.**

The add-in has no alternative here. A browser auth-code flow needs a
redirect URI registered on the app registration, and Entra only returns
`Access-Control-Allow-Origin` from the token endpoint when the caller's
origin matches a redirect URI of type `spa`. You can't add either to a
Microsoft-owned app, and Entra explicitly rejects `spa` redirect URIs in
non-SPA flows — so device code from the taskpane doesn't rescue you
either. The only way around it would be to stop authenticating in the
browser entirely and broker tokens through a local dvload process.

- Microsoft Entra admin center → App registrations → New registration
- Name: `<your tool name>` (whatever appears on the consent screen)
- Supported account types: *Accounts in any organizational directory
  (multi-tenant)*
- Redirect URIs:
  - Single-page application: every URL the add-in will be hosted at
    (`https://yourdomain.example/taskpane.html`, plus
    `https://localhost:3000/taskpane.html` if you keep dev sideloading)
  - Public client / native: leave the default `http://localhost`
- Authentication → Allow public client flows: **Yes**
- API permissions → Add → Dynamics CRM → user_impersonation (delegated)
- Branding & properties → set your publisher name, support URL, privacy
  statement URL, terms of service URL. These show on the consent screen.

Then:

- Set `DATAVERSE_LOAD_CLIENT_ID` at webpack build time so the add-in
  bundle carries your app id (`packages/addin/webpack.config.js` fails a
  production build if it doesn't).
- Optionally update `DVLOAD_CLIENT_ID` in `packages/cli/src/auth.ts` — this
  is only the CLI's *fallback*, used when the shared Microsoft client is
  blocked. Leaving it alone is fine.

**Confidential client (app-only flow) — optional.** Only needed if you
want `app-login` to work without users having to register their own
Entra app first. Most tools omit this from the default install and have
power users register their own; that's fine.

## 2. Office add-in manifest

`packages/addin/manifest.xml` is currently a dev manifest. Before
sideloading anywhere outside your machine, and definitely before
AppSource submission:

- Replace `<Id>00000000-0000-0000-0000-000000000000</Id>` with a real
  GUID. Generate with `[guid]::NewGuid()` in PowerShell or `uuidgen`.
- Replace every `https://localhost:3000/...` URL with your production
  add-in host.
- Update `<ProviderName>`, `<DisplayName>`, `<Description>`,
  `<SupportUrl>`, and the resources at the bottom (`GetStarted.Title`,
  `GetStarted.Description`, `GetStarted.LearnMoreUrl`, etc.).
- Add real 16x16, 32x32, and 80x80 icons under `packages/addin/assets/`
  and verify they load.
- Validate the manifest: `npm --workspace=@dvload/addin run validate`.

## 3. AppSource submission (only if listing publicly)

- Add a privacy policy and terms of service hosted somewhere durable.
- Test the add-in in Excel desktop (Win + Mac), Excel for the web, and
  Excel on iPad if claiming mobile support.
- Submit through Partner Center → Office Add-ins.
- Expect 5–10 business days for review; reviewers do test data ops, so
  make sure a tester Dataverse environment is reachable.

## 4. CLI distribution

- Bundle as a single executable: `npm run bundle --workspace=@dvload/cli`,
  then the Node SEA steps in `.github/workflows/release.yml` (keytar is
  gone, so the bundle has no native addons and SEA works cleanly).
- Code-sign the .exe so Windows SmartScreen doesn't yell at users.
  EV certs are nice but not strictly required for personal-scale distribution.
- Publish to GitHub Releases with checksums (SHA-256) in the release notes.
- Document the install path (suggested: drop into `%USERPROFILE%\bin`
  which most users already have on PATH, or use `winget` if you want to
  set that up).

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
