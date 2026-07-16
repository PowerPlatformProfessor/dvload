# Pre-release checklist

Things that must change before this tool is shared with anyone outside your
own machine. Track these like blockers — at minimum, the auth and manifest
items are non-negotiable.

## How to know you're not ready yet

You'll see one or more of these:

- CLI prints `[dev mode] Using "Microsoft PowerApps" public client …` when
  you run `login` or `whoami`.
- Add-in shows a yellow banner at the top: `Dev mode — signing in as
  "Microsoft PowerApps"`.
- Browser dev console logs `[dvload dev mode]` on add-in start.

All three are the same signal: you're still using a borrowed Microsoft
public client id. Fix that first.

## 1. Register your own Entra ID app(s)

You need at least one app registration, possibly two depending on whether
the CLI will support unattended (app-only) auth out of the box for users.

**Public client (delegated flow) — required.** Used by the add-in and by
`dvload login`.

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

- Replace the well-known fallback in `packages/cli/src/auth.ts`
  (`DEFAULT_PUBLIC_CLIENT_ID`) with your new app id.
- Replace the same constant in `packages/addin/src/auth.ts` (`CLIENT_ID`).
- Or: keep the env-var override and ship the app id via build-time
  injection (`process.env.DATAVERSE_LOAD_CLIENT_ID`) so it's not hardcoded
  in source.

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

- Update `README.md` to remove the "if you don't want to register your own
  app, use this well-known client id" shortcut. That paragraph is the
  thing telling people they're allowed to do something this checklist
  forbids.
- Add a CHANGELOG.md with at least the v0.1.0 entry.
- Add a LICENSE file. MIT is the path of least resistance.

## 7. Sanity checks before tagging v1

- `dvload whoami --env <prod-env>` does NOT print `[dev mode]`.
- Add-in does NOT show the yellow `Dev mode` banner.
- Browser console on add-in load does NOT log `[dvload dev mode]`.
- Manifest GUID is not all zeros.
- Manifest URLs do not contain `localhost`.
- README mentions your registered app, not the PowerApps fallback.

If all of the above are true, you're shippable.
