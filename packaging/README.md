# Distribution

The install story, in order of user friction:

## 1. npm (developers)

```
npm i -g dvload
```

Requires Node 20+. No native modules anymore (keytar was replaced with a
DPAPI-protected file store), so this works on locked-down machines where
node-gyp would fail.

## 2. Single .exe (everyone else)

`release.yml` builds `dvload.exe` with Node's SEA pipeline on every version
tag. Users download one file; no Node install needed.

**Code signing is not optional.** An unsigned exe triggers SmartScreen's
"Windows protected your PC" wall, which loses exactly the semi-technical
Dataverse admins this tool targets. Cheapest route in 2026: Azure Trusted
Signing (~$10/mo, integrates with GitHub Actions). Wire it into the marked
step in `release.yml`.

## 3. winget (recommended default in docs)

Once a signed release exists, submit a manifest PR to
[microsoft/winget-pkgs](https://github.com/microsoft/winget-pkgs):

```
wingetcreate new https://github.com/<you>/dvload/releases/download/v<ver>/dvload.exe
```

Users then install with `winget install dvload` — PATH and upgrades handled.

## 4. Scoop (dev-adjacent crowd)

`scoop/dvload.json` is a working manifest template; fill in the version,
URL, and sha256 per release (or automate with a `checkver`/`autoupdate`
block, already stubbed).

## Excel add-in

Do NOT ask users to sideload. Host `packages/addin/dist/` on any static
HTTPS host (Azure Static Web Apps / GitHub Pages), point `manifest.xml` at
it, then either:

- **Known orgs:** M365 admin center → Integrated apps → upload the
  manifest (Centralized Deployment). The add-in appears in everyone's
  Excel automatically — zero per-user steps.
- **Strangers:** AppSource via Partner Center (only worth it for public
  distribution).

Production builds refuse the borrowed dev client id — set
`DATAVERSE_LOAD_CLIENT_ID` to your own Entra app registration first.
