# Distribution

The install story, in order of user friction.

## 1. winget (the default in the docs)

```
winget install dvload
```

PATH and upgrades handled, nothing to unblock. This is what the README tells
people to run, so it is the path that has to work.

The three-file manifest lives in `winget/`. Submission is a PR to
[microsoft/winget-pkgs](https://github.com/microsoft/winget-pkgs) copying
those files into `manifests/p/PowerPlatformProfessor/dvload/<version>/`:

```
npm run stamp-release -- 0.2.0          # after the GitHub release exists
winget validate --manifest packaging/winget
wingetcreate submit packaging/winget
```

`InstallerType: portable` — winget shims the bare exe onto PATH itself, so
there is no MSI to author for a file that needs no install-time work.

## 2. Single .exe

`release.yml` builds `dvload.exe` with Node's SEA pipeline on every version
tag and attaches it, plus a `.sha256`, to the release. Users download one
file; no Node install needed.

The add-in UI is **embedded in the exe as SEA assets** (see
`packages/cli/scripts/bundle.mjs`), so `dvload serve` and `dvload gui` work
from the binary alone with nothing extracted beside it. The release smoke
test runs `dvload gui --check-ui` from an otherwise-empty directory to prove
it, because the failure mode is otherwise invisible until a user opens the
task pane.

**Code signing is not optional.** An unsigned exe triggers SmartScreen's
"Windows protected your PC" wall, which loses exactly the semi-technical
Dataverse admins this tool targets. Cheapest route in 2026: Azure Trusted
Signing (~$10/mo). The `Sign` step in `release.yml` is wired for it and
activates as soon as these are set:

| Kind      | Name                                                                      |
| --------- | ------------------------------------------------------------------------- |
| Secrets   | `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`               |
| Variables | `AZURE_SIGNING_ENDPOINT`, `AZURE_SIGNING_ACCOUNT`, `AZURE_SIGNING_PROFILE` |

Until then the build still succeeds and the smoke test emits a warning, so
an unsigned release is a visible choice rather than an accident.

## 3. npm (developers)

```
npm i -g dvload
```

Requires Node 20+. No native modules (keytar was replaced with a
DPAPI-protected file store), so this works on locked-down machines where
node-gyp would fail. Here the UI ships as `build/web/` next to the bundle
rather than embedded — same resolution logic in `serve.ts`, different
branch.

## 4. Scoop (dev-adjacent crowd)

`scoop/dvload.json` is a working manifest; `npm run stamp-release` fills in
the version and hash. `autoupdate` reads the `.sha256` asset from the
release, so later versions update without another manual hash.

## Releasing

1. `npm run verify` locally. `release.yml` runs it again before building —
   a tag is not a promise that the commit was ever green.
2. Bump versions, tag `v<x.y.z>`, push the tag.
3. Wait for `release.yml`, then confirm the run's signature check said
   `Valid`.
4. `npm run stamp-release -- <x.y.z>` — stamps the version and the published
   artifact's hash into the winget and Scoop manifests.
5. Submit the winget PR; update your Scoop bucket.
6. `npm publish --workspace=dvload` to keep the npm path in sync.

## Excel add-in

Do NOT ask users to sideload, and do NOT host the UI on a public origin.

Since the UI moved behind the local sidecar it is served from loopback by
the CLI itself (`dvload serve`), which is what lets it borrow the CLI's auth
instead of needing its own Entra app registration with an `spa` redirect URI
— see the header comment in `packages/cli/src/commands/serve.ts` and
`docs/AUTH-NOTES.md`. The manifest's `SourceLocation` points at
`https://localhost:44321`, so the only thing to distribute is the manifest.

- **Known orgs:** M365 admin center → Integrated apps → upload
  `manifest.prod.xml` (Centralized Deployment). The pane appears in
  everyone's Excel automatically; each user still needs `dvload` installed
  and `dvload serve` running.
- **Strangers:** AppSource via Partner Center, only worth it for public
  distribution.

There is no hosted copy of the pane: the UI ships only inside the exe and
is served from loopback by `dvload serve`. (A GitHub Pages deployment
existed once for previews and was removed — nothing outside the exe needs
the build.)

## Client id

Nothing is baked in at build time. The CLI resolves a public client at
runtime — shared Microsoft client first, then dvload's own registered app —
and `DATAVERSE_LOAD_CLIENT_ID` or `--client-id` overrides both. See
`resolveClientIdChain` in `packages/cli/src/auth.ts`.
