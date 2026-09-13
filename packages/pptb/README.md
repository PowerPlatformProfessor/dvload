# dvload — Excel to Dataverse (Power Platform ToolBox tool)

The [dvload](https://github.com/PowerPlatformProfessor/dvload) mapping UI as a
[Power Platform ToolBox](https://www.powerplatformtoolbox.com/) tool: pick an
`.xlsx` / `.csv` / `.tsv`, map its columns to a Dataverse table, and run the
import against the ToolBox's active connection — no sign-in, no sidecar, no
install beyond the tool itself.

It is the same interface and the same engine as the Excel add-in and the
browser UI (`dvload gui`). Mappings are portable both ways: save a
`.dvmap.json` here and run it nightly with the dvload CLI, or open one built
elsewhere.

## What's different inside the ToolBox

The ToolBox deliberately never hands tools an access token: every Dataverse
call crosses its `dataverseAPI` bridge. dvload's engine runs on a gateway
adapter over that bridge (`packages/addin/src/pptb/pptb-client.ts`), which
changes a few things:

- **No account/environment steps.** The ToolBox connection decides both; the
  bar at the bottom of the pane shows where records will land. Switching the
  connection reloads the tool.
- **`$batch` is emulated.** Operations run individually through the bridge
  (the engine's per-row accounting semantics are preserved — core already
  isolates every operation in its own changeset). Expect large loads to be
  slower than the add-in/CLI, which use real `$batch`.
- **Upserts are emulated** (probe, then update-or-create), so they are not
  atomic: a record created by someone else between probe and write surfaces
  as a per-row error rather than being overwritten.
- **Not supported here** (use the Excel add-in or the CLI):
  - *Bypass plugins / Run as user* — per-request headers can't cross the bridge.
  - *Alternate-key creation* in "create table from source" — the table and
    columns are created; add the key in Power Apps (Table → Keys).
- **Sync mode caveat:** removing missing records lists the whole target
  table. Paging past 5,000 rows depends on the bridge passing
  `@odata.nextLink` through `queryData`; a truncated listing makes sync
  remove *fewer* records than it should (never more). Verify against your
  ToolBox version before relying on sync for tables that size.
- Reading the *open workbook* and Power Query extraction are Excel-only, as
  in the browser UI. "Add files…" covers both.

Dataflow import works fully: the conversion runs in-page through the bridge
instead of through `dvload serve`.

## Build

```powershell
npm install            # from the repo root
npm run build --workspace=@dvload/pptb
```

`dist/` is the complete tool package: `index.html`, the bundle, `icon.svg`,
and a `package.json` generated from `tool.package.json` (the PPTB manifest —
displayName, icon, `features.minAPI`).

## Try it locally

1. Power Platform ToolBox → **Settings** → enable **Show Debug Menu**.
2. **Debug** in the sidebar → **Load Local Tool** → Browse to
   **`packages/pptb/publish`** → Load Tool.
3. Pick a connection, open the tool. After a rebuild, close and reopen the
   tool tab.

`publish/` is what the build assembles and what `npm publish` uploads, so the
tool you debug is byte-for-byte the tool that ships.

Both the loader and the registry want the same shape — a tool *project*, not
the built output:

```
package.json        the PPTB manifest (from tool.package.json)
dist/index.html     the entry named by `main`, beside the bundle and icons/
```

Handing either of them `dist/` itself fails: the loader says "No
dist/index.html found" (it appends `dist/` to whatever folder you pick), and
submission fails `structure_validation` with "dist folder is required but not
found in the package". `main` and `icon` are resolved relative to `dist/`.

To load a *published* version rather than the working tree, install it
somewhere outside this repo — inside it, npm resolves the name to the
workspace and you would be testing your own source — and point the loader at
the package folder, which already has the right shape:

```powershell
mkdir pptb-verify; cd pptb-verify; npm init -y; npm i @dvload/pptb
```

Then load `pptb-verify\node_modules\@dvload\pptb` (0.2.1 and later).

## Publish

The published artifact is `publish/`, assembled by the build from `dist/` plus
`tool.package.json`. This workspace's own `package.json` is only the build
harness — it is `private`, and its dependencies (`@dvload/core`,
`@dvload/addin`) are never published, so publishing it directly would give
consumers a package npm cannot install.

```powershell
npm run build --workspace=@dvload/pptb
cd packages/pptb/publish
npx @pptb/validate
npm publish --access public
```

npm versions are immutable, so bump `version` in `tool.package.json` for every
publish. Before the first one, check the scope — `@dvload` must be one you own.

Then submit it at <https://www.powerplatformtoolbox.com/submit-tool> (login
required): npm package name plus up to three tags. Automated checks run first
— including the structure validation described above — then a human review,
quoted at 48–72 hours.
