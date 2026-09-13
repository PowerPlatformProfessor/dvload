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
2. **Debug** in the sidebar → **Load Local Tool** → Browse to **this package's
   own folder** (`packages/pptb`) → Load Tool.
3. Pick a connection, open the tool. After a rebuild, close and reopen the
   tool tab.

The loader wants a tool *project*, not the built output: it reads
`<folder>/package.json` and `<folder>/dist/index.html`, and refuses anything
else ("No dist/index.html found" / "No package.json found"). So always pick
the folder that contains `dist/` — never `dist/` itself.

That contract is why the **published** package cannot be loaded as it comes:
the npm package root *is* the dist content. To test what actually shipped,
install it outside this repo — inside it, npm resolves the name to the
workspace and you would be testing your own source — and rebuild the project
shape around it:

```powershell
mkdir pptb-verify; cd pptb-verify; npm init -y; npm i @dvload/pptb
mkdir published\dist
xcopy node_modules\@dvload\pptb published\dist /s /e /y
copy published\dist\package.json published\package.json
```

Then load `pptb-verify\published`. Verified against the published 0.2.0.

## Publish

The published artifact is `dist/`, not this workspace package (this
`package.json` is only the build harness, and it is `private`).

```powershell
npm run build --workspace=@dvload/pptb
cd packages/pptb/dist
npm publish --access public
```

Then submit it to the ToolBox registry via the Tool Submission Form on
powerplatformtoolbox.com (npm package name, display name, description, repo
URL, tags). Before first publish, check the npm scope in
`tool.package.json` — `@dvload` must be a scope you own, or rename.
