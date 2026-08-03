// Pull Power Platform dataflows straight out of a connected environment.
//
//   dataflows          list what's there
//   import-dataflow    turn one into an .xlsx of its queries and/or
//                      .dvmap.json files for its Dataverse mappings
//
// This is the live-environment counterpart to `import-pqt`. The difference
// matters: a .pqt file carries the M code alone, so `import-pqt` can only
// ever emit mappings with no columns. Dataverse still holds the destination
// config, so this path produces mappings that are actually populated. See
// core/src/dataflow.ts.

import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import kleur from "kleur";
import {
  DataverseClient,
  buildWorkbookWithQueries,
  createMetadataResolver,
  findDataflowByName,
  getDataflow,
  listDataflows,
  mappingsFromDataflow,
  serializeMapping,
  validateMapping,
  type DataflowSummary,
} from "@dvload/core";
import { getTokenProvider } from "../auth.js";
import { resolveEnv } from "../profiles.js";
import { openInDefaultApp } from "../open-file.js";

interface EnvOpts {
  env?: string;
  profile?: string;
  user?: boolean;
  nonInteractive?: boolean;
}

async function connect(opts: EnvOpts): Promise<{ client: DataverseClient; envUrl: string }> {
  const envUrl = await resolveEnv(opts);
  const nonInteractive = opts.nonInteractive ?? !process.stdout.isTTY;
  const getToken = await getTokenProvider({
    environmentUrl: envUrl,
    forceUser: opts.user,
    silentOnly: nonInteractive,
  });
  return { client: new DataverseClient({ environmentUrl: envUrl, getToken }), envUrl };
}

/* -------------------------------------------------------------------------- */
/* dvload dataflows                                                            */
/* -------------------------------------------------------------------------- */

interface ListOpts extends EnvOpts {
  drafts?: boolean;
  json?: boolean;
}

export async function dataflowsCommand(opts: ListOpts): Promise<void> {
  const { client } = await connect(opts);
  const flows = await listDataflows(client, { includeDrafts: opts.drafts });

  if (opts.json) {
    console.log(JSON.stringify(flows, null, 2));
    return;
  }

  if (flows.length === 0) {
    console.log(kleur.yellow("No dataflows in this environment."));
    return;
  }

  for (const f of flows) {
    const state = f.state === "Active" ? "" : kleur.yellow(` [${f.state.toLowerCase()}]`);
    console.log(kleur.green(f.name) + state);
    console.log(kleur.gray(`  ${f.id}`));
    console.log(kleur.gray(`  ${describeTargets(f)}`));
  }
  if (!opts.drafts) {
    console.log(
      kleur.gray(
        `\n${flows.length} published dataflow(s). Editing drafts are hidden; pass --drafts to see them.`
      )
    );
  }
}

function describeTargets(f: DataflowSummary): string {
  if (f.loadTargets.length === 0) {
    return `${f.queryNames.length} quer(ies), none loading to Dataverse`;
  }
  const targets = f.loadTargets
    .map((t) => `${t.queryName} → ${t.entityName} (${t.fieldCount} fields)`)
    .join(", ");
  const staging = f.queryNames.length - f.loadTargets.length;
  return staging > 0 ? `${targets}; ${staging} staging quer(ies)` : targets;
}

/* -------------------------------------------------------------------------- */
/* dvload import-dataflow                                                      */
/* -------------------------------------------------------------------------- */

interface ImportOpts extends EnvOpts {
  out?: string;
  /** Default: both, matching the task pane's two checkboxes. */
  xlsx?: boolean;
  mapping?: boolean;
  drafts?: boolean;
  open?: boolean;
}

export async function importDataflowCommand(
  nameOrId: string,
  opts: ImportOpts
): Promise<void> {
  const { client, envUrl } = await connect(opts);

  // Both flags default on. Commander sets them to false only when the user
  // passes --no-xlsx / --no-mapping, so an absent flag must not be read as
  // "off" — that would make the bare command produce nothing.
  const wantXlsx = opts.xlsx !== false;
  const wantMapping = opts.mapping !== false;
  if (!wantXlsx && !wantMapping) {
    throw new Error("Nothing to do: --no-xlsx and --no-mapping were both passed.");
  }

  const id = isGuid(nameOrId)
    ? nameOrId
    : (await findDataflowByName(client, nameOrId, { includeDrafts: opts.drafts })).id;
  const detail = await getDataflow(client, id);

  const outDir = path.resolve(opts.out ?? ".");
  await mkdir(outDir, { recursive: true });
  const stem = slug(detail.name);

  if (wantXlsx) {
    const xlsxPath = path.join(outDir, `${stem}.xlsx`);
    await writeFile(xlsxPath, await buildWorkbookWithQueries(detail.archive));
    console.log(kleur.green(`Wrote ${xlsxPath}`));
    console.log(
      kleur.gray(
        `  ${detail.queryNames.length} quer(ies): ${detail.queryNames.join(", ")}\n` +
          `  Data → Queries & Connections in Excel; "Load To…" per query.\n` +
          `  The M keeps the dataflow's SharePoint URLs, so Excel will ask for those credentials.`
      )
    );
    if (opts.open) await openQuietly(xlsxPath);
  }

  if (wantMapping) {
    // Metadata is only read when mappings are requested: resolving alternate
    // keys and lookup targets costs several round trips, and the workbook
    // path needs none of it.
    const mappings = await mappingsFromDataflow(detail, {
      environmentUrl: envUrl,
      resolver: createMetadataResolver(client),
    });

    const names = Object.keys(mappings);
    if (names.length === 0) {
      console.log(
        kleur.yellow(
          `No query in "${detail.name}" loads to Dataverse, so there are no mappings to write.`
        )
      );
      return;
    }

    for (const [queryName, mapping] of Object.entries(mappings)) {
      const file = path.join(outDir, `${slug(queryName)}.dvmap.json`);
      await writeFile(file, serializeMapping(mapping));
      console.log(
        kleur.green(`Wrote ${file}`) +
          kleur.gray(
            ` (${mapping.columns.length} columns, ${mapping.conflictMode}` +
              `${mapping.upsertKey ? ` on ${mapping.upsertKey.join("+")}` : ""})`
          )
      );

      // Surfaced now rather than at run time: an unresolved lookup or a
      // missing entity set is something to fix in the task pane before the
      // first load, not something to discover mid-import.
      const problems = validateMapping(mapping);
      for (const p of problems) console.log(kleur.yellow(`  ! ${p}`));
      if (mapping.description) console.log(kleur.gray(`  ${mapping.description}`));
    }
  }
}

function isGuid(s: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s);
}

function slug(s: string): string {
  return s.replace(/\W+/g, "-").replace(/^-+|-+$/g, "").toLowerCase() || "dataflow";
}

async function openQuietly(file: string): Promise<void> {
  try {
    await openInDefaultApp(file);
  } catch (e) {
    // The file is already written; failing to launch Excel is a nuisance.
    console.log(kleur.yellow(`  Couldn't open it automatically (${(e as Error).message}).`));
  }
}
