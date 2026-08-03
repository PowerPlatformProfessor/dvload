// Power Query Template (.pqt) CLI commands.
//
//   extract-pqt   read Excel → write a .pqt with the M code and (optionally)
//                 inject FieldsMetadata from a .dvmap.json
//   import-pqt    read a .pqt → write a .dvmap.json from its FieldsMetadata,
//                 useful when migrating from a Dataverse Dataflow

import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import kleur from "kleur";
import { resolveEnv } from "../profiles.js";
import { promptYesNo } from "../prompt.js";
import { openInDefaultApp } from "../open-file.js";
import {
  buildWorkbookWithQueries,
  extractPqtFromXlsx,
  injectMappingIntoPqt,
  mappingFromPqt,
  mappingsFromPqtAll,
  parseMapping,
  parseQueryNames,
  readPqt,
  serializeMapping,
  writePqt,
} from "@dvload/core";

interface ExtractPqtOpts {
  out?: string;
  name?: string;
  description?: string;
  version?: string;
  locale?: string;
  /** Optional .dvmap.json to inject as FieldsMetadata. */
  mapping?: string;
}

export async function extractPqtCommand(
  workbook: string,
  opts: ExtractPqtOpts
): Promise<void> {
  const xlsxPath = path.resolve(workbook);
  const xlsxBytes = await readFile(xlsxPath);
  const archive = await extractPqtFromXlsx(xlsxBytes, {
    name: opts.name,
    description: opts.description,
    version: opts.version,
    locale: opts.locale,
  });

  if (opts.mapping) {
    const mappingRaw = await readFile(path.resolve(opts.mapping), "utf8");
    const mapping = parseMapping(JSON.parse(mappingRaw));
    injectMappingIntoPqt(archive, mapping);
    console.log(
      kleur.gray(
        `Injected ${mapping.columns.length} field mapping(s) for query "${mapping.sourceTable}".`
      )
    );
  }

  const outPath = opts.out
    ? path.resolve(opts.out)
    : path.join(
        path.dirname(xlsxPath),
        path.basename(xlsxPath, path.extname(xlsxPath)) + ".pqt"
      );

  await mkdir(path.dirname(outPath), { recursive: true });
  await writeFile(outPath, await writePqt(archive));

  const queries = Object.keys(archive.mashupMetadata.QueriesMetadata);
  console.log(kleur.green(`Wrote ${outPath}`));
  console.log(kleur.gray(`  ${queries.length} quer${queries.length === 1 ? "y" : "ies"}: ${queries.join(", ")}`));
}

interface ImportPqtOpts {
  query?: string;
  env?: string;
  profile?: string;
  entity?: string;
  out?: string;
  /** Also extract MashupDocument.pq next to the mapping. */
  emitM?: boolean;
  /** Emit one .dvmap.json per query instead of picking one. */
  allQueries?: boolean;
}

export async function importPqtCommand(pqt: string, opts: ImportPqtOpts): Promise<void> {
  const envUrl = await resolveEnv(opts);
  const pqtPath = path.resolve(pqt);
  const archive = await readPqt(await readFile(pqtPath));

  const queries = Object.keys(archive.mashupMetadata.QueriesMetadata);
  if (queries.length === 0) {
    // Synthesize from the M document if MashupMetadata had nothing.
    queries.push(...parseQueryNames(archive.mashupDocument));
  }

  if (opts.allQueries) {
    const mappings = mappingsFromPqtAll(archive, { environmentUrl: envUrl });
    const outDir = opts.out ? path.resolve(opts.out) : path.dirname(pqtPath);
    await mkdir(outDir, { recursive: true });
    let withColumns = 0;
    for (const [name, m] of Object.entries(mappings)) {
      const file = path.join(outDir, `${name.replace(/\W+/g, "-").toLowerCase()}.dvmap.json`);
      await writeFile(file, serializeMapping(m));
      if (m.columns.length > 0) withColumns++;
      console.log(
        kleur.green(`Wrote ${file}`) +
          kleur.gray(` (${m.columns.length} column mapping(s))`)
      );
    }
    console.log(
      kleur.gray(
        `${Object.keys(mappings).length} quer(ies) exported, ${withColumns} with Dataverse field mappings.`
      )
    );
    if (opts.emitM) {
      const mPath = path.join(outDir, `${stemOf(pqtPath)}.pq`);
      await writeFile(mPath, archive.mashupDocument);
      console.log(kleur.green(`Wrote ${mPath}`));
    }
    return;
  }

  const queryName = opts.query ?? pickPrimaryQuery(archive, queries);
  const mapping = mappingFromPqt(archive, queryName, {
    environmentUrl: envUrl,
    targetEntitySet: opts.entity,
  });

  const outPath = opts.out
    ? path.resolve(opts.out)
    : path.join(
        path.dirname(pqtPath),
        `${queryName.replace(/\W+/g, "-").toLowerCase()}.dvmap.json`
      );
  await mkdir(path.dirname(outPath), { recursive: true });
  await writeFile(outPath, serializeMapping(mapping));
  console.log(kleur.green(`Wrote ${outPath}`));
  console.log(
    kleur.gray(`  ${mapping.columns.length} column mapping(s) from query "${queryName}"`)
  );

  if (mapping.description?.includes("DeleteExistingDataOnLoad")) {
    console.log(
      kleur.yellow(
        "Note: the source Dataflow used truncate-and-reload (DeleteExistingDataOnLoad). " +
          "The mapping was written with conflictMode=insert; adjust to upsert + upsertKey " +
          "if you need update semantics."
      )
    );
  }

  if (opts.emitM) {
    const mPath = outPath.replace(/\.dvmap\.json$/, ".pq");
    await writeFile(mPath, archive.mashupDocument);
    console.log(kleur.green(`Wrote ${mPath}`));
  }

  if (mapping.columns.length === 0) {
    console.log(
      kleur.yellow(
        "Warning: no FieldsMetadata in the .pqt for that query, so the output .dvmap.json " +
          "has no columns. Only Dataverse Dataflow exports record which source column feeds " +
          "which attribute; a .pqt saved from Power Query Online or extracted from Excel " +
          "carries the M code alone. Map the columns in the add-in (or by hand) and use " +
          "extract-pqt --mapping to write them back into a .pqt."
      )
    );
  }
}

/**
 * EXPERIMENTAL: turn a .pqt into an .xlsx whose Power Query editor contains
 * every query, ready for "Load To…". The DataMashup part is synthesized per
 * MS-QDEFF; if Excel refuses the file, fall back to `import-pqt --emit-m`
 * and paste the M into a Blank Query's Advanced Editor.
 */
export async function pqtToXlsxCommand(
  pqt: string,
  opts: { out?: string; open?: boolean }
): Promise<void> {
  const pqtPath = path.resolve(pqt);
  const archive = await readPqt(await readFile(pqtPath));
  const outPath = opts.out
    ? path.resolve(opts.out)
    : path.join(path.dirname(pqtPath), `${stemOf(pqtPath)}.xlsx`);

  const bytes = await buildWorkbookWithQueries(archive);
  await mkdir(path.dirname(outPath), { recursive: true });
  await writeFile(outPath, bytes);

  const queries = Object.keys(archive.mashupMetadata.QueriesMetadata);
  const names = queries.length > 0 ? queries : parseQueryNames(archive.mashupDocument);
  console.log(kleur.green(`Wrote ${outPath}`));
  console.log(
    kleur.gray(
      `  ${names.length} quer(ies) embedded: ${names.join(", ")}\n` +
        `  Open in Excel → Data → Queries & Connections to see them; use "Load To…" per query.\n` +
        `  (Experimental QDEFF writer — if Excel complains, use import-pqt --emit-m and paste the M.)`
    )
  );

  // --open forces, --no-open suppresses, and bare interactive runs ask. On a
  // non-TTY (CI, cron, piped) promptYesNo returns the default without reading
  // stdin, so an unattended run never blocks here.
  const shouldOpen = opts.open ?? (await promptYesNo("Open it in Excel now?", false));
  if (!shouldOpen) return;

  try {
    await openInDefaultApp(outPath);
    console.log(kleur.gray("  Handed off to Excel."));
  } catch (e) {
    // The file is already written; failing to launch is a nuisance, not an error.
    console.log(
      kleur.yellow(`  Couldn't open it automatically (${(e as Error).message}). Open it yourself:`)
    );
    console.log(kleur.gray(`  ${outPath}`));
  }
}

function stemOf(file: string): string {
  return path.basename(file, path.extname(file));
}

function pickPrimaryQuery(
  archive: { mashupMetadata: { QueriesMetadata: Record<string, { LoadEnabled?: boolean }> } },
  queries: string[]
): string {
  // Prefer a query with LoadEnabled=true (the one that actually writes data).
  const loadable = queries.find(
    (n) => archive.mashupMetadata.QueriesMetadata[n]?.LoadEnabled
  );
  if (loadable) return loadable;
  if (queries.length === 1) return queries[0];
  throw new Error(
    `Multiple queries in .pqt: ${queries.join(", ")}. Pass --query <name> to pick one.`
  );
}
