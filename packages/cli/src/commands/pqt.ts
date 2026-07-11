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
import {
  extractPqtFromXlsx,
  injectMappingIntoPqt,
  mappingFromPqt,
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
        "Warning: no FieldsMetadata in the .pqt for that query. The output .dvmap.json " +
          "has no columns. This usually means the .pqt was extracted from Excel (Excel " +
          "doesn't know about Dataverse mappings) rather than from a Dataverse Dataflow."
      )
    );
  }
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
