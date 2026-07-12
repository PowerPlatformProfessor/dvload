// Run several mappings in declared order — needed when lookups create
// dependencies between tables (accounts before contacts, etc.).
//
// Manifest shape (.json):
// {
//   "stopOnError": true,            // default true
//   "runs": [
//     { "mapping": "./accounts.dvmap.json", "workbook": "./data.xlsx", "refresh": true },
//     { "mapping": "./contacts.dvmap.json", "workbook": "./data.xlsx" }
//   ]
// }
// Relative paths resolve against the manifest file's directory.

import { readFile } from "node:fs/promises";
import path from "node:path";
import kleur from "kleur";
import { executeRun, type RunOpts } from "./run.js";

interface ManifestRun {
  mapping: string;
  workbook: string;
  refresh?: boolean;
}

interface Manifest {
  stopOnError?: boolean;
  runs: ManifestRun[];
}

export interface RunAllOpts {
  dryRun?: boolean;
  user?: boolean;
  notifyUrl?: string;
  failedRows?: boolean;
}

export async function runAllCommand(manifestPath: string, opts: RunAllOpts): Promise<void> {
  const file = path.resolve(manifestPath);
  const dir = path.dirname(file);
  const manifest = parseManifest(JSON.parse(await readFile(file, "utf8")));
  const stopOnError = manifest.stopOnError ?? true;

  let failedRuns = 0;
  for (let i = 0; i < manifest.runs.length; i++) {
    const r = manifest.runs[i];
    const mappingPath = path.resolve(dir, r.mapping);
    console.log("");
    console.log(kleur.bold(`[${i + 1}/${manifest.runs.length}] ${r.mapping}`));

    const runOpts: RunOpts = {
      workbook: path.resolve(dir, r.workbook),
      refresh: r.refresh,
      dryRun: opts.dryRun,
      user: opts.user,
      notifyUrl: opts.notifyUrl,
      failedRows: opts.failedRows,
    };

    let ok = false;
    try {
      const result = await executeRun(mappingPath, runOpts);
      ok = result !== null && result.failed === 0;
    } catch (e) {
      console.error(kleur.red((e as Error).message));
    }

    if (!ok) {
      failedRuns++;
      if (stopOnError) {
        console.error(
          kleur.red(
            `Stopping: run ${i + 1} failed and stopOnError is set. ` +
              `${manifest.runs.length - i - 1} run(s) not attempted.`
          )
        );
        break;
      }
    }
  }

  if (failedRuns > 0) process.exitCode = 1;
  console.log("");
  console.log(
    failedRuns === 0
      ? kleur.green(`All ${manifest.runs.length} runs completed successfully.`)
      : kleur.red(`${failedRuns} run(s) had failures.`)
  );
}

function parseManifest(input: unknown): Manifest {
  if (typeof input !== "object" || input === null || Array.isArray(input)) {
    throw new Error("Manifest must be a JSON object with a runs[] array.");
  }
  const m = input as Record<string, unknown>;
  if (!Array.isArray(m.runs) || m.runs.length === 0) {
    throw new Error("Manifest needs a non-empty runs[] array.");
  }
  const runs: ManifestRun[] = m.runs.map((r, i) => {
    if (typeof r !== "object" || r === null) throw new Error(`runs[${i}] must be an object`);
    const rr = r as Record<string, unknown>;
    if (typeof rr.mapping !== "string" || typeof rr.workbook !== "string") {
      throw new Error(`runs[${i}] needs "mapping" and "workbook" string paths`);
    }
    return {
      mapping: rr.mapping,
      workbook: rr.workbook,
      refresh: rr.refresh === true,
    };
  });
  return { stopOnError: m.stopOnError !== false, runs };
}
