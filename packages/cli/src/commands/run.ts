import { mkdir, readFile, writeFile, unlink, stat } from "node:fs/promises";
import path from "node:path";
import kleur from "kleur";
import {
  parseMapping,
  validateMapping,
  mappingWarnings,
  DataverseClient,
  loadRows,
  readTableFromFile,
  writeRowsToFile,
  type LoadResult,
  type Mapping,
  type RequestLogEntry,
  type RowSuccess,
} from "@dvload/core";
import { getTokenProvider } from "../auth.js";
import { refreshWorkbook } from "../refresh.js";

export interface RunOpts {
  workbook: string;
  refresh?: boolean;
  dryRun?: boolean;
  user?: boolean;
  maxErrors?: number;
  concurrency?: number;
  /** Resume from the last checkpoint if one exists for this mapping+workbook. */
  resume?: boolean;
  /** POST a {text} summary here after the run (overrides mapping.notifyUrl). */
  notifyUrl?: string;
  /** Set false (--no-failed-rows) to suppress the failed-rows .xlsx. */
  failedRows?: boolean;
  /**
   * Never fall back to an interactive device-code prompt. Defaults to true
   * when stdout is not a TTY (scheduled/piped runs) so an expired token
   * fails fast instead of hanging on a prompt nobody can see.
   */
  nonInteractive?: boolean;
  /** Emit the LoadResult as JSON on stdout; suppress pretty output. */
  json?: boolean;
}

export async function runCommand(mappingPath: string, opts: RunOpts): Promise<void> {
  const result = await executeRun(mappingPath, opts);
  if (result && result.failed > 0) process.exitCode = 1;
}

/**
 * The actual run pipeline, shared by `run` and `run-all`.
 * Returns null when the mapping failed schema validation (exitCode set).
 */
export async function executeRun(mappingPath: string, opts: RunOpts): Promise<LoadResult | null> {
  const quiet = opts.json === true;
  const log = quiet ? () => {} : console.log.bind(console);
  const write = quiet ? () => {} : process.stdout.write.bind(process.stdout);
  const mappingFile = path.resolve(mappingPath);
  const raw = await readFile(mappingFile, "utf8");
  const mapping = parseMapping(JSON.parse(raw));

  const schemaErrs = validateMapping(mapping);
  if (schemaErrs.length > 0) {
    console.error(kleur.red("Mapping has schema errors:"));
    for (const e of schemaErrs) console.error("  - " + e);
    process.exitCode = 2;
    return null;
  }

  for (const w of mappingWarnings(mapping)) log(kleur.yellow("! " + w));

  if (typeof opts.maxErrors === "number") mapping.maxErrors = opts.maxErrors;
  if (typeof opts.concurrency === "number") mapping.concurrency = opts.concurrency;

  const workbookPath = path.resolve(opts.workbook);
  const logDir = path.resolve(path.dirname(mappingFile), mapping.logDir);
  const checkpointFile = path.join(
    logDir,
    `${stem(mappingFile)}.checkpoint.json`
  );

  // Resume: read the checkpoint and verify the workbook hasn't changed.
  let startOffset = 0;
  if (opts.resume) {
    const cp = await readCheckpoint(checkpointFile, workbookPath);
    if (cp !== null) {
      startOffset = cp;
      log(kleur.yellow(`Resuming from row offset ${startOffset} (checkpoint found).`));
      if (opts.refresh) {
        log(
          kleur.yellow("  --refresh skipped: refreshing would change the data under a resume.")
        );
        opts = { ...opts, refresh: false };
      }
    } else {
      log(kleur.gray("No usable checkpoint — starting from the beginning."));
    }
  }

  if (opts.refresh) {
    log(kleur.gray("Refreshing Power Query in Excel..."));
    await refreshWorkbook({ workbook: workbookPath });
  }

  log(kleur.gray(`Reading table "${mapping.sourceTable}" from ${workbookPath}`));
  const { headers, rows } = await readTableFromFile(workbookPath, {
    tableName: mapping.sourceTable,
    sheetName: mapping.sourceSheet,
  });
  log(kleur.gray(`  ${rows.length} row(s) loaded.`));

  // Unattended runs (scheduled tasks, pipes, CI) must never block on a
  // device-code prompt written into a log nobody reads.
  const nonInteractive = opts.nonInteractive ?? !process.stdout.isTTY;
  const getToken = await getTokenProvider({
    environmentUrl: mapping.environmentUrl,
    forceUser: opts.user,
    silentOnly: nonInteractive,
  });

  const requestLog: RequestLogEntry[] = [];
  const successLog: RowSuccess[] = [];

  const client = new DataverseClient({
    environmentUrl: mapping.environmentUrl,
    getToken,
    onRequest: (entry) => requestLog.push(entry),
    retry: {
      onRetry: (info) => {
        log(
          kleur.yellow(
            `\n  [retry ${info.attempt}] ${info.status} ${info.url} — ` +
              `waiting ${(info.delayMs / 1000).toFixed(1)}s (${info.source})`
          )
        );
      },
    },
  });

  await mkdir(logDir, { recursive: true });
  const wbStat = await stat(workbookPath);
  let checkpointChain: Promise<void> = Promise.resolve();

  const result = await loadRows({
    mapping,
    rows,
    client,
    dryRun: opts.dryRun,
    startOffset,
    onProgress: (e) => {
      if (e.type === "row-success") {
        successLog.push(e.success);
      } else if (e.type === "checkpoint") {
        if (!opts.dryRun) {
          const payload = JSON.stringify({
            offset: e.offset,
            workbookSize: wbStat.size,
            workbookMtimeMs: wbStat.mtimeMs,
          });
          // Serialize writes so a fast run can't interleave them.
          checkpointChain = checkpointChain
            .then(() => writeFile(checkpointFile, payload, "utf8"))
            .catch(() => {});
        }
      } else if (e.type === "sync") {
        write(
          `\r  sync: ${e.removed}/${e.toRemove} removed (${e.checked} target records checked)     `
        );
      } else if (e.type === "batch") {
        const pct = Math.round((e.processed / Math.max(1, e.total)) * 100);
        write(
          `\r  ${pct}%  created=${e.created}  updated=${e.updated}  ` +
            `skipped=${e.skipped}  failed=${e.failed}  ${e.processed}/${e.total}     `
        );
      } else if (e.type === "done") {
        write("\n");
      }
    },
  });

  // The run completed (even if rows failed) — the checkpoint is obsolete.
  await checkpointChain;
  await unlink(checkpointFile).catch(() => {});

  await writeRunLog(mappingFile, mapping.logDir, workbookPath, requestLog, successLog, result);

  // Failed-rows re-run file: same columns as the source table, one row per
  // failure, so you can fix values and run just this file.
  const failedSourceRows = result.errors.filter((e) => e.rowIndex >= 0).map((e) => e.sourceRow);
  let failedFile: string | undefined;
  if (failedSourceRows.length > 0 && opts.failedRows !== false && !opts.dryRun) {
    failedFile = path.join(logDir, `failed_${stem(workbookPath)}_${timestamp(result.startedAt)}.xlsx`);
    try {
      await writeRowsToFile(failedFile, headers, failedSourceRows, mapping.sourceTable);
    } catch (e) {
      console.error(kleur.yellow(`Could not write failed-rows file: ${(e as Error).message}`));
      failedFile = undefined;
    }
  }

  if (quiet) {
    // Machine-readable result for pipelines/monitoring.
    console.log(JSON.stringify({ ...result, failedRowsFile: failedFile ?? null }));
  } else {
    printSummary(mapping, result, failedFile);
  }

  const notifyUrl = opts.notifyUrl ?? mapping.notifyUrl;
  if (notifyUrl && !opts.dryRun) {
    await notify(notifyUrl, mapping, result);
  }

  return result;
}

/* -------------------------------------------------------------------------- */
/* Helpers                                                                     */
/* -------------------------------------------------------------------------- */

function stem(file: string): string {
  return path.basename(file, path.extname(file)).replace(/\.dvmap$/, "").replace(/[^\w.-]/g, "_");
}

function timestamp(iso: string): string {
  return iso.slice(0, 19).replace(/:/g, "-").replace("T", "_");
}

async function readCheckpoint(checkpointFile: string, workbookPath: string): Promise<number | null> {
  try {
    const cp = JSON.parse(await readFile(checkpointFile, "utf8")) as {
      offset?: number;
      workbookSize?: number;
      workbookMtimeMs?: number;
    };
    if (typeof cp.offset !== "number" || cp.offset <= 0) return null;
    const s = await stat(workbookPath);
    if (cp.workbookSize !== s.size || cp.workbookMtimeMs !== s.mtimeMs) {
      console.log(kleur.yellow("Checkpoint ignored: the workbook changed since the interrupted run."));
      return null;
    }
    return cp.offset;
  } catch {
    return null;
  }
}

function printSummary(mapping: Mapping, result: LoadResult, failedFile?: string): void {
  console.log("");
  console.log(kleur.bold("Run summary"));
  console.log(`  total:     ${result.total}`);
  console.log(`  created:   ${kleur.green(String(result.created))}`);
  console.log(`  updated:   ${kleur.green(String(result.updated))}`);
  console.log(`  skipped:   ${result.skipped}${result.unchanged > 0 ? kleur.gray(` (${result.unchanged} unchanged)`) : ""}`);
  if (mapping.conflictMode === "sync") {
    console.log(`  removed:   ${result.removed} (${mapping.syncAction ?? "deactivate"})`);
  }
  console.log(`  failed:    ${result.failed > 0 ? kleur.red(String(result.failed)) : "0"}`);
  console.log(`  duration:  ${duration(result.startedAt, result.finishedAt)}`);
  if (result.errors.length > 0) {
    console.log("");
    console.log(kleur.red("First few errors:"));
    for (const e of result.errors.slice(0, 5)) {
      console.log(`  ${e.rowIndex >= 0 ? `row ${e.rowIndex}` : "sync"}: ${e.message}`);
    }
    if (failedFile) {
      console.log("");
      console.log(kleur.yellow(`Failed rows written to: ${failedFile}`));
      console.log(kleur.gray("  Fix the values and re-run against that file."));
    }
  }
}

async function notify(url: string, mapping: Mapping, result: LoadResult): Promise<void> {
  const status = result.failed > 0 ? "⚠ completed with errors" : "✓ succeeded";
  const text =
    `dvload: "${mapping.name}" → ${mapping.targetEntitySet} ${status}. ` +
    `${result.created} created, ${result.updated} updated, ${result.skipped} skipped` +
    (result.unchanged > 0 ? ` (${result.unchanged} unchanged)` : "") +
    (result.removed > 0 ? `, ${result.removed} removed` : "") +
    `, ${result.failed} failed of ${result.total} rows in ${duration(result.startedAt, result.finishedAt)}.`;
  try {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text }),
    });
    if (!res.ok) {
      console.error(kleur.yellow(`Notification webhook returned ${res.status}.`));
    }
  } catch (e) {
    console.error(kleur.yellow(`Notification failed: ${(e as Error).message}`));
  }
}

async function writeRunLog(
  mappingFile: string,
  logDir: string,
  workbookPath: string,
  requestEntries: RequestLogEntry[],
  successEntries: RowSuccess[],
  result: LoadResult
): Promise<void> {
  const dir = path.resolve(path.dirname(mappingFile), logDir);
  await mkdir(dir, { recursive: true });
  const excelStem = stem(workbookPath);
  const startedAt = new Date(result.startedAt);
  const datePart = startedAt.toISOString().slice(0, 10);
  const timePart = startedAt.toISOString().slice(11, 19).replace(/:/g, "-");
  const file = path.join(dir, `${excelStem}_${datePart}_${timePart}.jsonl`);
  const lines: string[] = [];
  lines.push(JSON.stringify({ event: "summary", ...result, errors: undefined }));
  for (const r of requestEntries) lines.push(JSON.stringify(r));
  for (const s of successEntries) lines.push(JSON.stringify({ event: "success", ...s }));
  for (const e of result.errors) lines.push(JSON.stringify({ event: "error", ...e }));
  await writeFile(file, lines.join("\n") + "\n", "utf8");
}

function duration(start: string, end: string): string {
  const ms = new Date(end).getTime() - new Date(start).getTime();
  if (ms < 1000) return `${ms}ms`;
  return `${(ms / 1000).toFixed(1)}s`;
}
