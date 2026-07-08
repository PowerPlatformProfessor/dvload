import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import kleur from "kleur";
import {
  parseMapping,
  validateMapping,
  DataverseClient,
  loadRows,
  readTableFromFile,
  type LoadResult,
  type RequestLogEntry,
  type RowSuccess,
} from "@dvload/core";
import { getTokenProvider } from "../auth.js";
import { refreshWorkbook } from "../refresh.js";

interface RunOpts {
  workbook: string;
  refresh?: boolean;
  dryRun?: boolean;
  user?: boolean;
  maxErrors?: number;
}

export async function runCommand(mappingPath: string, opts: RunOpts): Promise<void> {
  const mappingFile = path.resolve(mappingPath);
  const raw = await readFile(mappingFile, "utf8");
  const mapping = parseMapping(JSON.parse(raw));

  const schemaErrs = validateMapping(mapping);
  if (schemaErrs.length > 0) {
    console.error(kleur.red("Mapping has schema errors:"));
    for (const e of schemaErrs) console.error("  - " + e);
    process.exitCode = 2;
    return;
  }

  if (typeof opts.maxErrors === "number") mapping.maxErrors = opts.maxErrors;

  const workbookPath = path.resolve(opts.workbook);

  if (opts.refresh) {
    console.log(kleur.gray("Refreshing Power Query in Excel..."));
    await refreshWorkbook({ workbook: workbookPath });
  }

  console.log(kleur.gray(`Reading table "${mapping.sourceTable}" from ${workbookPath}`));
  const { rows } = await readTableFromFile(workbookPath, {
    tableName: mapping.sourceTable,
    sheetName: mapping.sourceSheet,
  });
  console.log(kleur.gray(`  ${rows.length} row(s) loaded.`));

  const getToken = await getTokenProvider({
    environmentUrl: mapping.environmentUrl,
    forceUser: opts.user,
  });

  const requestLog: RequestLogEntry[] = [];
  const successLog: RowSuccess[] = [];

  const client = new DataverseClient({
    environmentUrl: mapping.environmentUrl,
    getToken,
    onRequest: (entry) => requestLog.push(entry),
    retry: {
      onRetry: (info) => {
        console.log(
          kleur.yellow(
            `\n  [retry ${info.attempt}] ${info.status} ${info.url} — ` +
              `waiting ${(info.delayMs / 1000).toFixed(1)}s (${info.source})`
          )
        );
      },
    },
  });

  const result = await loadRows({
    mapping,
    rows,
    client,
    dryRun: opts.dryRun,
    onProgress: (e) => {
      if (e.type === "row-success") {
        successLog.push(e.success);
      } else if (e.type === "batch") {
        const pct = Math.round((e.processed / Math.max(1, e.total)) * 100);
        process.stdout.write(
          `\r  ${pct}%  created=${e.created}  updated=${e.updated}  ` +
            `skipped=${e.skipped}  failed=${e.failed}  ${e.processed}/${e.total}     `
        );
      } else if (e.type === "done") {
        process.stdout.write("\n");
      }
    },
  });

  await writeRunLog(mappingFile, mapping.logDir, workbookPath, requestLog, successLog, result);

  console.log("");
  console.log(kleur.bold("Run summary"));
  console.log(`  total:     ${result.total}`);
  console.log(`  created:   ${kleur.green(String(result.created))}`);
  console.log(`  updated:   ${kleur.green(String(result.updated))}`);
  console.log(`  skipped:   ${result.skipped}`);
  console.log(`  failed:    ${result.failed > 0 ? kleur.red(String(result.failed)) : "0"}`);
  console.log(`  duration:  ${duration(result.startedAt, result.finishedAt)}`);
  if (result.errors.length > 0) {
    console.log("");
    console.log(kleur.red("First few errors:"));
    for (const e of result.errors.slice(0, 5)) {
      console.log(`  row ${e.rowIndex}: ${e.message}`);
    }
    process.exitCode = 1;
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
  const excelStem = path.basename(workbookPath, path.extname(workbookPath)).replace(/[^\w.-]/g, "_");
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
