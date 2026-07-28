// Read a named Excel table (the kind Power Query writes its output to)
// and yield rows keyed by header. Two entry points:
//   - readTableFromFile: Node-only (fs path)
//   - readTableFromBuffer: works in browser (Office.js add-in) too
//
// We use ExcelJS because it preserves Date types and reads tables natively.
// ExcelJS is untyped from our perspective (it ships types but they don't
// resolve reliably under Node classic module resolution). We interact with
// it via `any` — safe enough because ExcelJS is plain JS underneath.

// eslint-disable-next-line @typescript-eslint/no-explicit-any
import ExcelJS from "exceljs";
import type { SourceRow } from "./types.js";

export interface ReadTableOptions {
  /**
   * The table name (Excel "ListObject" name). Omit to take the first table in
   * the workbook, falling back to the used range of the first worksheet —
   * useful when a file is picked by a human who doesn't know or care whether
   * the sheet has a named table on it.
   */
  tableName?: string;
  /** Optional sheet to constrain the search to. */
  sheetName?: string;
}

export interface ReadTableResult {
  headers: string[];
  rows: SourceRow[];
}

export async function readTableFromBuffer(
  buffer: ArrayBuffer | Uint8Array,
  opts: ReadTableOptions
): Promise<ReadTableResult> {
  const wb = new ExcelJS.Workbook();
  const buf = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  await wb.xlsx.load(buf as unknown as ArrayBuffer);
  return extractTable(wb, opts);
}

export async function readTableFromFile(
  path: string,
  opts: ReadTableOptions
): Promise<ReadTableResult> {
  // Delimited-text sources: same output shape, no workbook/table concepts
  // (tableName and sheetName are ignored).
  if (/\.(csv|tsv)$/i.test(path)) {
    const { readFile } = await import("node:fs/promises");
    const { readTableFromCsvString } = await import("./csv-reader.js");
    const text = await readFile(path, "utf8");
    return readTableFromCsvString(text, {
      delimiter: /\.tsv$/i.test(path) ? "\t" : ",",
    });
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  return extractTable(wb, opts);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractTable(wb: any, opts: ReadTableOptions): ReadTableResult {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheets: any[] = opts.sheetName
    ? [wb.getWorksheet(opts.sheetName)].filter((s) => !!s)
    : wb.worksheets;

  for (const sheet of sheets) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const tables = ((sheet as { tables?: Record<string, any> }).tables) ?? {};
    // No name given: take whatever table this sheet has.
    const name = opts.tableName ?? Object.keys(tables)[0];
    const tbl = name === undefined ? undefined : tables[name];
    if (!tbl) continue;
    // ExcelJS wraps the table definition under a `table` property in some versions
    const tblDef = (tbl as { table?: unknown }).table ?? tbl;
    const ref = (tblDef as { tableRef?: string }).tableRef;
    if (!ref) continue;
    return readByRange(sheet, ref);
  }

  // Still nothing, and the caller didn't insist on a specific table: treat the
  // first non-empty sheet as a header row plus data.
  if (opts.tableName === undefined) {
    for (const sheet of sheets) {
      const range = usedRange(sheet);
      if (range) return readByRange(sheet, range);
    }
    throw new Error(
      opts.sheetName
        ? `Sheet "${opts.sheetName}" is empty.`
        : "The workbook has no tables and no data to read."
    );
  }

  throw new Error(
    `Table "${opts.tableName}" not found${
      opts.sheetName ? ` on sheet "${opts.sheetName}"` : ""
    }.`
  );
}

/** "A1:D57" for a sheet's populated cells, or null when it's empty. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function usedRange(sheet: any): string | null {
  const rowCount: number = sheet?.actualRowCount ?? sheet?.rowCount ?? 0;
  const colCount: number = sheet?.actualColumnCount ?? sheet?.columnCount ?? 0;
  if (!rowCount || !colCount || rowCount < 2) return null;
  return `A1:${columnLetter(colCount)}${rowCount}`;
}

/** 1 -> "A", 27 -> "AA". */
function columnLetter(n: number): string {
  let out = "";
  let rest = n;
  while (rest > 0) {
    const rem = (rest - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    rest = Math.floor((rest - 1) / 26);
  }
  return out;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function readByRange(sheet: any, ref: string): ReadTableResult {
  const [start, end] = ref.split(":");
  const startCell = parseAddress(start);
  const endCell = parseAddress(end);

  const headers: string[] = [];
  for (let c = startCell.col; c <= endCell.col; c++) {
    const cell = sheet.getRow(startCell.row).getCell(c);
    headers.push(String(cell.value ?? "").trim() || `col_${c}`);
  }

  const rows: SourceRow[] = [];
  for (let r = startCell.row + 1; r <= endCell.row; r++) {
    const row: SourceRow = {};
    let allBlank = true;
    for (let c = startCell.col; c <= endCell.col; c++) {
      const v = unwrap(sheet.getRow(r).getCell(c).value);
      row[headers[c - startCell.col]] = v;
      if (v !== null && v !== undefined && v !== "") allBlank = false;
    }
    if (!allBlank) rows.push(row);
  }

  return { headers, rows };
}

/* -------------------------------------------------------------------------- */
/* Enumerating what's in a workbook                                            */
/* -------------------------------------------------------------------------- */

export interface WorkbookTableInfo {
  /** Excel table (ListObject) name, or the sheet name when the sheet has none. */
  name: string;
  sheetName: string;
  /**
   * Whether `name` refers to a real table or to a whole sheet's used range.
   * Callers need this to read the table back: a `sheet` entry has no table
   * name to pass, only a sheet.
   */
  kind: "table" | "sheet";
  /** Data rows, excluding the header. Approximate: blank rows aren't excluded. */
  rowCount: number;
  columns: string[];
}

/**
 * List everything readable in a workbook, without parsing the data.
 *
 * `readTableFromBuffer` answers "give me *the* table", which is right for a
 * saved mapping that already names one. This answers "what's in here?", so a
 * UI can offer a choice — including across several files at once.
 *
 * Only the header row of each range is read, so this stays cheap on a large
 * workbook. `rowCount` comes from the range dimensions rather than a scan,
 * which means it counts blank rows the reader would later drop; it is for
 * display, not for accounting.
 */
export async function listTablesFromBuffer(
  buffer: ArrayBuffer | Uint8Array
): Promise<WorkbookTableInfo[]> {
  const wb = new ExcelJS.Workbook();
  const buf = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  await wb.xlsx.load(buf as unknown as ArrayBuffer);
  return listTables(wb);
}

export async function listTablesFromFile(path: string): Promise<WorkbookTableInfo[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  return listTables(wb);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function listTables(wb: any): WorkbookTableInfo[] {
  const out: WorkbookTableInfo[] = [];

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  for (const sheet of wb.worksheets as any[]) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const tables = ((sheet as { tables?: Record<string, any> }).tables) ?? {};
    const names = Object.keys(tables);

    for (const name of names) {
      const tbl = tables[name];
      const tblDef = (tbl as { table?: unknown }).table ?? tbl;
      const ref = (tblDef as { tableRef?: string }).tableRef;
      if (!ref) continue;
      const described = describeRange(sheet, ref);
      if (described) out.push({ name, sheetName: sheet.name, kind: "table", ...described });
    }

    // A sheet with no named table is still usable — Power Query writes
    // tables, but people paste data too. Only offered when there's no table,
    // matching what extractTable would fall back to.
    if (names.length === 0) {
      const range = usedRange(sheet);
      if (range) {
        const described = describeRange(sheet, range);
        if (described) {
          out.push({ name: sheet.name, sheetName: sheet.name, kind: "sheet", ...described });
        }
      }
    }
  }

  return out;
}

/** Header row and dimensions for a range, without reading the data. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function describeRange(sheet: any, ref: string): { rowCount: number; columns: string[] } | null {
  const [start, end] = ref.split(":");
  if (!start || !end) return null;
  const startCell = parseAddress(start);
  const endCell = parseAddress(end);

  const columns: string[] = [];
  for (let c = startCell.col; c <= endCell.col; c++) {
    const cell = sheet.getRow(startCell.row).getCell(c);
    // Same blank-header rule as readByRange, so the names a caller sees here
    // are the ones the rows will actually be keyed by.
    columns.push(String(cell.value ?? "").trim() || `col_${c}`);
  }

  return { rowCount: Math.max(0, endCell.row - startCell.row), columns };
}

/** Build the failed-rows workbook in memory. Shared by both writers. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function buildRowsWorkbook(headers: string[], rows: SourceRow[], tableName: string): any {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet("Failed rows");
  sheet.addTable({
    name: tableName,
    ref: "A1",
    headerRow: true,
    columns: headers.map((h) => ({ name: h })),
    rows: rows.map((r) => headers.map((h) => (r[h] === undefined ? null : r[h]))),
  });
  return wb;
}

/**
 * Write rows to a new .xlsx with a table named `tableName` (default the
 * same shape dvload reads). Used for the failed-rows re-run file.
 */
export async function writeRowsToFile(
  path: string,
  headers: string[],
  rows: SourceRow[],
  tableName = "FailedRows"
): Promise<void> {
  await buildRowsWorkbook(headers, rows, tableName).xlsx.writeFile(path);
}

/**
 * Same workbook as `writeRowsToFile`, returned as bytes instead of written to
 * disk — the add-in runs in a browser and can only hand the user a Blob.
 */
export async function writeRowsToBuffer(
  headers: string[],
  rows: SourceRow[],
  tableName = "FailedRows"
): Promise<ArrayBuffer> {
  const buf = await buildRowsWorkbook(headers, rows, tableName).xlsx.writeBuffer();
  return buf as ArrayBuffer;
}

/** "A1" -> { col: 1, row: 1 }. */
function parseAddress(addr: string): { col: number; row: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(addr);
  if (!m) throw new Error(`Bad cell address: ${addr}`);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]) };
}

/** Unwrap ExcelJS rich/formula cell values into plain primitives. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function unwrap(value: any): unknown {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value;
  if (typeof value === "object") {
    const v = value as Record<string, unknown>;
    if ("text" in v) return v.text;
    if ("result" in v) return v.result;
    if ("richText" in v && Array.isArray(v.richText)) {
      return (v.richText as Array<{ text?: string }>).map((r) => r.text ?? "").join("");
    }
    if ("hyperlink" in v && "text" in v) return v.text;
  }
  return value;
}
