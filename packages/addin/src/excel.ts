// Office.js helpers: list tables in the workbook, read a table's headers
// and rows. Power Query writes its output as a table (ListObject), so this
// is how we get from "user finished their PQ transform" to rows we can map.

import type { SourceRow } from "@dvload/core";

export interface TableInfo {
  name: string;
  worksheetName: string;
  rowCount: number;
  columns: string[];
}

export async function listTables(): Promise<TableInfo[]> {
  return Excel.run(async (ctx) => {
    const tables = ctx.workbook.tables.load(["items/name", "items/worksheet/name"]);
    await ctx.sync();

    const out: TableInfo[] = [];
    for (const t of tables.items) {
      const headerRange = t.getHeaderRowRange().load("values");
      const dataRange = t.getDataBodyRange().load("rowCount");
      await ctx.sync();
      out.push({
        name: t.name,
        worksheetName: t.worksheet.name,
        rowCount: dataRange.rowCount,
        columns: (headerRange.values[0] ?? []).map((h) => String(h)),
      });
    }
    return out;
  });
}

export async function readTable(name: string): Promise<{ headers: string[]; rows: SourceRow[] }> {
  return Excel.run(async (ctx) => {
    const t = ctx.workbook.tables.getItem(name);
    const headerRange = t.getHeaderRowRange().load("values");
    const dataRange = t.getDataBodyRange().load(["values", "valueTypes"]);
    await ctx.sync();

    const headers = (headerRange.values[0] ?? []).map((h) => String(h));
    const rows: SourceRow[] = [];

    const values = dataRange.values;
    const types = dataRange.valueTypes;

    for (let r = 0; r < values.length; r++) {
      const row: SourceRow = {};
      let allBlank = true;
      for (let c = 0; c < headers.length; c++) {
        const v = unwrapCell(values[r][c], types?.[r]?.[c]);
        row[headers[c]] = v;
        if (v !== null && v !== undefined && v !== "") allBlank = false;
      }
      if (!allBlank) rows.push(row);
    }
    return { headers, rows };
  });
}

/**
 * Excel returns dates as serial numbers (days since 1899-12-30). We detect
 * those via valueTypes and convert to a JS Date so coerce.ts sees the right
 * shape.
 */
function unwrapCell(value: unknown, type: string | undefined): unknown {
  if (value === null || value === "") return null;
  if (type === "Double" && typeof value === "number" && value > 25569 && value < 80000) {
    // Heuristic: looks like an Excel date serial. Caller can override via
    // explicit `kind: dateonly` mappings; this just preserves the type signal.
    return excelSerialToDate(value);
  }
  return value;
}

function excelSerialToDate(serial: number): Date {
  // Excel epoch is 1899-12-30 (accounting for the 1900 leap-year bug).
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + serial * 86_400_000);
}
