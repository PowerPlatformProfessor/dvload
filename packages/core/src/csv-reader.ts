// Minimal RFC 4180 CSV/TSV reader — no dependency. Values stay strings
// (plus null for empty cells); the coercion layer already converts strings
// into the wire types Dataverse expects, so no type sniffing is done here.

import type { SourceRow } from "./types.js";
import type { ReadTableResult } from "./xlsx-reader.js";

export interface ReadCsvOptions {
  /** Field delimiter. Default "," ("\t" for .tsv files). */
  delimiter?: string;
}

/** Parse CSV text into rows of fields, honoring quoted fields per RFC 4180. */
export function parseCsv(text: string, delimiter = ","): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  // Strip a UTF-8 BOM (Excel writes one).
  let i = text.charCodeAt(0) === 0xfeff ? 1 : 0;

  for (; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
    } else if (ch === '"' && field === "") {
      inQuotes = true;
    } else if (ch === delimiter) {
      row.push(field);
      field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      field = "";
      rows.push(row);
      row = [];
    } else {
      field += ch;
    }
  }
  // Trailing field/row without a final newline.
  if (field !== "" || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/**
 * Read delimited text into the same shape the xlsx reader produces: first
 * row is the header, blank rows are dropped, empty cells become null.
 */
export function readTableFromCsvString(text: string, opts?: ReadCsvOptions): ReadTableResult {
  const raw = parseCsv(text, opts?.delimiter ?? ",");
  if (raw.length === 0) {
    throw new Error("CSV file is empty — expected a header row.");
  }
  const headers = raw[0].map((h, i) => (h.trim() === "" ? `col_${i + 1}` : h.trim()));
  const rows: SourceRow[] = [];
  for (let r = 1; r < raw.length; r++) {
    const cells = raw[r];
    const row: SourceRow = {};
    let allBlank = true;
    for (let c = 0; c < headers.length; c++) {
      const v = cells[c] !== undefined && cells[c] !== "" ? cells[c] : null;
      row[headers[c]] = v;
      if (v !== null) allBlank = false;
    }
    if (!allBlank) rows.push(row);
  }
  return { headers, rows };
}
