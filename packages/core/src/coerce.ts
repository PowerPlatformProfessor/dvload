// Cell-level type coercion. Excel/Power Query produces JS dates, numbers,
// booleans, strings, or nulls; this module converts them into the JSON
// shape Dataverse expects on the wire.

import type { ColumnMapping } from "./mapping.js";

export class CoerceError extends Error {
  constructor(message: string, public column: string, public value: unknown) {
    super(`[${column}] ${message}: ${JSON.stringify(value)}`);
    this.name = "CoerceError";
  }
}

const isBlank = (v: unknown): boolean =>
  v === null || v === undefined || (typeof v === "string" && v.trim() === "");

/**
 * Coerce one cell value into the JSON form Dataverse expects for `column.kind`.
 * Returns either the coerced value, `null` (if treatEmptyAsNull and value is
 * blank), or `undefined` to mean "omit this attribute from the payload."
 */
export function coerceValue(value: unknown, column: ColumnMapping): unknown {
  if (isBlank(value)) {
    return column.treatEmptyAsNull ? null : undefined;
  }

  switch (column.kind) {
    case "string":
    case "memo":
      return String(value);

    case "integer":
    case "status":
    case "state": {
      const n = typeof value === "number" ? value : Number(String(value).trim());
      if (!Number.isFinite(n)) throw new CoerceError("not an integer", column.target, value);
      return Math.trunc(n);
    }

    case "decimal":
    case "money":
    case "double": {
      const n = typeof value === "number" ? value : Number(String(value).trim());
      if (!Number.isFinite(n)) throw new CoerceError("not a number", column.target, value);
      return n;
    }

    case "boolean": {
      if (typeof value === "boolean") return value;
      const s = String(value).trim().toLowerCase();
      if (["true", "yes", "y", "1"].includes(s)) return true;
      if (["false", "no", "n", "0"].includes(s)) return false;
      throw new CoerceError("not a boolean", column.target, value);
    }

    case "datetime": {
      const d = value instanceof Date ? value : new Date(String(value));
      if (Number.isNaN(d.getTime())) throw new CoerceError("invalid datetime", column.target, value);
      return d.toISOString();
    }

    case "dateonly": {
      const d = value instanceof Date ? value : new Date(String(value));
      if (Number.isNaN(d.getTime())) throw new CoerceError("invalid date", column.target, value);
      // Dataverse DateOnly: yyyy-MM-dd (no time, no offset)
      const yyyy = d.getUTCFullYear().toString().padStart(4, "0");
      const mm = (d.getUTCMonth() + 1).toString().padStart(2, "0");
      const dd = d.getUTCDate().toString().padStart(2, "0");
      return `${yyyy}-${mm}-${dd}`;
    }

    case "uniqueidentifier":
      return String(value).trim().toLowerCase();

    case "choice": {
      if (typeof value === "number") return value;
      if (column.optionMap && Object.prototype.hasOwnProperty.call(column.optionMap, String(value))) {
        return column.optionMap[String(value)];
      }
      const n = Number(String(value).trim());
      if (Number.isFinite(n)) return n;
      throw new CoerceError("unknown option", column.target, value);
    }

    case "multichoice": {
      const parts = Array.isArray(value)
        ? value
        : String(value).split(/[,;]/).map((s) => s.trim()).filter(Boolean);
      const ints = parts.map((p) => {
        if (typeof p === "number") return p;
        if (column.optionMap && Object.prototype.hasOwnProperty.call(column.optionMap, p)) {
          return column.optionMap[p];
        }
        const n = Number(p);
        if (!Number.isFinite(n)) throw new CoerceError("unknown option", column.target, p);
        return n;
      });
      // Dataverse multi-select serializes as comma-separated integer string.
      return ints.join(",");
    }

    case "lookup":
      // Lookups are handled in load.ts (they require resolution and a
      // navigation-property bind, not a direct value). Pass through here.
      return value;

    default: {
      const _exhaustive: never = column.kind;
      throw new Error(`Unhandled field kind: ${String(_exhaustive)}`);
    }
  }
}

/**
 * Coerce an entire row into the partial Dataverse payload, EXCLUDING lookups.
 * Lookups are bound separately by the loader because they depend on async
 * resolution against Dataverse.
 */
export function coerceRow(
  row: Record<string, unknown>,
  columns: ColumnMapping[]
): { payload: Record<string, unknown>; lookups: ColumnMapping[] } {
  const payload: Record<string, unknown> = {};
  const lookups: ColumnMapping[] = [];

  for (const col of columns) {
    if (col.kind === "lookup") {
      lookups.push(col);
      continue;
    }
    const coerced = coerceValue(row[col.source], col);
    if (coerced !== undefined) payload[col.target] = coerced;
  }

  return { payload, lookups };
}
