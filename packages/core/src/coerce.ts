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
      // Explicit format wins: unambiguous, locale-independent parsing.
      if (typeof value === "string" && column.format) {
        const p = parseWithFormat(value.trim(), column.format);
        if (!p) {
          throw new CoerceError(`does not match format "${column.format}"`, column.target, value);
        }
        // Parsed parts are treated as UTC — deterministic regardless of the
        // machine running the load. Use an offset in your data if you need
        // a specific zone.
        return new Date(Date.UTC(p.y, p.mo - 1, p.d, p.h, p.mi, p.s)).toISOString();
      }
      const d = value instanceof Date ? value : new Date(String(value));
      if (Number.isNaN(d.getTime())) throw new CoerceError("invalid datetime", column.target, value);
      return d.toISOString();
    }

    case "dateonly": {
      if (typeof value === "string") {
        const s = value.trim();
        if (column.format) {
          const p = parseWithFormat(s, column.format);
          if (!p) {
            throw new CoerceError(`does not match format "${column.format}"`, column.target, value);
          }
          return `${pad4(p.y)}-${pad2(p.mo)}-${pad2(p.d)}`;
        }
        // ISO-shaped strings: take the date part verbatim. Going through
        // `new Date()` + UTC getters can shift a day for strings that JS
        // parses as *local* time (e.g. "2025-01-01 00:00").
        const iso = /^(\d{4})-(\d{2})-(\d{2})(?:$|[T ])/.exec(s);
        if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
      }
      const d = value instanceof Date ? value : new Date(String(value));
      if (Number.isNaN(d.getTime())) throw new CoerceError("invalid date", column.target, value);
      // Dataverse DateOnly: yyyy-MM-dd (no time, no offset). ExcelJS
      // constructs cell dates as UTC, so UTC getters are correct here.
      return `${pad4(d.getUTCFullYear())}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
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

/* -------------------------------------------------------------------------- */
/* Explicit date formats                                                       */
/* -------------------------------------------------------------------------- */

const pad2 = (n: number): string => String(n).padStart(2, "0");
const pad4 = (n: number): string => String(n).padStart(4, "0");

interface DateParts {
  y: number;
  mo: number;
  d: number;
  h: number;
  mi: number;
  s: number;
}

const FORMAT_TOKENS: Array<[token: string, re: string, field: keyof DateParts]> = [
  ["yyyy", "(\\d{4})", "y"],
  ["MM", "(\\d{2})", "mo"],
  ["M", "(\\d{1,2})", "mo"],
  ["dd", "(\\d{2})", "d"],
  ["d", "(\\d{1,2})", "d"],
  ["HH", "(\\d{2})", "h"],
  ["H", "(\\d{1,2})", "h"],
  ["mm", "(\\d{2})", "mi"],
  ["ss", "(\\d{2})", "s"],
];

const formatCache = new Map<string, { re: RegExp; fields: Array<keyof DateParts> }>();

function compileFormat(format: string): { re: RegExp; fields: Array<keyof DateParts> } {
  const cached = formatCache.get(format);
  if (cached) return cached;
  let pattern = "";
  const fields: Array<keyof DateParts> = [];
  let i = 0;
  outer: while (i < format.length) {
    for (const [token, re, field] of FORMAT_TOKENS) {
      if (format.startsWith(token, i)) {
        pattern += re;
        fields.push(field);
        i += token.length;
        continue outer;
      }
    }
    pattern += format[i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    i++;
  }
  const compiled = { re: new RegExp(`^${pattern}$`), fields };
  formatCache.set(format, compiled);
  return compiled;
}

/**
 * Parse `raw` against a `column.format` string built from the tokens
 * yyyy, MM/M, dd/d, HH/H, mm, ss (all other characters are literals),
 * e.g. "dd/MM/yyyy" or "yyyy-MM-dd HH:mm". Returns null on mismatch or
 * an impossible date (e.g. Feb 30).
 */
export function parseWithFormat(raw: string, format: string): DateParts | null {
  const { re, fields } = compileFormat(format);
  const m = re.exec(raw);
  if (!m) return null;
  const parts: DateParts = { y: 1970, mo: 1, d: 1, h: 0, mi: 0, s: 0 };
  fields.forEach((field, idx) => {
    parts[field] = Number(m[idx + 1]);
  });
  // Reject impossible dates: Date.UTC silently rolls Feb 30 → Mar 2.
  const probe = new Date(Date.UTC(parts.y, parts.mo - 1, parts.d, parts.h, parts.mi, parts.s));
  if (
    probe.getUTCFullYear() !== parts.y ||
    probe.getUTCMonth() !== parts.mo - 1 ||
    probe.getUTCDate() !== parts.d ||
    probe.getUTCHours() !== parts.h ||
    probe.getUTCMinutes() !== parts.mi ||
    probe.getUTCSeconds() !== parts.s
  ) {
    return null;
  }
  return parts;
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
