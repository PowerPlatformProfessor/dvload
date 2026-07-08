// The .dvmap.json schema. This file is the contract between the add-in
// (which writes mappings) and the CLI (which reads + executes them). Treat
// any breaking change here as a schema-version bump.
//
// Validation is hand-rolled to avoid pulling in a runtime schema library
// (we had brittle install issues with zod). The trade-off is verbose
// checks below in exchange for zero dependencies.

import type { DataverseFieldKind, ConflictMode } from "./types.js";

export const SCHEMA_VERSION = 1 as const;

export interface ColumnMapping {
  /** Source column header in the Excel table. */
  source: string;
  /** Target Dataverse logical name (for lookups, the navigation property). */
  target: string;
  /** Dataverse type to coerce into. */
  kind: DataverseFieldKind;
  /** For lookups: the entity set name to bind to (e.g. "accounts"). */
  bindEntitySet?: string;
  /** For lookups: how to resolve the source value into a Dataverse record id. */
  lookupResolution?: "guid" | "alternateKey";
  /** When lookupResolution = alternateKey: the attribute to match on. */
  keyAttribute?: string;
  /** For choice/multichoice/status/state: optional label-to-int map. */
  optionMap?: Record<string, number>;
  /** If true, empty source values are sent as null (clears the field). */
  treatEmptyAsNull: boolean;
  /** Override default coercion (e.g. an explicit date format string). */
  format?: string;
  /** Notes shown in the add-in UI; ignored by the engine. */
  notes?: string;
}

export interface Mapping {
  schemaVersion: typeof SCHEMA_VERSION;
  /** Human-readable name for this mapping. */
  name: string;
  /** Optional description shown in the UI / CLI. */
  description?: string;
  /** Created/updated timestamps (ISO 8601, UTC). */
  createdAt?: string;
  updatedAt?: string;
  /** Dataverse environment URL, e.g. "https://contoso.crm.dynamics.com". */
  environmentUrl: string;
  /** Target entity set name (plural logical name), e.g. "contacts". */
  targetEntitySet: string;
  /** Source table name in the Excel workbook. */
  sourceTable: string;
  /** Optional sheet name; if omitted, the table is searched workbook-wide. */
  sourceSheet?: string;
  /** Column-to-field mappings (in apply order). */
  columns: ColumnMapping[];
  /** Insert vs upsert behavior. */
  conflictMode: ConflictMode;
  /** For upsert: the alternate-key attribute name(s) used to match. */
  upsertKey?: string[];
  /** Rows per $batch changeset. Dataverse hard-caps at 1000. */
  batchSize: number;
  /** Stop after this many row errors (0 = unlimited). */
  maxErrors: number;
  /** Where to write run logs (relative to mapping file). */
  logDir: string;
}

/* -------------------------------------------------------------------------- */
/* Constants                                                                   */
/* -------------------------------------------------------------------------- */

const FIELD_KINDS: readonly DataverseFieldKind[] = [
  "string",
  "memo",
  "integer",
  "decimal",
  "money",
  "double",
  "boolean",
  "datetime",
  "dateonly",
  "uniqueidentifier",
  "lookup",
  "choice",
  "multichoice",
  "status",
  "state",
];

const CONFLICT_MODES: readonly ConflictMode[] = ["insert", "upsert", "skip-if-exists"];

/* -------------------------------------------------------------------------- */
/* Parsing helpers                                                             */
/* -------------------------------------------------------------------------- */

export class MappingParseError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "MappingParseError";
  }
}

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

function requireString(v: unknown, path: string): string {
  if (typeof v !== "string" || v.length === 0) {
    throw new MappingParseError(`${path} must be a non-empty string`);
  }
  return v;
}

function optionalString(v: unknown, path: string): string | undefined {
  if (v === undefined || v === null) return undefined;
  if (typeof v !== "string") throw new MappingParseError(`${path} must be a string if present`);
  return v;
}

function requireNumberInRange(v: unknown, path: string, min: number, max: number): number {
  if (typeof v !== "number" || !Number.isFinite(v) || !Number.isInteger(v)) {
    throw new MappingParseError(`${path} must be an integer`);
  }
  if (v < min || v > max) throw new MappingParseError(`${path} must be in [${min}, ${max}]`);
  return v;
}

function optionalBoolean(v: unknown, def: boolean, path: string): boolean {
  if (v === undefined || v === null) return def;
  if (typeof v !== "boolean") throw new MappingParseError(`${path} must be a boolean`);
  return v;
}

function optionalStringArray(v: unknown, path: string): string[] | undefined {
  if (v === undefined || v === null) return undefined;
  if (!Array.isArray(v) || v.some((x) => typeof x !== "string")) {
    throw new MappingParseError(`${path} must be an array of strings`);
  }
  return v as string[];
}

/* -------------------------------------------------------------------------- */
/* Public API                                                                  */
/* -------------------------------------------------------------------------- */

/** Parse + validate a mapping document loaded from JSON. */
export function parseMapping(input: unknown): Mapping {
  if (!isPlainObject(input)) throw new MappingParseError("mapping must be an object");

  if (input.schemaVersion !== SCHEMA_VERSION) {
    throw new MappingParseError(
      `schemaVersion must be ${SCHEMA_VERSION}, got ${JSON.stringify(input.schemaVersion)}`
    );
  }

  if (!Array.isArray(input.columns) || input.columns.length === 0) {
    throw new MappingParseError("columns must be a non-empty array");
  }
  const columns: ColumnMapping[] = input.columns.map((c, i) =>
    parseColumnMapping(c, `columns[${i}]`)
  );

  const conflictMode = input.conflictMode ?? "insert";
  if (typeof conflictMode !== "string" || !CONFLICT_MODES.includes(conflictMode as ConflictMode)) {
    throw new MappingParseError(
      `conflictMode must be one of: ${CONFLICT_MODES.join(", ")}`
    );
  }

  return {
    schemaVersion: SCHEMA_VERSION,
    name: requireString(input.name, "name"),
    description: optionalString(input.description, "description"),
    createdAt: optionalString(input.createdAt, "createdAt"),
    updatedAt: optionalString(input.updatedAt, "updatedAt"),
    environmentUrl: requireString(input.environmentUrl, "environmentUrl"),
    targetEntitySet: requireString(input.targetEntitySet, "targetEntitySet"),
    sourceTable: requireString(input.sourceTable, "sourceTable"),
    sourceSheet: optionalString(input.sourceSheet, "sourceSheet"),
    columns,
    conflictMode: conflictMode as ConflictMode,
    upsertKey: optionalStringArray(input.upsertKey, "upsertKey"),
    batchSize:
      input.batchSize === undefined
        ? 100
        : requireNumberInRange(input.batchSize, "batchSize", 1, 1000),
    maxErrors:
      input.maxErrors === undefined
        ? 0
        : requireNumberInRange(input.maxErrors, "maxErrors", 0, Number.MAX_SAFE_INTEGER),
    logDir: optionalString(input.logDir, "logDir") ?? "./logs",
  };
}

function parseColumnMapping(input: unknown, path: string): ColumnMapping {
  if (!isPlainObject(input)) throw new MappingParseError(`${path} must be an object`);

  const kind = input.kind;
  if (typeof kind !== "string" || !FIELD_KINDS.includes(kind as DataverseFieldKind)) {
    throw new MappingParseError(`${path}.kind must be one of: ${FIELD_KINDS.join(", ")}`);
  }

  const lookupResolution = input.lookupResolution;
  if (
    lookupResolution !== undefined &&
    lookupResolution !== "guid" &&
    lookupResolution !== "alternateKey"
  ) {
    throw new MappingParseError(`${path}.lookupResolution must be "guid" or "alternateKey"`);
  }

  let optionMap: Record<string, number> | undefined;
  if (input.optionMap !== undefined && input.optionMap !== null) {
    if (!isPlainObject(input.optionMap)) {
      throw new MappingParseError(`${path}.optionMap must be an object`);
    }
    optionMap = {};
    for (const [k, v] of Object.entries(input.optionMap)) {
      if (typeof v !== "number" || !Number.isFinite(v)) {
        throw new MappingParseError(`${path}.optionMap["${k}"] must be a number`);
      }
      optionMap[k] = v;
    }
  }

  return {
    source: requireString(input.source, `${path}.source`),
    target: requireString(input.target, `${path}.target`),
    kind: kind as DataverseFieldKind,
    bindEntitySet: optionalString(input.bindEntitySet, `${path}.bindEntitySet`),
    lookupResolution: lookupResolution as "guid" | "alternateKey" | undefined,
    keyAttribute: optionalString(input.keyAttribute, `${path}.keyAttribute`),
    optionMap,
    treatEmptyAsNull: optionalBoolean(input.treatEmptyAsNull, true, `${path}.treatEmptyAsNull`),
    format: optionalString(input.format, `${path}.format`),
    notes: optionalString(input.notes, `${path}.notes`),
  };
}

/** Serialize a mapping to a stable JSON form for writing to disk. */
export function serializeMapping(m: Mapping): string {
  return JSON.stringify({ ...m, updatedAt: new Date().toISOString() }, null, 2) + "\n";
}

/** Cheap structural validation for use in the add-in's UI before save. */
export function validateMapping(m: Mapping): string[] {
  const errors: string[] = [];
  const seen = new Set<string>();
  for (const col of m.columns) {
    if (seen.has(col.target)) errors.push(`Duplicate target attribute: ${col.target}`);
    seen.add(col.target);

    if (col.kind === "lookup") {
      if (!col.bindEntitySet) errors.push(`Lookup column ${col.target} missing bindEntitySet`);
      if (!col.lookupResolution) errors.push(`Lookup column ${col.target} missing lookupResolution`);
      if (col.lookupResolution === "alternateKey" && !col.keyAttribute) {
        errors.push(`Lookup column ${col.target} uses alternateKey but no keyAttribute set`);
      }
    }
  }
  if (m.conflictMode === "upsert" && (!m.upsertKey || m.upsertKey.length === 0)) {
    errors.push("conflictMode=upsert requires at least one upsertKey attribute");
  }
  return errors;
}
