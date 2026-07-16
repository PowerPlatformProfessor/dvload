// The .dvmap.json schema. This file is the contract between the add-in
// (which writes mappings) and the CLI (which reads + executes them). Treat
// any breaking change here as a schema-version bump.
//
// Validation is hand-rolled to avoid pulling in a runtime schema library
// (we had brittle install issues with zod). The trade-off is verbose
// checks below in exchange for zero dependencies.

import type { DataverseFieldKind, ConflictMode, SyncAction } from "./types.js";

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
  /**
   * For lookups: how to resolve the source value into a Dataverse record id.
   *  - "guid": the cell already contains the record GUID
   *  - "alternateKey": match on a defined Dataverse alternate key
   *  - "text": match on ANY text attribute via $filter (KingswaySoft-style
   *    text lookup); requires keyAttribute; supports createIfMissing.
   */
  lookupResolution?: "guid" | "alternateKey" | "text";
  /** When lookupResolution = alternateKey or text: the attribute to match on. */
  keyAttribute?: string;
  /**
   * lookupResolution=text only: create a record in bindEntitySet (with
   * keyAttribute = source value) when no match is found. Default false.
   */
  createIfMissing?: boolean;
  /**
   * lookupResolution=text only: what to do when the text matches 2+ records.
   * "error" (default) fails the row; "first" takes the first match.
   */
  duplicateBehavior?: "error" | "first";
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
  /** Concurrent $batch requests in flight (1-8). Default 1 (sequential). */
  concurrency: number;
  /**
   * Send MSCRM.BypassCustomPluginExecution and
   * MSCRM.SuppressCallbackRegistrationExpanderJob on every operation.
   * Skips synchronous plugins and Power Automate triggers during the load.
   * The app user / signed-in user needs the prvBypassCustomPlugins privilege.
   */
  bypassCustomLogic: boolean;
  /**
   * Impersonate this systemuser (GUID) — records show them as creator/owner.
   * Sent as the MSCRMCallerID header. Caller needs prvActOnBehalfOfAnotherUser.
   */
  impersonateUserId?: string;
  /**
   * upsert/sync only: pre-read each target record and drop attributes whose
   * values already match; skip the row entirely if nothing changed. Reduces
   * audit noise and plugin churn at the cost of one GET per row.
   * Lookup columns are always sent (navigation-property values can't be
   * compared cheaply).
   */
  skipUnchanged: boolean;
  /** conflictMode=sync only: what to do with missing records. Default "deactivate". */
  syncAction?: SyncAction;
  /** POST a {text: summary} JSON to this webhook after each run (Teams/Slack compatible). */
  notifyUrl?: string;
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

const CONFLICT_MODES: readonly ConflictMode[] = ["insert", "upsert", "skip-if-exists", "sync"];

const SYNC_ACTIONS: readonly SyncAction[] = ["deactivate", "delete"];

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

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

  const syncAction = input.syncAction ?? (conflictMode === "sync" ? "deactivate" : undefined);
  if (
    syncAction !== undefined &&
    (typeof syncAction !== "string" || !SYNC_ACTIONS.includes(syncAction as SyncAction))
  ) {
    throw new MappingParseError(`syncAction must be one of: ${SYNC_ACTIONS.join(", ")}`);
  }

  const impersonateUserId = optionalString(input.impersonateUserId, "impersonateUserId");
  if (impersonateUserId !== undefined && !GUID_RE.test(impersonateUserId)) {
    throw new MappingParseError("impersonateUserId must be a systemuser GUID");
  }

  // The run summary is POSTed to notifyUrl, so a shared mapping file can
  // point it anywhere. Require TLS (or localhost for testing).
  const notifyUrl = optionalString(input.notifyUrl, "notifyUrl");
  if (notifyUrl !== undefined && !/^(https:\/\/|http:\/\/localhost(:\d+)?\/)/i.test(notifyUrl)) {
    throw new MappingParseError("notifyUrl must be an https:// URL (or http://localhost for testing)");
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
    concurrency:
      input.concurrency === undefined
        ? 1
        : requireNumberInRange(input.concurrency, "concurrency", 1, 8),
    bypassCustomLogic: optionalBoolean(input.bypassCustomLogic, false, "bypassCustomLogic"),
    impersonateUserId,
    skipUnchanged: optionalBoolean(input.skipUnchanged, false, "skipUnchanged"),
    syncAction: syncAction as SyncAction | undefined,
    notifyUrl,
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
    lookupResolution !== "alternateKey" &&
    lookupResolution !== "text"
  ) {
    throw new MappingParseError(
      `${path}.lookupResolution must be "guid", "alternateKey", or "text"`
    );
  }

  const duplicateBehavior = input.duplicateBehavior;
  if (duplicateBehavior !== undefined && duplicateBehavior !== "error" && duplicateBehavior !== "first") {
    throw new MappingParseError(`${path}.duplicateBehavior must be "error" or "first"`);
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
    lookupResolution: lookupResolution as "guid" | "alternateKey" | "text" | undefined,
    keyAttribute: optionalString(input.keyAttribute, `${path}.keyAttribute`),
    createIfMissing: optionalBoolean(input.createIfMissing, false, `${path}.createIfMissing`) || undefined,
    duplicateBehavior: duplicateBehavior as "error" | "first" | undefined,
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
      if (
        (col.lookupResolution === "alternateKey" || col.lookupResolution === "text") &&
        !col.keyAttribute
      ) {
        errors.push(
          `Lookup column ${col.target} uses ${col.lookupResolution} but no keyAttribute set`
        );
      }
    } else {
      if (col.createIfMissing) {
        errors.push(`Column ${col.target}: createIfMissing only applies to lookup columns`);
      }
      if (col.duplicateBehavior) {
        errors.push(`Column ${col.target}: duplicateBehavior only applies to lookup columns`);
      }
    }
    if (col.kind === "lookup" && col.lookupResolution !== "text") {
      if (col.createIfMissing) {
        errors.push(`Lookup column ${col.target}: createIfMissing requires lookupResolution=text`);
      }
      if (col.duplicateBehavior) {
        errors.push(`Lookup column ${col.target}: duplicateBehavior requires lookupResolution=text`);
      }
    }
  }
  if (
    (m.conflictMode === "upsert" || m.conflictMode === "sync") &&
    (!m.upsertKey || m.upsertKey.length === 0)
  ) {
    errors.push(`conflictMode=${m.conflictMode} requires at least one upsertKey attribute`);
  }
  if (m.syncAction && m.conflictMode !== "sync") {
    errors.push("syncAction only applies when conflictMode=sync");
  }
  if (m.skipUnchanged && m.conflictMode !== "upsert" && m.conflictMode !== "sync") {
    errors.push("skipUnchanged requires conflictMode=upsert or sync");
  }
  // upsertKey attributes must exist in columns — catch at validate time, not per-row.
  if (m.upsertKey) {
    for (const attr of m.upsertKey) {
      if (!m.columns.some((c) => c.target === attr)) {
        errors.push(`upsertKey attribute "${attr}" is not mapped in columns`);
      }
    }
  }
  return errors;
}
