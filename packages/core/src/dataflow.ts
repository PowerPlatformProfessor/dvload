// Read Power Platform dataflows straight out of Dataverse.
//
// WHY THIS EXISTS
//
// A .pqt export carries the M code and nothing else — Dataverse strips the
// destination config on the way out (see pqt.ts). The mappings a user spent
// an afternoon building in the dataflow editor are simply not in the file.
//
// They are, however, sitting in Dataverse. Dataflows are ordinary rows in
// `msdyn_dataflow`, and two of its columns hold everything we need:
//
//   msdyn_mashupdocument   the M code, byte-identical to MashupDocument.pq
//   msdyn_mashupsettings   a JSON blob whose shape *is* MashupMetadata
//
// That second column is the win. It deserializes directly into the
// `MashupMetadata` interface pqt.ts already defines — same name-keyed
// `QueriesMetadata`, same per-query `FieldsMetadata`. So a live dataflow can
// be reshaped into a `PqtArchive` and every existing consumer
// (`buildWorkbookWithQueries`, `mappingsFromPqtAll`, `injectMappingIntoPqt`)
// works on it unchanged. A dataflow and a .pqt become the same thing to the
// rest of the codebase; only this file knows the difference.
//
// THREE THINGS THE SETTINGS BLOB HAS THAT A .pqt DOES NOT
//
//   1. HostProperties.MappingEditorProperties — a *stringified* JSON object
//      carrying `primaryKeyLogicalName`, the alternate key the dataflow
//      upserts on. mappingFromPqt() has to hardcode conflictMode=insert
//      precisely because a .pqt can't supply this. We can.
//   2. Lookups, encoded as dotted targets ("ParentAccountId.asker_act_id").
//      Naive parsing types these as whatever DestinationFieldType says
//      (usually "Memo") and produces a mapping that fails at load time.
//   3. DataflowMetadata.DataflowState — every dataflow exists twice, once as
//      "Source" (the editing draft) and once as "Active" (what actually
//      runs), cross-linked by RelatedDataflowId. Listing without filtering
//      shows each dataflow twice under an identical name.

import type { DataverseClient } from "./dataverse.js";
import type { ColumnMapping, Mapping } from "./mapping.js";
import { SCHEMA_VERSION } from "./mapping.js";
import {
  STANDARD_CONTENT_TYPES,
  dataflowTypeToDataverseKind,
  normalizeQueriesMetadata,
  parseQueryNames,
  type MashupMetadata,
  type PqtArchive,
  type QueryMetadata,
} from "./pqt.js";

/* -------------------------------------------------------------------------- */
/* Types                                                                       */
/* -------------------------------------------------------------------------- */

/** Publication state, from DataflowMetadata.DataflowState. */
export type DataflowState = "Active" | "Source" | "Unknown";

/**
 * msdyn_mashupsettings is MashupMetadata plus a few Dataverse-only keys that
 * never appear in a .pqt. `DataflowMetadata` is the one we read: publication
 * state, ownership and connection references, stored as an escaped JSON
 * string rather than a nested object.
 */
export type DataflowSettings = MashupMetadata & { DataflowMetadata?: unknown };

/** The msdyn_dataflow columns we read. */
export const DATAFLOW_LIST_SELECT = "msdyn_dataflowid,msdyn_name,msdyn_mashupsettings,modifiedon";

export const DATAFLOW_FULL_SELECT =
  "msdyn_dataflowid,msdyn_name,msdyn_mashupdocument,msdyn_mashupsettings,modifiedon";

/** A raw msdyn_dataflow row, only the fields we care about. */
export interface DataflowRow {
  msdyn_dataflowid?: string;
  msdyn_name?: string;
  msdyn_mashupdocument?: string | null;
  msdyn_mashupsettings?: string | null;
  modifiedon?: string;
}

/** One entry in the picker. Cheap to build; no metadata calls. */
export interface DataflowSummary {
  id: string;
  name: string;
  state: DataflowState;
  modifiedOn?: string;
  /** Every query in the dataflow, in declaration order. */
  queryNames: string[];
  /**
   * Queries that actually write to Dataverse (LoadEnabled), with their
   * destination. Staging queries used only as join sources are excluded.
   */
  loadTargets: Array<{ queryName: string; entityName: string; fieldCount: number }>;
}

/** Everything needed to build outputs for one dataflow. */
export interface DataflowDetail extends DataflowSummary {
  archive: PqtArchive;
  /** Parsed msdyn_mashupsettings. */
  settings: DataflowSettings;
}

/**
 * Metadata lookups needed to turn dataflow field metadata into a valid
 * `.dvmap.json`. Injected rather than taking a client directly so the
 * conversion is unit-testable without a live environment — and so a caller
 * that doesn't care about lookups can pass nothing and take the fallbacks.
 */
export interface DataflowMetadataResolver {
  /** account -> accounts. */
  entitySetName(entityLogicalName: string): Promise<string | undefined>;
  /** Attributes spanned by an alternate key, by the key's logical name. */
  keyAttributes(entityLogicalName: string, keyLogicalName: string): Promise<string[]>;
  /** Entity set a lookup attribute points at, for bindEntitySet. */
  lookupEntitySet(entityLogicalName: string, attrLogicalName: string): Promise<string | undefined>;
}

/* -------------------------------------------------------------------------- */
/* Reading from Dataverse                                                      */
/* -------------------------------------------------------------------------- */

export interface ListDataflowsOptions {
  /**
   * Include "Source" rows — the editing drafts. Off by default: every
   * dataflow has one, it duplicates the Active row under the same name, and
   * it can be arbitrarily far ahead of what is actually published.
   */
  includeDrafts?: boolean;
}

/**
 * List dataflows in the connected environment, newest first.
 *
 * Note this pulls `msdyn_mashupsettings` for every row (~50-100KB each)
 * because publication state lives inside that blob and nowhere cheaper.
 * Environments hold tens of dataflows, not thousands, so the cost is a
 * one-off page load rather than something worth optimising around.
 */
export async function listDataflows(
  client: DataverseClient,
  opts: ListDataflowsOptions = {}
): Promise<DataflowSummary[]> {
  const rows = (await client.queryAll(
    `msdyn_dataflows?$select=${DATAFLOW_LIST_SELECT}&$orderby=modifiedon desc`
  )) as DataflowRow[];

  const all = rows.map(summarizeDataflow);
  if (opts.includeDrafts) return all;

  const active = all.filter((d) => d.state === "Active");
  // A dataflow that has never been published has no Active row at all.
  // Dropping it entirely would hide it from the picker with no explanation,
  // so fall back to the draft for names that have no published counterpart.
  const publishedNames = new Set(active.map((d) => d.name));
  const orphanDrafts = all.filter((d) => d.state !== "Active" && !publishedNames.has(d.name));
  return [...active, ...orphanDrafts];
}

/** Fetch one dataflow, including its M document. */
export async function getDataflow(client: DataverseClient, dataflowId: string): Promise<DataflowDetail> {
  const rows = (await client.queryAll(
    `msdyn_dataflows?$select=${DATAFLOW_FULL_SELECT}` + `&$filter=msdyn_dataflowid eq ${dataflowId}`
  )) as DataflowRow[];
  const row = rows[0];
  if (!row) throw new Error(`No dataflow with id ${dataflowId} in this environment.`);
  return detailFromRow(row);
}

/**
 * Find a dataflow by display name (exact, then case-insensitive). Names are
 * not unique in Dataverse, so an ambiguous match is an error rather than a
 * silent pick — importing the wrong dataflow is worse than a failed command.
 */
export async function findDataflowByName(
  client: DataverseClient,
  name: string,
  opts: ListDataflowsOptions = {}
): Promise<DataflowSummary> {
  const all = await listDataflows(client, opts);
  const exact = all.filter((d) => d.name === name);
  const matches = exact.length > 0 ? exact : all.filter((d) => d.name.toLowerCase() === name.toLowerCase());

  if (matches.length === 0) {
    throw new Error(
      `No dataflow named "${name}". Available: ${all.map((d) => d.name).join(", ") || "(none)"}`
    );
  }
  if (matches.length > 1) {
    throw new Error(
      `"${name}" matches ${matches.length} dataflows. Pass the id instead: ` +
        matches.map((d) => d.id).join(", ")
    );
  }
  return matches[0];
}

/* -------------------------------------------------------------------------- */
/* Row -> summary / archive                                                    */
/* -------------------------------------------------------------------------- */

export function summarizeDataflow(row: DataflowRow): DataflowSummary {
  const settings = parseMashupSettings(row.msdyn_mashupsettings);
  const queries = settings.QueriesMetadata;
  const queryNames = Object.keys(queries);

  return {
    id: String(row.msdyn_dataflowid ?? ""),
    name: String(row.msdyn_name ?? "(unnamed)"),
    state: readDataflowState(settings),
    modifiedOn: row.modifiedon,
    queryNames: queryNames.length > 0 ? queryNames : parseQueryNames(row.msdyn_mashupdocument ?? ""),
    loadTargets: queryNames
      .filter((n) => queries[n]?.LoadEnabled)
      .map((n) => ({
        queryName: n,
        entityName: String(queries[n].EntityName ?? ""),
        fieldCount: Object.keys(queries[n].FieldsMetadata ?? {}).length,
      })),
  };
}

export function detailFromRow(row: DataflowRow): DataflowDetail {
  const settings = parseMashupSettings(row.msdyn_mashupsettings);
  return {
    ...summarizeDataflow(row),
    settings,
    archive: dataflowToPqtArchive(row, settings),
  };
}

/**
 * Reshape a dataflow row into the archive type the .pqt pipeline already
 * speaks. Deliberately a thin adapter: the settings blob *is* MashupMetadata,
 * so there is no translation layer to get wrong.
 */
export function dataflowToPqtArchive(
  row: DataflowRow,
  settings = parseMashupSettings(row.msdyn_mashupsettings)
): PqtArchive {
  return {
    mashupDocument: row.msdyn_mashupdocument ?? "",
    mashupMetadata: settings,
    metadata: {
      Name: String(row.msdyn_name ?? "Dataflow"),
      Description: "",
      Version: "1.0.0.0",
    },
    contentTypes: STANDARD_CONTENT_TYPES,
    // Dataverse always emits the name-keyed object form.
    queriesMetadataShape: "object",
  };
}

/**
 * Parse msdyn_mashupsettings into MashupMetadata.
 *
 * Tolerant by design: a dataflow whose settings are missing or malformed
 * should degrade to "no queries described" — callers then fall back to
 * parsing `shared` declarations out of the M document — rather than taking
 * down the whole listing because one row is odd.
 */
export function parseMashupSettings(raw: string | null | undefined): DataflowSettings {
  const empty: DataflowSettings = {
    QueryGroups: [],
    DocumentLocale: "en-US",
    QueriesMetadata: {},
    FastCombine: false,
    AllowNativeQueries: false,
  };
  if (!raw) return empty;

  let parsed: (DataflowSettings & { QueriesMetadata?: unknown }) | undefined;
  try {
    parsed = JSON.parse(raw) as DataflowSettings & { QueriesMetadata?: unknown };
  } catch {
    return empty;
  }
  if (!parsed || typeof parsed !== "object") return empty;

  const { queriesMetadata } = normalizeQueriesMetadata(parsed.QueriesMetadata);
  return { ...empty, ...parsed, QueriesMetadata: queriesMetadata };
}

/**
 * Publication state, out of the doubly-encoded DataflowMetadata string.
 * Unknown when absent — treated as a draft by the listing filter, which errs
 * toward hiding rather than showing a dataflow twice.
 */
export function readDataflowState(settings: DataflowSettings): DataflowState {
  const meta = parseNestedJson(settings.DataflowMetadata);
  const state = meta?.DataflowState;
  if (state === "Active" || state === "Source") return state;
  return "Unknown";
}

/** Alternate key the dataflow upserts on, if the editor recorded one. */
export function readPrimaryKeyLogicalName(q: QueryMetadata): string | undefined {
  const props = parseNestedJson(q.HostProperties?.MappingEditorProperties);
  const key = props?.primaryKeyLogicalName;
  return typeof key === "string" && key.length > 0 ? key : undefined;
}

/**
 * These blobs are JSON *inside* a JSON string — `MappingEditorProperties` and
 * `DataflowMetadata` are both stored escaped. Anything unparseable is treated
 * as absent; the caller's fallback path is always the safe one.
 */
function parseNestedJson(value: unknown): Record<string, unknown> | undefined {
  if (value && typeof value === "object") return value as Record<string, unknown>;
  if (typeof value !== "string" || value.length === 0) return undefined;
  try {
    const parsed = JSON.parse(value) as unknown;
    return parsed && typeof parsed === "object" ? (parsed as Record<string, unknown>) : undefined;
  } catch {
    return undefined;
  }
}

/* -------------------------------------------------------------------------- */
/* Dataflow -> .dvmap.json                                                     */
/* -------------------------------------------------------------------------- */

export interface DataflowMappingOptions {
  environmentUrl: string;
  /**
   * Metadata lookups. Omit to get a best-effort mapping with no network
   * calls: lookups keep their dotted target and are flagged in `notes`,
   * entity sets fall back to a naive pluralisation, and upsert is not
   * inferred. `validateMapping` will surface anything left incomplete.
   */
  resolver?: DataflowMetadataResolver;
}

/** Build a mapping for one query in a dataflow. */
export async function mappingFromDataflow(
  detail: Pick<DataflowDetail, "name" | "settings">,
  queryName: string,
  opts: DataflowMappingOptions
): Promise<Mapping> {
  const q = detail.settings.QueriesMetadata[queryName];
  if (!q) throw new Error(`Query "${queryName}" not found in dataflow "${detail.name}".`);

  const entityLogical = String(q.EntityName ?? "").toLowerCase();
  const notes: string[] = [];

  const columns: ColumnMapping[] = [];
  for (const [target, fm] of Object.entries(q.FieldsMetadata ?? {})) {
    columns.push(
      await buildColumn(target, fm.SourceColumnName, String(fm.DestinationFieldType), {
        entityLogical,
        resolver: opts.resolver,
        notes,
      })
    );
  }

  const { conflictMode, upsertKey, keyNote } = await resolveConflictMode(
    q,
    entityLogical,
    columns,
    opts.resolver
  );
  if (keyNote) notes.push(keyNote);

  if (q.DeleteExistingDataOnLoad) {
    notes.push(
      "The source dataflow used DeleteExistingDataOnLoad (truncate-and-reload). " +
        "dvload has no equivalent and does not truncate anything."
    );
  }

  return {
    schemaVersion: SCHEMA_VERSION,
    name: `${detail.name} - ${queryName}`,
    description: notes.length > 0 ? notes.join(" ") : undefined,
    environmentUrl: opts.environmentUrl,
    targetEntitySet: await resolveEntitySet(entityLogical, opts.resolver),
    sourceTable: queryName,
    columns,
    conflictMode,
    upsertKey,
    batchSize: 100,
    maxErrors: 0,
    logDir: "./logs",
    concurrency: 1,
    bypassCustomLogic: false,
    skipUnchanged: false,
  };
}

/**
 * One mapping per load-enabled query. Staging queries (LoadEnabled=false)
 * are skipped: they exist to be joined against, have no destination and no
 * FieldsMetadata, and would only produce empty mappings.
 */
export async function mappingsFromDataflow(
  detail: Pick<DataflowDetail, "name" | "settings">,
  opts: DataflowMappingOptions
): Promise<Record<string, Mapping>> {
  const out: Record<string, Mapping> = {};
  for (const [name, q] of Object.entries(detail.settings.QueriesMetadata)) {
    if (!q.LoadEnabled) continue;
    out[name] = await mappingFromDataflow(detail, name, opts);
  }
  return out;
}

interface ColumnContext {
  entityLogical: string;
  resolver?: DataflowMetadataResolver;
  notes: string[];
}

async function buildColumn(
  target: string,
  source: string,
  destinationFieldType: string,
  ctx: ColumnContext
): Promise<ColumnMapping> {
  const dot = target.indexOf(".");
  if (dot <= 0) {
    return {
      source,
      target,
      kind: dataflowTypeToDataverseKind(destinationFieldType),
      treatEmptyAsNull: true,
    };
  }

  // "ParentAccountId.asker_act_id": bind the ParentAccountId lookup by
  // matching asker_act_id on the referenced record. DestinationFieldType is
  // the *key column's* type here ("Memo"), not the lookup's — trusting it
  // would produce a text column that writes a GUID-shaped string into a
  // relationship field and fails at load time.
  const navProperty = target.slice(0, dot);
  const keyAttribute = target.slice(dot + 1);
  const bindEntitySet = await ctx.resolver?.lookupEntitySet(ctx.entityLogical, navProperty.toLowerCase());

  if (!bindEntitySet) {
    ctx.notes.push(
      `Lookup "${target}" needs a bindEntitySet; the referenced table could not be ` +
        `resolved from metadata. Set it in the task pane before running.`
    );
  }

  return {
    source,
    target: navProperty,
    kind: "lookup",
    lookupResolution: "alternateKey",
    keyAttribute,
    bindEntitySet,
    treatEmptyAsNull: true,
  };
}

async function resolveEntitySet(entityLogical: string, resolver?: DataflowMetadataResolver): Promise<string> {
  if (!entityLogical) return "";
  const resolved = await resolver?.entitySetName(entityLogical);
  if (resolved) return resolved;
  // Dataverse's own pluralisation is irregular enough that this is a
  // placeholder, not a rule — "account" -> "accounts" holds, "opportunity" ->
  // "opportunities" does not follow from naive suffixing alone.
  return entityLogical.endsWith("y")
    ? `${entityLogical.slice(0, -1)}ies`
    : entityLogical.endsWith("s")
      ? `${entityLogical}es`
      : `${entityLogical}s`;
}

/**
 * Translate the dataflow's alternate key into a dvload upsertKey.
 *
 * `primaryKeyLogicalName` names the *key*, not its attributes — a key called
 * `asker_actid` may span `asker_act_id`. Emitting it verbatim produces a
 * mapping that fails validation ("upsertKey attribute is not mapped in
 * columns"), so it has to go through metadata. Without a resolver, or when
 * the key's attributes aren't all mapped, fall back to insert and say why:
 * an upsert that silently matches nothing would mass-create duplicates.
 */
async function resolveConflictMode(
  q: QueryMetadata,
  entityLogical: string,
  columns: ColumnMapping[],
  resolver?: DataflowMetadataResolver
): Promise<{ conflictMode: Mapping["conflictMode"]; upsertKey?: string[]; keyNote?: string }> {
  const keyLogicalName = readPrimaryKeyLogicalName(q);
  if (!keyLogicalName) return { conflictMode: "insert" };

  if (!resolver) {
    return {
      conflictMode: "insert",
      keyNote:
        `The dataflow upserts on alternate key "${keyLogicalName}". Set conflictMode=upsert ` +
        `and upsertKey to that key's attributes to match its behaviour.`,
    };
  }

  const attrs = await resolver.keyAttributes(entityLogical, keyLogicalName);
  if (attrs.length === 0) {
    return {
      conflictMode: "insert",
      keyNote:
        `The dataflow upserts on alternate key "${keyLogicalName}", which was not found on ` +
        `${entityLogical}. Left as insert.`,
    };
  }

  const mapped = new Set(columns.map((c) => c.target));
  const missing = attrs.filter((a) => !mapped.has(a));
  if (missing.length > 0) {
    return {
      conflictMode: "insert",
      keyNote:
        `The dataflow upserts on "${keyLogicalName}" (${attrs.join(", ")}), but ` +
        `${missing.join(", ")} is not among the mapped columns. Left as insert.`,
    };
  }

  return { conflictMode: "upsert", upsertKey: attrs };
}

/* -------------------------------------------------------------------------- */
/* Live resolver                                                               */
/* -------------------------------------------------------------------------- */

/**
 * A resolver backed by the Dataverse metadata endpoints, memoised per
 * instance. Building mappings for a multi-query dataflow otherwise re-reads
 * the entity list once per query and once per lookup column.
 */
export function createMetadataResolver(client: DataverseClient): DataflowMetadataResolver {
  let entityList: Promise<Map<string, string>> | undefined;
  const keyCache = new Map<string, Promise<Array<{ LogicalName: string; KeyAttributes: string[] }>>>();
  const lookupCache = new Map<string, Promise<string | undefined>>();

  const entitySets = (): Promise<Map<string, string>> =>
    (entityList ??= client
      .listEntities()
      .then((list) => new Map(list.map((e) => [e.LogicalName.toLowerCase(), e.EntitySetName]))));

  const keysFor = (entity: string): Promise<Array<{ LogicalName: string; KeyAttributes: string[] }>> => {
    let p = keyCache.get(entity);
    if (!p) {
      // A table with no alternate keys is a normal answer, not a failure.
      p = client.getEntityKeys(entity).catch(() => []);
      keyCache.set(entity, p);
    }
    return p;
  };

  return {
    async entitySetName(entityLogicalName) {
      return (await entitySets()).get(entityLogicalName.toLowerCase());
    },

    async keyAttributes(entityLogicalName, keyLogicalName) {
      const keys = await keysFor(entityLogicalName);
      const wanted = keyLogicalName.toLowerCase();
      return keys.find((k) => k.LogicalName.toLowerCase() === wanted)?.KeyAttributes ?? [];
    },

    async lookupEntitySet(entityLogicalName, attrLogicalName) {
      const cacheKey = `${entityLogicalName}:${attrLogicalName}`;
      let p = lookupCache.get(cacheKey);
      if (!p) {
        p = (async () => {
          const targets = await client
            .getLookupTargets(entityLogicalName, attrLogicalName)
            .catch(() => [] as string[]);
          // Polymorphic lookups (customerid -> account | contact) have no
          // single answer. Guessing one would bind half the rows to the
          // wrong table, so leave it unset and let the user choose.
          if (targets.length !== 1) return undefined;
          return (await entitySets()).get(targets[0].toLowerCase());
        })();
        lookupCache.set(cacheKey, p);
      }
      return p;
    },
  };
}
