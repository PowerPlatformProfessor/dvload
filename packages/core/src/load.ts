// Orchestration: take a parsed mapping, source rows, and a Dataverse client,
// and load the data using batched OData. This is the function both the
// add-in's "Run import" button and the CLI's `run` command call.
//
// Feature notes:
//  - Lookups resolve by GUID, alternate key, or text match (with optional
//    create-if-missing and duplicate handling).
//  - skipUnchanged pre-reads each target record and strips attributes whose
//    values already match, skipping no-op rows entirely.
//  - conflictMode "sync" runs an upsert pass, then deactivates or deletes
//    target records whose keys are absent from the source.
//  - mapping.concurrency batches run in parallel (bounded pool).
//  - bypassCustomLogic / impersonateUserId become per-operation headers.

import { sourceValue, type Mapping, type ColumnMapping } from "./mapping.js";
import type { SourceRow, LoadResult, RowError, RowSuccess, ProgressFn } from "./types.js";
import { coerceRow, coerceValue, CoerceError } from "./coerce.js";
import {
  DataverseClient,
  DataverseError,
  assertLogicalName,
  formatKeyLiteral,
  type BatchOperation,
} from "./dataverse.js";

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export interface LoadOptions {
  mapping: Mapping;
  rows: SourceRow[];
  client: DataverseClient;
  /** Optional progress callback. */
  onProgress?: ProgressFn;
  /** If true, do everything except actually call Dataverse. */
  dryRun?: boolean;
  /** Resume: skip rows before this offset (they count as skipped). */
  startOffset?: number;
  /**
   * Cooperative cancellation. Batches already in flight complete (their rows
   * were sent and will exist in Dataverse); no further batches are scheduled.
   * Unattempted rows are counted as skipped and the result is flagged
   * `cancelled`. In sync mode the removal pass is also skipped.
   */
  signal?: AbortSignal;
}

export async function loadRows(opts: LoadOptions): Promise<LoadResult> {
  const { mapping, rows, client, onProgress, dryRun } = opts;
  const startOffset = clampOffset(opts.startOffset ?? 0, rows.length, mapping.batchSize);
  const startedAt = new Date().toISOString();
  const errors: RowError[] = [];
  const counters = {
    succeeded: 0,
    created: 0,
    updated: 0,
    skipped: startOffset, // resumed-past rows were handled by a previous run
    unchanged: 0,
    removed: 0,
  };

  onProgress?.({ type: "start", total: rows.length });

  // Resolve lookups in advance: collect unique source keys per lookup column,
  // resolve them once, and reuse the GUID across all rows that reference them.
  const lookupCache = await buildLookupCache(mapping, rows, client, dryRun ?? false);

  const extraHeaders = buildOpHeaders(mapping);
  const kindByTarget = new Map(mapping.columns.map((c) => [c.target, c.kind]));

  // Pre-slice the work so a bounded pool can pick batches off the list.
  const batches: Array<{ offset: number; slice: SourceRow[] }> = [];
  for (let offset = startOffset; offset < rows.length; offset += mapping.batchSize) {
    batches.push({ offset, slice: rows.slice(offset, offset + mapping.batchSize) });
  }

  let processed = startOffset;
  let nextBatch = 0;
  const completedThrough: boolean[] = new Array(batches.length).fill(false);
  let checkpointFrontier = 0;

  const emitBatchProgress = (): void => {
    onProgress?.({
      type: "batch",
      processed,
      total: rows.length,
      succeeded: counters.succeeded,
      created: counters.created,
      updated: counters.updated,
      failed: errors.length,
      skipped: counters.skipped,
    });
  };

  const runBatch = async (batchIdx: number): Promise<void> => {
    const { offset, slice } = batches[batchIdx];
    const ops: BatchOperation[] = [];
    const opIndex: number[] = []; // position in ops -> rowIndex within full `rows`
    const opKeyExpr: Array<string | undefined> = [];

    for (let i = 0; i < slice.length; i++) {
      const rowIndex = offset + i;
      const sourceRow = slice[i];
      try {
        const built = buildOperation(mapping, sourceRow, lookupCache, ops.length + 1, extraHeaders);
        if (built) {
          ops.push(built.op);
          opIndex.push(rowIndex);
          opKeyExpr.push(built.keyExpr);
        }
      } catch (e) {
        errors.push(toRowError(e, rowIndex, sourceRow));
        onProgress?.({ type: "row-error", error: errors[errors.length - 1] });
      }
    }

    // Delta detection: drop attributes that already match the target record;
    // drop whole ops with no differences.
    let effectiveOps = ops;
    let effectiveIndex = opIndex;
    if (!dryRun && mapping.skipUnchanged && (mapping.conflictMode === "upsert" || mapping.conflictMode === "sync")) {
      const filtered = await applySkipUnchanged(
        client, mapping, kindByTarget, ops, opIndex, opKeyExpr, counters
      );
      effectiveOps = filtered.ops;
      effectiveIndex = filtered.opIndex;
    }

    if (dryRun || effectiveOps.length === 0) {
      if (dryRun) {
        // Pretend the batch was a create-only insert for accounting purposes.
        counters.succeeded += effectiveOps.length;
        counters.created += effectiveOps.length;
      }
      processed += slice.length;
      emitBatchProgress();
      return;
    }

    try {
      // insert mode = plain POST creates: replaying after a 503/504 that
      // arrived mid-execution would duplicate records, so only 429 retries.
      const results = await client.batch(effectiveOps, {
        idempotent: mapping.conflictMode !== "insert",
      });
      const seen = new Set<number>();
      for (const r of results) {
        seen.add(r.contentId);
        const pos = effectiveOps.findIndex((o) => o.contentId === r.contentId);
        const rowIndex = effectiveIndex[pos];
        if (r.ok) {
          counters.succeeded++;
          // With Prefer: return=representation:
          //   201 Created → POST or PATCH-upsert that created a new record
          //   200 OK      → PATCH that found and updated an existing record
          if (r.status === 201) counters.created++;
          else if (r.status === 200 || r.status === 204) counters.updated++;
          const success: RowSuccess = {
            rowIndex,
            sourceRow: rows[rowIndex],
            status: r.status,
            ...(r.id ? { id: r.id } : {}),
          };
          onProgress?.({ type: "row-success", success });
        } else if (r.status === 412 && mapping.conflictMode === "skip-if-exists") {
          // If-None-Match: * matched an existing record. That's the
          // explicit skip the user asked for, not an error.
          counters.skipped++;
        } else {
          errors.push({
            rowIndex,
            sourceRow: rows[rowIndex],
            message: r.errorMessage ?? `HTTP ${r.status}`,
            httpStatus: r.status,
          });
          onProgress?.({ type: "row-error", error: errors[errors.length - 1] });
        }
      }
      // Any op the server returned no response for is unaccounted — surface
      // it as an error rather than letting the totals silently not add up.
      for (let i = 0; i < effectiveOps.length; i++) {
        if (seen.has(effectiveOps[i].contentId)) continue;
        const rowIndex = effectiveIndex[i];
        errors.push({
          rowIndex,
          sourceRow: rows[rowIndex],
          message: "no response returned for this operation in the $batch reply",
        });
        onProgress?.({ type: "row-error", error: errors[errors.length - 1] });
      }
    } catch (e) {
      // Whole-batch failure (network, auth). Mark every op in this batch failed.
      for (const idx of effectiveIndex) {
        errors.push(toRowError(e, idx, rows[idx]));
      }
    }

    processed += slice.length;
    emitBatchProgress();
  };

  // Bounded worker pool. JS is single-threaded, so counter mutation is safe;
  // only the HTTP requests overlap.
  const poolSize = Math.max(1, Math.min(mapping.concurrency ?? 1, batches.length || 1));
  const worker = async (): Promise<void> => {
    for (;;) {
      if (opts.signal?.aborted) return;
      if (mapping.maxErrors > 0 && errors.length >= mapping.maxErrors) return;
      const idx = nextBatch++;
      if (idx >= batches.length) return;
      await runBatch(idx);
      completedThrough[idx] = true;
      // Advance the contiguous-completion frontier for checkpointing.
      while (checkpointFrontier < batches.length && completedThrough[checkpointFrontier]) {
        checkpointFrontier++;
      }
      const committedOffset =
        checkpointFrontier < batches.length
          ? batches[checkpointFrontier].offset
          : rows.length;
      onProgress?.({ type: "checkpoint", offset: committedOffset });
    }
  };
  await Promise.all(Array.from({ length: poolSize }, worker));

  // Batches never started (maxErrors tripped or cancelled): count their rows
  // skipped rather than letting the totals silently not add up.
  const cancelled = opts.signal?.aborted ?? false;
  if (cancelled || (mapping.maxErrors > 0 && errors.length >= mapping.maxErrors)) {
    const attempted = batches
      .filter((_, i) => completedThrough[i])
      .reduce((n, b) => n + b.slice.length, 0);
    const remaining = rows.length - startOffset - attempted;
    if (remaining > 0) counters.skipped += remaining;
  }

  // Sync pass: remove target records whose keys are absent from the source.
  // Never on a cancelled run — deleting/deactivating based on a partially
  // loaded source would remove records that simply weren't reached.
  if (mapping.conflictMode === "sync" && !dryRun && !cancelled) {
    await syncRemoveMissing(client, mapping, rows, extraHeaders, counters, errors, onProgress);
  }

  const result: LoadResult = {
    ...(cancelled ? { cancelled: true } : {}),
    total: rows.length,
    succeeded: counters.succeeded,
    created: counters.created,
    updated: counters.updated,
    failed: errors.length,
    skipped: counters.skipped,
    unchanged: counters.unchanged,
    removed: counters.removed,
    startedAt,
    finishedAt: new Date().toISOString(),
    errors,
  };
  onProgress?.({ type: "done", result });
  return result;
}

/* -------------------------------------------------------------------------- */
/* Operation headers (bypass + impersonation)                                  */
/* -------------------------------------------------------------------------- */

function buildOpHeaders(mapping: Mapping): Record<string, string> {
  const h: Record<string, string> = {};
  if (mapping.bypassCustomLogic) {
    h["MSCRM.BypassCustomPluginExecution"] = "true";
    h["MSCRM.SuppressCallbackRegistrationExpanderJob"] = "true";
  }
  if (mapping.impersonateUserId) {
    h["MSCRMCallerID"] = mapping.impersonateUserId;
  }
  return h;
}

/* -------------------------------------------------------------------------- */
/* Lookup cache                                                                */
/* -------------------------------------------------------------------------- */

interface LookupEntry {
  guid?: string;
  /** Populated instead of guid when resolution failed with a specific reason. */
  failure?: string;
}

type LookupCache = Map<string /* target */, Map<string /* sourceValueKey */, LookupEntry>>;

async function buildLookupCache(
  mapping: Mapping,
  rows: SourceRow[],
  client: DataverseClient,
  dryRun: boolean
): Promise<LookupCache> {
  const cache: LookupCache = new Map();
  const lookups = mapping.columns.filter((c) => c.kind === "lookup");
  const extraHeaders = buildOpHeaders(mapping);

  for (const col of lookups) {
    const inner = new Map<string, LookupEntry>();
    cache.set(col.target, inner);

    if (dryRun) continue;
    if (col.lookupResolution === "guid") continue; // resolved per-row

    const uniques = new Set<string>();
    for (const row of rows) {
      const v = sourceValue(row, col);
      if (v === null || v === undefined || v === "") continue;
      uniques.add(String(v));
    }

    if (col.lookupResolution === "alternateKey") {
      for (const value of uniques) {
        try {
          const id = await client.resolveByKey(col.bindEntitySet!, col.keyAttribute!, value);
          if (id) inner.set(value, { guid: id });
        } catch {
          // leave unresolved; the loader errors per-row using this column
        }
      }
      continue;
    }

    // Text resolution: batched `or`-filter queries (one request per ~15
    // unique values instead of one per value), then detect duplicates and
    // optionally create missing records.
    const info = await client.getEntitySetInfo(col.bindEntitySet!);
    const uniqueList = [...uniques];
    let resolved: Map<string, string[]>;
    try {
      resolved = await client.resolveManyByText(
        col.bindEntitySet!,
        col.keyAttribute!,
        uniqueList,
        info.primaryIdAttribute
      );
    } catch (e) {
      const failure = `text lookup failed: ${e instanceof Error ? e.message : String(e)}`;
      for (const value of uniqueList) inner.set(value, { failure });
      continue;
    }
    for (const value of uniqueList) {
      const ids = resolved.get(value.toLowerCase()) ?? [];
      try {
        if (ids.length === 1) {
          inner.set(value, { guid: ids[0] });
        } else if (ids.length > 1) {
          if (col.duplicateBehavior === "first") {
            inner.set(value, { guid: ids[0] });
          } else {
            inner.set(value, {
              failure: `ambiguous text lookup: 2+ ${col.bindEntitySet} records have ${col.keyAttribute} = ${JSON.stringify(value)}`,
            });
          }
        } else if (col.createIfMissing) {
          const created = await client.create(
            col.bindEntitySet!,
            { [col.keyAttribute!]: value },
            extraHeaders
          );
          if (created.id) inner.set(value, { guid: created.id });
          else inner.set(value, { failure: "create-if-missing returned no id" });
        }
        // else: no match, no create → left unset; row errors as unresolved.
      } catch (e) {
        inner.set(value, {
          failure: `text lookup failed: ${e instanceof Error ? e.message : String(e)}`,
        });
      }
    }
  }

  return cache;
}

/* -------------------------------------------------------------------------- */
/* Delta detection (skipUnchanged)                                             */
/* -------------------------------------------------------------------------- */

async function applySkipUnchanged(
  client: DataverseClient,
  mapping: Mapping,
  kindByTarget: Map<string, string>,
  ops: BatchOperation[],
  opIndex: number[],
  opKeyExpr: Array<string | undefined>,
  counters: { skipped: number; unchanged: number }
): Promise<{ ops: BatchOperation[]; opIndex: number[] }> {
  const keptOps: BatchOperation[] = [];
  const keptIndex: number[] = [];

  // Fetch existing records in small parallel chunks to bound connection use.
  const CHUNK = 10;
  for (let start = 0; start < ops.length; start += CHUNK) {
    const chunk = ops.slice(start, start + CHUNK);
    const fetched = await Promise.all(
      chunk.map(async (op, j) => {
        const keyExpr = opKeyExpr[start + j];
        if (!keyExpr || op.method !== "PATCH") return { existing: null as Record<string, unknown> | null };
        const body = op.body as Record<string, unknown>;
        const select = Object.keys(body).filter(
          (k) => !k.includes("@") && kindByTarget.get(k) !== "lookup"
        );
        if (select.length === 0) return { existing: null };
        try {
          return { existing: await client.getRecord(mapping.targetEntitySet, keyExpr, select) };
        } catch {
          return { existing: null }; // on read failure, fall back to sending everything
        }
      })
    );

    for (let j = 0; j < chunk.length; j++) {
      const op = chunk[j];
      const existing = fetched[j].existing;
      if (!existing) {
        keptOps.push(op);
        keptIndex.push(opIndex[start + j]);
        continue;
      }
      const body = op.body as Record<string, unknown>;
      const newBody: Record<string, unknown> = {};
      let changes = 0;
      for (const [k, v] of Object.entries(body)) {
        if (k.includes("@")) {
          // Lookup binds are always sent — comparing navigation values is
          // not worth an extra metadata round-trip per column.
          newBody[k] = v;
          changes++;
          continue;
        }
        if (!valuesEqual(kindByTarget.get(k), v, existing[k])) {
          newBody[k] = v;
          changes++;
        }
      }
      if (changes === 0) {
        counters.skipped++;
        counters.unchanged++;
        continue; // drop the op entirely
      }
      keptOps.push({ ...op, body: newBody });
      keptIndex.push(opIndex[start + j]);
    }
  }

  return { ops: keptOps, opIndex: keptIndex };
}

/** Loose comparison between our wire value and the value Dataverse returned. */
export function valuesEqual(kind: string | undefined, ours: unknown, theirs: unknown): boolean {
  const oursNull = ours === null || ours === undefined;
  const theirsNull = theirs === null || theirs === undefined;
  if (oursNull || theirsNull) return oursNull === theirsNull;

  switch (kind) {
    case "datetime": {
      const a = Date.parse(String(ours));
      const b = Date.parse(String(theirs));
      return !Number.isNaN(a) && a === b;
    }
    case "multichoice": {
      const norm = (v: unknown): string =>
        String(v)
          .split(",")
          .map((s) => s.trim())
          .filter(Boolean)
          .sort()
          .join(",");
      return norm(ours) === norm(theirs);
    }
    case "integer":
    case "decimal":
    case "money":
    case "double":
    case "choice":
    case "status":
    case "state":
      return Number(ours) === Number(theirs);
    case "boolean":
      return Boolean(ours) === Boolean(theirs);
    case "uniqueidentifier":
      return String(ours).toLowerCase() === String(theirs).toLowerCase();
    default:
      return String(ours) === String(theirs);
  }
}

/* -------------------------------------------------------------------------- */
/* Sync: remove target records missing from the source                         */
/* -------------------------------------------------------------------------- */

async function syncRemoveMissing(
  client: DataverseClient,
  mapping: Mapping,
  rows: SourceRow[],
  extraHeaders: Record<string, string>,
  counters: { removed: number },
  errors: RowError[],
  onProgress?: ProgressFn
): Promise<void> {
  if (rows.length === 0) {
    throw new Error(
      "conflictMode=sync refuses to run with 0 source rows — that would " +
        `${mapping.syncAction ?? "deactivate"} every record in ${mapping.targetEntitySet}. ` +
        "If the table is intentionally empty, run the removal manually."
    );
  }
  const keyAttrs = mapping.upsertKey!;
  for (const a of keyAttrs) assertLogicalName(a, "upsertKey attribute");
  const action = mapping.syncAction ?? "deactivate";

  // Build the set of key tuples present in the source. Tuple parts are
  // joined with NUL — a printable separator (like a space) lets values
  // containing it collide across parts (["a b","c"] vs ["a","b c"]).
  const SEP = "\u0000";
  const sourceKeys = new Set<string>();
  for (const row of rows) {
    const tuple = keyAttrs.map((attr) => {
      const col = mapping.columns.find((c) => c.target === attr)!;
      const coerced = coerceValue(sourceValue(row, col), col);
      return normalizeKeyPart(coerced);
    });
    sourceKeys.add(tuple.join(SEP));
  }

  const info = await client.getEntitySetInfo(mapping.targetEntitySet);
  // Deduped: when the upsert key IS the primary id, both lists name it.
  const select = [...new Set([info.primaryIdAttribute, ...keyAttrs])].join(",");
  // For deactivate, only active records are candidates.
  const filter = action === "deactivate" ? "&$filter=statecode eq 0" : "";
  const targets = await client.queryAll(
    `${mapping.targetEntitySet}?$select=${select}${filter}`
  );

  const toRemove: string[] = [];
  for (const t of targets) {
    const tuple = keyAttrs.map((attr) => normalizeKeyPart(t[attr]));
    if (!sourceKeys.has(tuple.join(SEP))) {
      const id = String(t[info.primaryIdAttribute] ?? "");
      if (GUID_RE.test(id)) toRemove.push(id);
    }
  }

  onProgress?.({ type: "sync", checked: targets.length, toRemove: toRemove.length, removed: 0 });

  for (let offset = 0; offset < toRemove.length; offset += mapping.batchSize) {
    const slice = toRemove.slice(offset, offset + mapping.batchSize);
    const ops: BatchOperation[] = slice.map((id, i) => ({
      contentId: i + 1,
      method: action === "delete" ? "DELETE" : "PATCH",
      url: `${mapping.targetEntitySet}(${id})`,
      body: action === "delete" ? undefined : { statecode: 1 },
      headers: extraHeaders,
    }));
    try {
      const results = await client.batch(ops);
      for (const r of results) {
        if (r.ok) {
          counters.removed++;
        } else {
          errors.push({
            rowIndex: -1,
            sourceRow: { [info.primaryIdAttribute]: slice[r.contentId - 1] },
            message: `sync ${action} failed: ${r.errorMessage ?? `HTTP ${r.status}`}`,
            httpStatus: r.status,
          });
        }
      }
    } catch (e) {
      for (const id of slice) {
        errors.push({
          rowIndex: -1,
          sourceRow: { [info.primaryIdAttribute]: id },
          message: `sync ${action} failed: ${e instanceof Error ? e.message : String(e)}`,
        });
      }
    }
    onProgress?.({
      type: "sync",
      checked: targets.length,
      toRemove: toRemove.length,
      removed: counters.removed,
    });
  }
}

function normalizeKeyPart(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return String(v);
  const s = String(v).trim();
  // GUIDs are case-insensitive, and Dataverse returns them lowercase. Without
  // this, a spreadsheet holding uppercase ids would look absent from the
  // source and sync would deactivate/delete every row it was meant to keep.
  return GUID_RE.test(s) ? s.toLowerCase() : s;
}

/* -------------------------------------------------------------------------- */
/* Operation building                                                          */
/* -------------------------------------------------------------------------- */

function buildOperation(
  mapping: Mapping,
  row: SourceRow,
  lookupCache: LookupCache,
  contentId: number,
  extraHeaders: Record<string, string>
): { op: BatchOperation; keyExpr?: string } | null {
  assertLogicalName(mapping.targetEntitySet, "targetEntitySet");
  const { payload, lookups } = coerceRow(row, mapping.columns);

  // Bind lookups: foo@odata.bind = "/accounts(<guid>)"
  for (const col of lookups) {
    const value = sourceValue(row, col);
    if (value === null || value === undefined || value === "") {
      if (col.treatEmptyAsNull) {
        // A null `@odata.bind` annotation is REJECTED by the OData
        // deserializer ("undeclared property ... only has property
        // annotations but no property value"). Clearing a lookup uses the
        // plain single-valued navigation property with a null value, which
        // is only meaningful when the operation can update an existing
        // record. Plain creates have nothing to clear — omit the attribute.
        const canUpdate =
          mapping.conflictMode === "upsert" || mapping.conflictMode === "sync";
        if (canUpdate) payload[col.target] = null;
      }
      continue;
    }

    let guid: string | undefined;
    if (col.lookupResolution === "guid") {
      guid = String(value).trim().toLowerCase();
      // Cell data goes straight into a URL — must be a real GUID, nothing else.
      if (!GUID_RE.test(guid)) {
        throw new CoerceError("not a valid GUID", col.target, value);
      }
    } else {
      const entry = lookupCache.get(col.target)?.get(String(value));
      if (entry?.failure) {
        throw new CoerceError(entry.failure, col.target, value);
      }
      guid = entry?.guid;
    }

    if (!guid) {
      throw new CoerceError(
        `lookup unresolved (entitySet=${col.bindEntitySet}, key=${col.keyAttribute ?? "<n/a>"})`,
        col.target,
        value
      );
    }
    assertLogicalName(col.bindEntitySet!, "bindEntitySet");
    payload[`${col.target}@odata.bind`] = `/${col.bindEntitySet}(${guid})`;
  }

  const headers = Object.keys(extraHeaders).length > 0 ? { ...extraHeaders } : undefined;

  switch (mapping.conflictMode) {
    case "insert":
      return {
        op: {
          contentId,
          method: "POST",
          url: mapping.targetEntitySet,
          body: payload,
          ...(headers ? { headers } : {}),
        },
      };

    case "upsert":
    case "sync":
    case "skip-if-exists": {
      if (!mapping.upsertKey || mapping.upsertKey.length === 0) {
        throw new Error("conflictMode requires upsertKey");
      }
      const keyExpr = buildKeyExpression(mapping.upsertKey, mapping.columns, row);
      // The record id is already in the URL, and Dataverse rejects a write to
      // the primary key on an existing row — which would fail every update
      // while letting creates through. Alternate-key attributes stay in the
      // body: they're ordinary columns and may legitimately be set on create.
      if (isPrimaryIdKey(mapping.upsertKey, mapping.columns)) {
        delete payload[mapping.upsertKey[0]];
      }
      return {
        op: {
          contentId,
          method: "PATCH",
          url: `${mapping.targetEntitySet}(${keyExpr})`,
          body: payload,
          headers: {
            ...(headers ?? {}),
            ...(mapping.conflictMode === "skip-if-exists"
              ? { "If-None-Match": "*" } // create-only; fail if exists
              : {}),
          },
        },
        keyExpr,
      };
    }
    default:
      throw new Error(`Unhandled conflictMode: ${String(mapping.conflictMode)}`);
  }
}

/**
 * True when the upsert key is the record's own id rather than an alternate
 * key: a single column carrying a GUID. Dataverse addresses those with the
 * canonical `accounts(<guid>)` form — the named `accounts(accountid=…)` form
 * is alternate-key syntax and there is no alternate key on the primary id.
 */
export function isPrimaryIdKey(upsertKey: string[], columns: ColumnMapping[]): boolean {
  if (upsertKey.length !== 1) return false;
  const col = columns.find((c) => c.target === upsertKey[0]);
  return col?.kind === "uniqueidentifier";
}

/**
 * The key part of the record URL: `<guid>` for a primary-id upsert, or
 * `attr='value',attr2=…` for an alternate key.
 */
export function buildKeyExpression(
  upsertKey: string[],
  columns: ColumnMapping[],
  row: SourceRow
): string {
  const readKey = (attr: string): { col: ColumnMapping; value: unknown } => {
    assertLogicalName(attr, "upsertKey attribute");
    const col = columns.find((c) => c.target === attr);
    if (!col) throw new Error(`upsertKey attribute "${attr}" not present in mapping columns`);
    const value = sourceValue(row, col);
    if (value === null || value === undefined || value === "") {
      throw new Error(`upsertKey attribute "${attr}" is blank in source row`);
    }
    return { col, value };
  };

  // Upsert on the primary key: PATCH accounts(<guid>) creates the record with
  // that id when it doesn't exist, updates it when it does. The GUID goes in
  // bare — an Edm.Guid key rejects a quoted string literal.
  if (isPrimaryIdKey(upsertKey, columns)) {
    const { value } = readKey(upsertKey[0]);
    const guid = String(value).trim().toLowerCase();
    if (!GUID_RE.test(guid)) {
      throw new Error(
        `upsertKey attribute "${upsertKey[0]}" must be a GUID when upserting on the record id, got "${String(value)}"`
      );
    }
    return guid;
  }

  return upsertKey
    .map((attr) => {
      const { value } = readKey(attr);
      // formatKeyLiteral percent-encodes the value so spreadsheet data cannot
      // inject CRLF/structure into the batch request line.
      return `${attr}=${formatKeyLiteral(value)}`;
    })
    .join(",");
}

function clampOffset(offset: number, total: number, batchSize: number): number {
  if (!Number.isInteger(offset) || offset <= 0) return 0;
  if (offset >= total) return total;
  // Align to a batch boundary so resume can't split a previous batch.
  return offset - (offset % batchSize);
}

function toRowError(e: unknown, rowIndex: number, sourceRow: SourceRow): RowError {
  if (e instanceof CoerceError) {
    return { rowIndex, sourceRow, message: e.message, code: "COERCE" };
  }
  if (e instanceof DataverseError) {
    return { rowIndex, sourceRow, message: e.message, code: e.code, httpStatus: e.status };
  }
  return { rowIndex, sourceRow, message: e instanceof Error ? e.message : String(e) };
}