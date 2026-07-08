// Orchestration: take a parsed mapping + source rows + a Dataverse client,
// and load the data using batched OData. This is the function both the
// add-in's "Run import" button and the CLI's `run` command call.

import type { Mapping, ColumnMapping } from "./mapping.js";
import type { SourceRow, LoadResult, RowError, RowSuccess, ProgressFn } from "./types.js";
import { coerceRow, CoerceError } from "./coerce.js";
import { DataverseClient, DataverseError, type BatchOperation } from "./dataverse.js";

export interface LoadOptions {
  mapping: Mapping;
  rows: SourceRow[];
  client: DataverseClient;
  /** Optional progress callback. */
  onProgress?: ProgressFn;
  /** If true, do everything except actually call Dataverse. */
  dryRun?: boolean;
}

export async function loadRows(opts: LoadOptions): Promise<LoadResult> {
  const { mapping, rows, client, onProgress, dryRun } = opts;
  const startedAt = new Date().toISOString();
  const errors: RowError[] = [];
  let succeeded = 0;
  let created = 0;
  let updated = 0;
  let skipped = 0;

  onProgress?.({ type: "start", total: rows.length });

  // Resolve lookups in advance: collect unique source keys per lookup column,
  // resolve them once, and reuse the GUID across all rows that reference them.
  const lookupCache = await buildLookupCache(mapping, rows, client, dryRun ?? false);

  // Slice rows into changesets of mapping.batchSize.
  for (let offset = 0; offset < rows.length; offset += mapping.batchSize) {
    if (mapping.maxErrors > 0 && errors.length >= mapping.maxErrors) {
      // Stop scheduling more batches.
      const remaining = rows.length - offset;
      skipped += remaining;
      break;
    }

    const slice = rows.slice(offset, offset + mapping.batchSize);
    const ops: BatchOperation[] = [];
    const opIndex: number[] = []; // contentId -> rowIndex within full `rows`

    for (let i = 0; i < slice.length; i++) {
      const rowIndex = offset + i;
      const sourceRow = slice[i];
      try {
        const op = buildOperation(mapping, sourceRow, lookupCache, rowIndex, ops.length + 1);
        if (op) {
          ops.push(op);
          opIndex.push(rowIndex);
        }
      } catch (e) {
        errors.push(toRowError(e, rowIndex, sourceRow));
        onProgress?.({ type: "row-error", error: errors[errors.length - 1] });
      }
    }

    if (dryRun || ops.length === 0) {
      // Pretend the batch was a create-only insert for accounting purposes.
      succeeded += ops.length;
      created += ops.length;
      onProgress?.({
        type: "batch",
        processed: offset + slice.length,
        total: rows.length,
        succeeded,
        created,
        updated,
        failed: errors.length,
        skipped,
      });
      continue;
    }

    try {
      const results = await client.batch(ops);
      for (const r of results) {
        const rowIndex = opIndex[r.contentId - 1];
        if (r.ok) {
          succeeded++;
          // With Prefer: return=representation:
          //   201 Created → POST or PATCH-upsert that created a new record
          //   200 OK      → PATCH that found and updated an existing record
          //   204 No Content → PATCH/DELETE without representation (we don't ask for this)
          if (r.status === 201) created++;
          else if (r.status === 200 || r.status === 204) updated++;
          const success: RowSuccess = { rowIndex, sourceRow: rows[rowIndex], status: r.status, ...(r.id ? { id: r.id } : {}) };
          onProgress?.({ type: "row-success", success });
        } else if (
          r.status === 412 &&
          mapping.conflictMode === "skip-if-exists"
        ) {
          // If-None-Match: * matched an existing record. That's the
          // explicit skip the user asked for, not an error.
          skipped++;
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
    } catch (e) {
      // Whole-batch failure (network, auth). Mark every op in this batch failed.
      for (const idx of opIndex) {
        errors.push(toRowError(e, idx, rows[idx]));
      }
    }

    onProgress?.({
      type: "batch",
      processed: offset + slice.length,
      total: rows.length,
      succeeded,
      created,
      updated,
      failed: errors.length,
      skipped,
    });
  }

  const result: LoadResult = {
    total: rows.length,
    succeeded,
    created,
    updated,
    failed: errors.length,
    skipped,
    startedAt,
    finishedAt: new Date().toISOString(),
    errors,
  };
  onProgress?.({ type: "done", result });
  return result;
}

/* -------------------------------------------------------------------------- */
/* Internals                                                                   */
/* -------------------------------------------------------------------------- */

type LookupCache = Map<string /* target */, Map<string /* sourceValueKey */, string /* guid */>>;

async function buildLookupCache(
  mapping: Mapping,
  rows: SourceRow[],
  client: DataverseClient,
  dryRun: boolean
): Promise<LookupCache> {
  const cache: LookupCache = new Map();
  const lookups = mapping.columns.filter((c) => c.kind === "lookup");

  for (const col of lookups) {
    const inner = new Map<string, string>();
    cache.set(col.target, inner);

    if (dryRun) continue;
    if (col.lookupResolution === "guid") continue; // resolved per-row

    const uniques = new Set<string>();
    for (const row of rows) {
      const v = row[col.source];
      if (v === null || v === undefined || v === "") continue;
      uniques.add(String(v));
    }

    for (const value of uniques) {
      try {
        const id = await client.resolveByKey(col.bindEntitySet!, col.keyAttribute!, value);
        if (id) inner.set(value, id);
      } catch {
        // leave unresolved; load.ts will error per-row using this column
      }
    }
  }

  return cache;
}

function buildOperation(
  mapping: Mapping,
  row: SourceRow,
  lookupCache: LookupCache,
  _rowIndex: number,
  contentId: number
): BatchOperation | null {
  const { payload, lookups } = coerceRow(row, mapping.columns);

  // Bind lookups: foo@odata.bind = "/accounts(<guid>)"
  for (const col of lookups) {
    const value = row[col.source];
    if (value === null || value === undefined || value === "") {
      if (col.treatEmptyAsNull) {
        // Setting a lookup to null clears it via the navigation property
        payload[`${col.target}@odata.bind`] = null;
      }
      continue;
    }

    let guid: string | undefined;
    if (col.lookupResolution === "guid") {
      guid = String(value).trim().toLowerCase();
    } else {
      guid = lookupCache.get(col.target)?.get(String(value));
    }

    if (!guid) {
      throw new CoerceError(
        `lookup unresolved (entitySet=${col.bindEntitySet}, key=${col.keyAttribute ?? "<n/a>"})`,
        col.target,
        value
      );
    }
    payload[`${col.target}@odata.bind`] = `/${col.bindEntitySet}(${guid})`;
  }

  switch (mapping.conflictMode) {
    case "insert":
      return {
        contentId,
        method: "POST",
        url: mapping.targetEntitySet,
        body: payload,
      };

    case "upsert":
    case "skip-if-exists": {
      if (!mapping.upsertKey || mapping.upsertKey.length === 0) {
        throw new Error("conflictMode requires upsertKey");
      }
      const keyExpr = buildKeyExpression(mapping.upsertKey, mapping.columns, row);
      return {
        contentId,
        method: "PATCH",
        url: `${mapping.targetEntitySet}(${keyExpr})`,
        body: payload,
        headers:
          mapping.conflictMode === "skip-if-exists"
            ? { "If-None-Match": "*" } // create-only; fail if exists
            : {},
      };
    }
    default:
      throw new Error(`Unhandled conflictMode: ${String(mapping.conflictMode)}`);
  }
}

function buildKeyExpression(
  upsertKey: string[],
  columns: ColumnMapping[],
  row: SourceRow
): string {
  const parts = upsertKey.map((attr) => {
    const col = columns.find((c) => c.target === attr);
    if (!col) throw new Error(`upsertKey attribute "${attr}" not present in mapping columns`);
    const v = row[col.source];
    if (v === null || v === undefined || v === "") {
      throw new Error(`upsertKey attribute "${attr}" is blank in source row`);
    }
    return `${attr}=${typeof v === "number" ? v : `'${String(v).replace(/'/g, "''")}'`}`;
  });
  return parts.join(",");
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
