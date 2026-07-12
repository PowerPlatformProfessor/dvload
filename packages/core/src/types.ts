// Shared, runtime-agnostic types. Keep this file dependency-free.

/** A single row read from an Excel table, keyed by column header. */
export type SourceRow = Record<string, unknown>;

/** A row prepared for Dataverse: keyed by logical/navigation name. */
export type DataversePayload = Record<string, unknown>;

export type DataverseFieldKind =
  | "string"
  | "memo"
  | "integer"
  | "decimal"
  | "money"
  | "double"
  | "boolean"
  | "datetime" // UTC, ISO 8601
  | "dateonly"
  | "uniqueidentifier"
  | "lookup" // requires bindNavigationProperty + bindEntitySet
  | "choice" // single-select option set; integer value
  | "multichoice" // multi-select; comma-joined integer values
  | "status"
  | "state";

/**
 * What to do when a row already exists (matched by alternate key).
 * "sync" = upsert + remove (deactivate or delete, per syncAction) any target
 * record whose key is absent from the source table.
 */
export type ConflictMode = "insert" | "upsert" | "skip-if-exists" | "sync";

/** What "sync" does with target records missing from the source. */
export type SyncAction = "deactivate" | "delete";

/** A successfully processed row from a load run. */
export interface RowSuccess {
  rowIndex: number;
  sourceRow: SourceRow;
  /** HTTP status returned for this row (201 Created, 200 Updated). */
  status: number;
  /** Created/updated entity GUID. */
  id?: string;
}

/** A failed row from a load run. */
export interface RowError {
  rowIndex: number; // 0-based index in the source table
  sourceRow: SourceRow;
  message: string;
  code?: string;
  httpStatus?: number;
}

/** Result of a load run. */
export interface LoadResult {
  total: number;
  succeeded: number;
  /** Subset of `succeeded` — rows the server returned 201 Created for. */
  created: number;
  /** Subset of `succeeded` — rows the server returned 200 OK for (PATCH found an existing row). */
  updated: number;
  failed: number;
  /**
   * Rows we deliberately didn't load. Currently this counts:
   *   - rows skipped because conflictMode=skip-if-exists matched an existing row (412)
   *   - rows we didn't attempt because maxErrors stopped the run early
   */
  skipped: number;
  /**
   * Rows skipped because skipUnchanged found no differing attributes.
   * Subset of `skipped`.
   */
  unchanged: number;
  /** conflictMode=sync only: target records deactivated or deleted. */
  removed: number;
  startedAt: string;
  finishedAt: string;
  errors: RowError[];
}

/** Progress callback shape used by the engine. */
export type ProgressFn = (event: ProgressEvent) => void;
export type ProgressEvent =
  | { type: "start"; total: number }
  | {
      type: "batch";
      processed: number;
      total: number;
      succeeded: number;
      created: number;
      updated: number;
      failed: number;
      skipped: number;
    }
  | { type: "row-success"; success: RowSuccess }
  | { type: "row-error"; error: RowError }
  /** conflictMode=sync: progress of the remove-missing pass. */
  | { type: "sync"; checked: number; toRemove: number; removed: number }
  /** Emitted after each committed batch; offset = next row offset to process. */
  | { type: "checkpoint"; offset: number }
  | { type: "done"; result: LoadResult };
