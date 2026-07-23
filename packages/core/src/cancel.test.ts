// Cancellation: aborting the signal stops scheduling new batches; rows in
// completed batches keep their real outcome, the rest count as skipped, and
// the result is flagged cancelled. Sync-mode removal must not run.

import { test } from "node:test";
import assert from "node:assert/strict";
import { loadRows } from "./load.js";
import { parseMapping } from "./mapping.js";
import type { Mapping } from "./mapping.js";
import type { DataverseClient, BatchOperation, BatchResultItem } from "./dataverse.js";

function mapping(overrides: Record<string, unknown> = {}): Mapping {
  return parseMapping({
    schemaVersion: 1,
    name: "t",
    environmentUrl: "https://unit.crm.dynamics.com",
    targetEntitySet: "contacts",
    sourceTable: "T",
    columns: [{ source: "Email", target: "emailaddress1", kind: "string" }],
    conflictMode: "insert",
    batchSize: 1, // one row per batch → cancellation boundaries are per row
    ...overrides,
  }) as Mapping;
}

test("abort stops scheduling; done rows counted, rest skipped, cancelled=true", async () => {
  const controller = new AbortController();
  let calls = 0;
  const client = {
    batch: async (ops: BatchOperation[]): Promise<BatchResultItem[]> => {
      calls++;
      if (calls === 2) controller.abort(); // cancel while batch 2 is "in flight"
      return ops.map((o) => ({ contentId: o.contentId, status: 201, ok: true }));
    },
  } as unknown as DataverseClient;

  const rows = [{ Email: "a@x" }, { Email: "b@x" }, { Email: "c@x" }, { Email: "d@x" }];
  const result = await loadRows({ mapping: mapping(), rows, client, signal: controller.signal });

  assert.equal(calls, 2); // batches 3 and 4 never sent
  assert.equal(result.cancelled, true);
  assert.equal(result.created, 2); // in-flight work kept its real outcome
  assert.equal(result.skipped, 2); // unattempted rows accounted for
  assert.equal(result.failed, 0);
  assert.equal(result.total, 4);
});

test("cancelled sync run skips the removal pass", async () => {
  const controller = new AbortController();
  const queried: string[] = [];
  const client = {
    batch: async (ops: BatchOperation[]): Promise<BatchResultItem[]> => {
      controller.abort();
      return ops.map((o) => ({ contentId: o.contentId, status: 201, ok: true }));
    },
    queryAll: async (path: string): Promise<Array<Record<string, unknown>>> => {
      queried.push(path);
      return [];
    },
    getEntitySetInfo: async () => ({ logicalName: "contact", primaryIdAttribute: "contactid" }),
  } as unknown as DataverseClient;

  const result = await loadRows({
    mapping: mapping({ conflictMode: "sync", upsertKey: ["emailaddress1"] }),
    rows: [{ Email: "a@x" }, { Email: "b@x" }],
    client,
    signal: controller.signal,
  });

  assert.equal(result.cancelled, true);
  assert.equal(result.removed, 0);
  assert.deepEqual(queried, []); // no target scan → nothing deactivated/deleted
});

test("no signal → no cancelled flag", async () => {
  const client = {
    batch: async (ops: BatchOperation[]): Promise<BatchResultItem[]> =>
      ops.map((o) => ({ contentId: o.contentId, status: 201, ok: true })),
  } as unknown as DataverseClient;
  const result = await loadRows({ mapping: mapping(), rows: [{ Email: "a@x" }], client });
  assert.equal(result.cancelled, undefined);
});
