// Regression test: empty lookup cells. Sending `nav@odata.bind: null` is
// rejected by the OData deserializer ("undeclared property ... only has
// property annotations but no property value"), which failed every row in
// the batch. Correct behavior: omit the attribute on plain creates; clear
// via the plain navigation property (`nav: null`) on upsert/sync.

import { test } from "vitest";
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
    columns: [
      { source: "Email", target: "emailaddress1", kind: "string" },
      {
        source: "Company",
        target: "parentcustomerid_account",
        kind: "lookup",
        bindEntitySet: "accounts",
        lookupResolution: "guid",
      },
    ],
    conflictMode: "insert",
    ...overrides,
  }) as Mapping;
}

function capture(): { client: DataverseClient; ops: BatchOperation[] } {
  const ops: BatchOperation[] = [];
  const client = {
    batch: async (batch: BatchOperation[]): Promise<BatchResultItem[]> => {
      ops.push(...batch);
      return batch.map((o) => ({ contentId: o.contentId, status: 201, ok: true }));
    },
  } as unknown as DataverseClient;
  return { client, ops };
}

test("insert: empty lookup cell omits the attribute entirely", async () => {
  const { client, ops } = capture();
  await loadRows({ mapping: mapping(), rows: [{ Email: "a@x", Company: "" }], client });

  const body = ops[0].body as Record<string, unknown>;
  assert.ok(!("parentcustomerid_account@odata.bind" in body), "no null @odata.bind annotation");
  assert.ok(!("parentcustomerid_account" in body), "nothing to clear on a create");
  assert.equal(body["emailaddress1"], "a@x");
});

test("upsert: empty lookup cell clears via plain nav property null, not @odata.bind", async () => {
  const { client, ops } = capture();
  await loadRows({
    mapping: mapping({ conflictMode: "upsert", upsertKey: ["emailaddress1"] }),
    rows: [{ Email: "a@x", Company: null }],
    client,
  });

  const body = ops[0].body as Record<string, unknown>;
  assert.ok(!("parentcustomerid_account@odata.bind" in body));
  assert.equal(body["parentcustomerid_account"], null);
});

test("treatEmptyAsNull=false: empty lookup cell is omitted even on upsert", async () => {
  const m = mapping({ conflictMode: "upsert", upsertKey: ["emailaddress1"] });
  m.columns[1].treatEmptyAsNull = false;
  const { client, ops } = capture();
  await loadRows({ mapping: m, rows: [{ Email: "a@x", Company: "" }], client });

  const body = ops[0].body as Record<string, unknown>;
  assert.ok(!("parentcustomerid_account@odata.bind" in body));
  assert.ok(!("parentcustomerid_account" in body));
});

test("non-empty lookup still binds via @odata.bind", async () => {
  const { client, ops } = capture();
  const guid = "11111111-1111-1111-1111-111111111111";
  await loadRows({ mapping: mapping(), rows: [{ Email: "a@x", Company: guid }], client });

  const body = ops[0].body as Record<string, unknown>;
  assert.equal(body["parentcustomerid_account@odata.bind"], `/accounts(${guid})`);
});
