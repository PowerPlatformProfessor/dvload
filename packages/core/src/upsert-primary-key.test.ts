// Upserting on the record's own id rather than an alternate key.
//
// Dataverse addresses a record by primary key as `accounts(<guid>)`; the
// `accounts(accountid='<guid>')` form is alternate-key syntax and fails when
// no alternate key exists on the id. PATCHing that URL creates the row with
// the given id if it's absent and updates it if it's present.

import { test } from "node:test";
import assert from "node:assert/strict";
import { parseMapping } from "./mapping.js";
import type { Mapping } from "./mapping.js";
import { loadRows, buildKeyExpression, isPrimaryIdKey } from "./load.js";
import type { DataverseClient, BatchOperation, BatchResultItem } from "./dataverse.js";

const GUID = "3f2504e0-4f89-11d3-9a0c-0305e82c3301";

function capture(): { client: DataverseClient; ops: BatchOperation[] } {
  const ops: BatchOperation[] = [];
  const client = {
    batch: async (batch: BatchOperation[]): Promise<BatchResultItem[]> => {
      ops.push(...batch);
      return batch.map((o) => ({ contentId: o.contentId, status: 204, ok: true }));
    },
  } as unknown as DataverseClient;
  return { client, ops };
}

function mapping(over: Partial<Mapping> = {}): Mapping {
  return parseMapping({
    schemaVersion: 1,
    name: "account",
    environmentUrl: "https://contoso.crm.dynamics.com",
    targetEntitySet: "accounts",
    sourceTable: "account",
    columns: [
      { source: "accountid", target: "accountid", kind: "uniqueidentifier", treatEmptyAsNull: true },
      { source: "Name", target: "name", kind: "string", treatEmptyAsNull: true },
    ],
    conflictMode: "upsert",
    upsertKey: ["accountid"],
    batchSize: 100,
    maxErrors: 0,
    logDir: "./logs",
    concurrency: 1,
    bypassCustomLogic: false,
    skipUnchanged: false,
    ...over,
  }) as Mapping;
}

test("upsert on a uniqueidentifier key addresses the record by bare GUID", async () => {
  const { client, ops } = capture();
  await loadRows({ mapping: mapping(), rows: [{ accountid: GUID, Name: "Contoso" }], client });

  assert.equal(ops.length, 1);
  assert.equal(ops[0].method, "PATCH");
  assert.equal(ops[0].url, `accounts(${GUID})`);
  // Explicitly NOT the alternate-key form, and never a quoted string literal.
  assert.ok(!ops[0].url.includes("accountid="));
  assert.ok(!ops[0].url.includes("'"));
});

test("the id is dropped from the body — Dataverse rejects writing the primary key", async () => {
  const { client, ops } = capture();
  await loadRows({ mapping: mapping(), rows: [{ accountid: GUID, Name: "Contoso" }], client });

  const body = ops[0].body as Record<string, unknown>;
  assert.equal(body.accountid, undefined);
  assert.equal(body.name, "Contoso");
});

test("GUID casing is normalised so the URL matches what Dataverse stores", async () => {
  const { client, ops } = capture();
  await loadRows({
    mapping: mapping(),
    rows: [{ accountid: `  ${GUID.toUpperCase()}  `, Name: "Contoso" }],
    client,
  });
  assert.equal(ops[0].url, `accounts(${GUID})`);
});

test("a malformed id fails its row instead of creating a record with a new id", async () => {
  const { client, ops } = capture();
  const result = await loadRows({
    mapping: mapping(),
    rows: [{ accountid: "not-a-guid", Name: "Contoso" }],
    client,
  });
  assert.equal(ops.length, 0);
  assert.equal(result.failed, 1);
  assert.match(result.errors[0].message, /must be a GUID/);
});

test("alternate keys are untouched — still named, quoted and encoded", async () => {
  const { client, ops } = capture();
  const m = mapping({
    upsertKey: ["accountnumber"],
    columns: [
      { source: "Num", target: "accountnumber", kind: "string", treatEmptyAsNull: true },
      { source: "Name", target: "name", kind: "string", treatEmptyAsNull: true },
    ],
  });
  await loadRows({ mapping: m, rows: [{ Num: "ACC-1", Name: "Contoso" }], client });

  assert.equal(ops[0].url, "accounts(accountnumber='ACC-1')");
  // Alternate-key attributes stay in the body: they're ordinary columns.
  assert.equal((ops[0].body as Record<string, unknown>).accountnumber, "ACC-1");
});

test("compound keys keep the alternate-key form even when one part is a GUID", () => {
  const columns = [
    { source: "a", target: "somegroupid", kind: "uniqueidentifier" as const, treatEmptyAsNull: true },
    { source: "b", target: "code", kind: "string" as const, treatEmptyAsNull: true },
  ];
  assert.equal(isPrimaryIdKey(["somegroupid", "code"], columns), false);
  assert.equal(
    buildKeyExpression(["somegroupid", "code"], columns, { a: GUID, b: "X" }),
    `somegroupid='${GUID}',code='X'`
  );
});

test("skip-if-exists uses the same bare-GUID URL plus If-None-Match", async () => {
  const { client, ops } = capture();
  await loadRows({
    mapping: mapping({ conflictMode: "skip-if-exists" }),
    rows: [{ accountid: GUID, Name: "Contoso" }],
    client,
  });
  assert.equal(ops[0].url, `accounts(${GUID})`);
  assert.equal(ops[0].headers?.["If-None-Match"], "*");
});
