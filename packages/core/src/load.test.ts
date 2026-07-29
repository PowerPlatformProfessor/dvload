// Tests for the loader features: text lookups (with create-if-missing and
// duplicate handling), skipUnchanged delta detection, sync mode, concurrency,
// and bypass/impersonation headers.

import { test } from "vitest";
import assert from "node:assert/strict";
import { loadRows, valuesEqual } from "./load.js";
import { parseMapping } from "./mapping.js";
import type { Mapping } from "./mapping.js";
import type { DataverseClient, BatchOperation, BatchResultItem } from "./dataverse.js";

const ENV = "https://unit.crm.dynamics.com";
const GUID_A = "11111111-1111-1111-1111-111111111111";
const GUID_B = "22222222-2222-2222-2222-222222222222";

function mapping(overrides: Partial<Mapping> = {}): Mapping {
  return {
    schemaVersion: 1,
    name: "t",
    environmentUrl: ENV,
    targetEntitySet: "contacts",
    sourceTable: "T",
    columns: [{ source: "Email", target: "emailaddress1", kind: "string", treatEmptyAsNull: true }],
    conflictMode: "insert",
    batchSize: 100,
    maxErrors: 0,
    logDir: "./logs",
    concurrency: 1,
    bypassCustomLogic: false,
    skipUnchanged: false,
    ...overrides,
  };
}

/** Fake client: every op succeeds as 201; captures batches and other calls. */
function fakeClient(overrides: Partial<Record<string, unknown>> = {}): {
  client: DataverseClient;
  batches: BatchOperation[][];
} {
  const batches: BatchOperation[][] = [];
  const base = {
    batch: async (ops: BatchOperation[]): Promise<BatchResultItem[]> => {
      batches.push(ops);
      return ops.map((o) => ({ contentId: o.contentId, status: 201, ok: true }));
    },
    ...overrides,
  };
  return { client: base as unknown as DataverseClient, batches };
}

/* -------------------------------------------------------------------------- */
/* Text lookups                                                                */
/* -------------------------------------------------------------------------- */

function textLookupMapping(col: Partial<Mapping["columns"][0]> = {}): Mapping {
  return mapping({
    columns: [
      {
        source: "Company",
        target: "parentcustomerid_account",
        kind: "lookup",
        bindEntitySet: "accounts",
        lookupResolution: "text",
        keyAttribute: "name",
        treatEmptyAsNull: true,
        ...col,
      },
    ],
  });
}

test("text lookup resolves a unique match and binds it", async () => {
  const { client, batches } = fakeClient({
    getEntitySetInfo: async () => ({ logicalName: "account", primaryIdAttribute: "accountid" }),
    resolveManyByText: async (_es: string, _attr: string, values: unknown[]) =>
      new Map(values.map((v) => [String(v).toLowerCase(), [GUID_A]])),
  });

  const result = await loadRows({
    mapping: textLookupMapping(),
    rows: [{ Company: "Contoso" }],
    client,
  });

  assert.equal(result.succeeded, 1);
  const body = batches[0][0].body as Record<string, unknown>;
  assert.equal(body["parentcustomerid_account@odata.bind"], `/accounts(${GUID_A})`);
});

test("ambiguous text lookup fails the row by default, resolves with duplicateBehavior=first", async () => {
  const make = () =>
    fakeClient({
      getEntitySetInfo: async () => ({ logicalName: "account", primaryIdAttribute: "accountid" }),
      resolveManyByText: async (_es: string, _attr: string, values: unknown[]) =>
        new Map(values.map((v) => [String(v).toLowerCase(), [GUID_A, GUID_B]])),
    });

  const strict = await loadRows({
    mapping: textLookupMapping(),
    rows: [{ Company: "Contoso" }],
    client: make().client,
  });
  assert.equal(strict.failed, 1);
  assert.match(strict.errors[0].message, /ambiguous/);

  const lenient = make();
  const first = await loadRows({
    mapping: textLookupMapping({ duplicateBehavior: "first" }),
    rows: [{ Company: "Contoso" }],
    client: lenient.client,
  });
  assert.equal(first.succeeded, 1);
  const body = lenient.batches[0][0].body as Record<string, unknown>;
  assert.equal(body["parentcustomerid_account@odata.bind"], `/accounts(${GUID_A})`);
});

test("createIfMissing creates the record and uses its id; without it the row fails", async () => {
  let createdBody: Record<string, unknown> | undefined;
  const withCreate = fakeClient({
    getEntitySetInfo: async () => ({ logicalName: "account", primaryIdAttribute: "accountid" }),
    resolveManyByText: async () => new Map(),
    create: async (_set: string, body: Record<string, unknown>) => {
      createdBody = body;
      return { id: GUID_B, entityUrl: "" };
    },
  });

  const created = await loadRows({
    mapping: textLookupMapping({ createIfMissing: true }),
    rows: [{ Company: "New Corp" }],
    client: withCreate.client,
  });
  assert.equal(created.succeeded, 1);
  assert.deepEqual(createdBody, { name: "New Corp" });
  const body = withCreate.batches[0][0].body as Record<string, unknown>;
  assert.equal(body["parentcustomerid_account@odata.bind"], `/accounts(${GUID_B})`);

  const noCreate = fakeClient({
    getEntitySetInfo: async () => ({ logicalName: "account", primaryIdAttribute: "accountid" }),
    resolveManyByText: async () => new Map(),
  });
  const failed = await loadRows({
    mapping: textLookupMapping(),
    rows: [{ Company: "New Corp" }],
    client: noCreate.client,
  });
  assert.equal(failed.failed, 1);
  assert.match(failed.errors[0].message, /lookup unresolved/);
});

/* -------------------------------------------------------------------------- */
/* Headers: bypass + impersonation                                             */
/* -------------------------------------------------------------------------- */

test("bypassCustomLogic and impersonateUserId become per-op headers", async () => {
  const { client, batches } = fakeClient();
  await loadRows({
    mapping: mapping({ bypassCustomLogic: true, impersonateUserId: GUID_A }),
    rows: [{ Email: "a@x.com" }],
    client,
  });
  const headers = batches[0][0].headers!;
  assert.equal(headers["MSCRM.BypassCustomPluginExecution"], "true");
  assert.equal(headers["MSCRM.SuppressCallbackRegistrationExpanderJob"], "true");
  assert.equal(headers["MSCRMCallerID"], GUID_A);
});

/* -------------------------------------------------------------------------- */
/* skipUnchanged                                                               */
/* -------------------------------------------------------------------------- */

test("skipUnchanged drops unchanged rows and strips unchanged attributes", async () => {
  // Note: key expressions carry percent-encoded literals (formatKeyLiteral).
  const existing: Record<string, Record<string, unknown>> = {
    "emailaddress1='same%40x.com'": { emailaddress1: "same@x.com", firstname: "Ann" },
    "emailaddress1='diff%40x.com'": { emailaddress1: "diff@x.com", firstname: "OLD" },
  };
  const { client, batches } = fakeClient({
    getRecord: async (_set: string, keyExpr: string) => existing[keyExpr] ?? null,
  });

  const m = mapping({
    conflictMode: "upsert",
    upsertKey: ["emailaddress1"],
    skipUnchanged: true,
    columns: [
      { source: "Email", target: "emailaddress1", kind: "string", treatEmptyAsNull: true },
      { source: "First", target: "firstname", kind: "string", treatEmptyAsNull: true },
    ],
  });

  const result = await loadRows({
    mapping: m,
    rows: [
      { Email: "same@x.com", First: "Ann" }, // identical → skip
      { Email: "diff@x.com", First: "NEW" }, // firstname differs → send only firstname... plus key
      { Email: "new@x.com", First: "Zed" }, // 404 → full payload
    ],
    client,
  });

  assert.equal(result.unchanged, 1);
  assert.equal(result.skipped, 1);
  assert.equal(result.succeeded, 2);
  const sent = batches.flat();
  assert.equal(sent.length, 2);
  const diffOp = sent.find((o) => o.url.includes("diff"))!;
  const diffBody = diffOp.body as Record<string, unknown>;
  assert.equal(diffBody.firstname, "NEW");
  assert.equal("emailaddress1" in diffBody, false, "unchanged key attribute should be stripped");
  const newOp = sent.find((o) => o.url.includes("new"))!;
  assert.deepEqual(newOp.body, { emailaddress1: "new@x.com", firstname: "Zed" });
});

test("valuesEqual handles datetimes, numbers, and multichoice ordering", () => {
  assert.ok(valuesEqual("datetime", "2026-01-01T00:00:00.000Z", "2026-01-01T00:00:00Z"));
  assert.ok(valuesEqual("integer", 5, "5"));
  assert.ok(valuesEqual("multichoice", "2,1", "1, 2"));
  assert.ok(valuesEqual("uniqueidentifier", GUID_A.toUpperCase(), GUID_A));
  assert.ok(!valuesEqual("string", "a", "b"));
  assert.ok(valuesEqual("string", null, undefined));
  assert.ok(!valuesEqual("string", null, "x"));
});

/* -------------------------------------------------------------------------- */
/* Sync mode                                                                   */
/* -------------------------------------------------------------------------- */

test("sync deactivates target records missing from the source", async () => {
  const { client, batches } = fakeClient({
    getEntitySetInfo: async () => ({ logicalName: "contact", primaryIdAttribute: "contactid" }),
    queryAll: async () => [
      { contactid: GUID_A, emailaddress1: "keep@x.com" },
      { contactid: GUID_B, emailaddress1: "gone@x.com" },
    ],
  });

  const result = await loadRows({
    mapping: mapping({ conflictMode: "sync", upsertKey: ["emailaddress1"], syncAction: "deactivate" }),
    rows: [{ Email: "keep@x.com" }],
    client,
  });

  assert.equal(result.removed, 1);
  const removeOps = batches.flat().filter((o) => o.url.includes(GUID_B));
  assert.equal(removeOps.length, 1);
  assert.equal(removeOps[0].method, "PATCH");
  assert.deepEqual(removeOps[0].body, { statecode: 1 });
});

test("sync with syncAction=delete issues DELETE ops", async () => {
  const { client, batches } = fakeClient({
    getEntitySetInfo: async () => ({ logicalName: "contact", primaryIdAttribute: "contactid" }),
    queryAll: async () => [{ contactid: GUID_B, emailaddress1: "gone@x.com" }],
  });

  await loadRows({
    mapping: mapping({ conflictMode: "sync", upsertKey: ["emailaddress1"], syncAction: "delete" }),
    rows: [{ Email: "keep@x.com" }],
    client,
  });

  const del = batches.flat().find((o) => o.method === "DELETE")!;
  assert.ok(del, "expected a DELETE op");
  assert.equal(del.url, `contacts(${GUID_B})`);
});

test("sync refuses to run with zero source rows", async () => {
  const { client } = fakeClient({
    getEntitySetInfo: async () => ({ logicalName: "contact", primaryIdAttribute: "contactid" }),
    queryAll: async () => [{ contactid: GUID_B, emailaddress1: "gone@x.com" }],
  });
  await assert.rejects(
    () =>
      loadRows({
        mapping: mapping({ conflictMode: "sync", upsertKey: ["emailaddress1"] }),
        rows: [],
        client,
      }),
    /refuses to run with 0 source rows/
  );
});

/* -------------------------------------------------------------------------- */
/* Concurrency + resume                                                        */
/* -------------------------------------------------------------------------- */

test("concurrent batches process every row exactly once", async () => {
  const { client, batches } = fakeClient();
  const rows = Array.from({ length: 25 }, (_, i) => ({ Email: `u${i}@x.com` }));
  const result = await loadRows({
    mapping: mapping({ batchSize: 5, concurrency: 3 }),
    rows,
    client,
  });
  assert.equal(result.succeeded, 25);
  assert.equal(batches.length, 5);
  const allEmails = batches
    .flat()
    .map((o) => (o.body as Record<string, unknown>).emailaddress1)
    .sort();
  assert.equal(new Set(allEmails).size, 25);
});

test("startOffset resumes past already-processed rows", async () => {
  const { client, batches } = fakeClient();
  const rows = Array.from({ length: 10 }, (_, i) => ({ Email: `u${i}@x.com` }));
  const result = await loadRows({
    mapping: mapping({ batchSize: 5 }),
    rows,
    client,
    startOffset: 5,
  });
  assert.equal(result.skipped, 5);
  assert.equal(result.succeeded, 5);
  assert.equal(batches.length, 1);
});

/* -------------------------------------------------------------------------- */
/* Mapping parse of new fields                                                 */
/* -------------------------------------------------------------------------- */

test("parseMapping validates the new fields", () => {
  const base = {
    schemaVersion: 1,
    name: "t",
    environmentUrl: ENV,
    targetEntitySet: "contacts",
    sourceTable: "T",
    columns: [{ source: "Email", target: "emailaddress1", kind: "string" }],
  };

  assert.equal(parseMapping({ ...base, conflictMode: "sync" }).syncAction, "deactivate");
  assert.equal(parseMapping(base).concurrency, 1);
  assert.throws(() => parseMapping({ ...base, concurrency: 99 }), /concurrency/);
  assert.throws(() => parseMapping({ ...base, impersonateUserId: "bob" }), /systemuser GUID/);
  assert.throws(() => parseMapping({ ...base, syncAction: "obliterate" }), /syncAction/);
  const withText = parseMapping({
    ...base,
    columns: [
      {
        source: "Company",
        target: "parentcustomerid_account",
        kind: "lookup",
        bindEntitySet: "accounts",
        lookupResolution: "text",
        keyAttribute: "name",
        createIfMissing: true,
        duplicateBehavior: "first",
      },
    ],
  });
  assert.equal(withText.columns[0].createIfMissing, true);
  assert.equal(withText.columns[0].duplicateBehavior, "first");
});
