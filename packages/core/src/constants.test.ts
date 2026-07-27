// Tests for constant-value columns: schema parsing, sourceValue resolution,
// payload coercion, and lookup binding of fixed values.

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  parseMapping,
  validateMapping,
  validateColumn,
  mappingWarnings,
  sourceValue,
  MappingParseError,
} from "./mapping.js";
import type { Mapping, ColumnMapping } from "./mapping.js";
import { coerceRow } from "./coerce.js";
import { writeRowsToBuffer, readTableFromBuffer } from "./xlsx-reader.js";
import { loadRows } from "./load.js";
import type { DataverseClient, BatchOperation, BatchResultItem } from "./dataverse.js";

const GUID_A = "11111111-1111-1111-1111-111111111111";

function baseMapping(columns: unknown[]): unknown {
  return {
    schemaVersion: 1,
    name: "t",
    environmentUrl: "https://unit.crm.dynamics.com",
    targetEntitySet: "contacts",
    sourceTable: "T",
    columns,
    conflictMode: "insert",
  };
}

/* -------------------------------------------------------------------------- */
/* Parsing                                                                     */
/* -------------------------------------------------------------------------- */

test("parseMapping accepts a constant column without source", () => {
  const m = parseMapping(
    baseMapping([
      { source: "Email", target: "emailaddress1", kind: "string" },
      { constant: "Imported", target: "description", kind: "string" },
    ])
  );
  assert.equal(m.columns[1].source, undefined);
  assert.equal(m.columns[1].constant, "Imported");
});

test("parseMapping rejects a column with both source and constant", () => {
  assert.throws(
    () =>
      parseMapping(
        baseMapping([{ source: "A", constant: "x", target: "description", kind: "string" }])
      ),
    MappingParseError
  );
});

test("parseMapping rejects a column with neither source nor constant", () => {
  assert.throws(
    () => parseMapping(baseMapping([{ target: "description", kind: "string" }])),
    MappingParseError
  );
});

test("parseMapping rejects non-scalar constants", () => {
  assert.throws(
    () =>
      parseMapping(baseMapping([{ constant: { a: 1 }, target: "description", kind: "string" }])),
    MappingParseError
  );
});

/* -------------------------------------------------------------------------- */
/* sourceValue + coercion                                                      */
/* -------------------------------------------------------------------------- */

test("sourceValue returns the cell for source columns and the constant otherwise", () => {
  const src: ColumnMapping = { source: "A", target: "x", kind: "string", treatEmptyAsNull: true };
  const cst: ColumnMapping = { constant: 42, target: "y", kind: "integer", treatEmptyAsNull: true };
  const row = { A: "hello" };
  assert.equal(sourceValue(row, src), "hello");
  assert.equal(sourceValue(row, cst), 42);
});

test("coerceRow includes constants in every payload", () => {
  const columns: ColumnMapping[] = [
    { source: "Email", target: "emailaddress1", kind: "string", treatEmptyAsNull: true },
    { constant: 3, target: "priority", kind: "integer", treatEmptyAsNull: true },
  ];
  const { payload } = coerceRow({ Email: "a@b.c" }, columns);
  assert.deepEqual(payload, { emailaddress1: "a@b.c", priority: 3 });
});

/* -------------------------------------------------------------------------- */
/* Constant lookups (the fixed-owner scenario)                                 */
/* -------------------------------------------------------------------------- */

test("a constant GUID lookup binds on every record", async () => {
  const batches: BatchOperation[][] = [];
  const client = {
    batch: async (ops: BatchOperation[]): Promise<BatchResultItem[]> => {
      batches.push(ops);
      return ops.map((o) => ({ contentId: o.contentId, status: 201, ok: true }));
    },
  } as unknown as DataverseClient;

  const m = parseMapping(
    baseMapping([
      { source: "Email", target: "emailaddress1", kind: "string" },
      {
        constant: GUID_A,
        target: "ownerid",
        kind: "lookup",
        bindEntitySet: "systemusers",
        lookupResolution: "guid",
      },
    ])
  ) as Mapping;

  const result = await loadRows({
    mapping: m,
    rows: [{ Email: "a@b.c" }, { Email: "d@e.f" }],
    client,
  });

  assert.equal(result.failed, 0);
  for (const op of batches.flat()) {
    assert.equal(
      (op.body as Record<string, unknown>)["ownerid@odata.bind"],
      `/systemusers(${GUID_A})`
    );
  }
});

/* -------------------------------------------------------------------------- */
/* validateMapping                                                             */
/* -------------------------------------------------------------------------- */

test("mappingWarnings flags overriddencreatedon with upsert, not with insert", () => {
  const cols = [
    { source: "Email", target: "emailaddress1", kind: "string" },
    { source: "Created", target: "overriddencreatedon", kind: "datetime" },
  ];

  const insert = parseMapping(baseMapping(cols)) as Mapping;
  const insertWarnings = mappingWarnings(insert);
  assert.ok(!insertWarnings.some((w) => w.includes("CREATE only")));
  assert.ok(insertWarnings.some((w) => w.includes("Override Created On"))); // privilege note always shows

  const upsert = parseMapping({
    ...(baseMapping(cols) as Record<string, unknown>),
    conflictMode: "upsert",
    upsertKey: ["emailaddress1"],
  }) as Mapping;
  assert.ok(mappingWarnings(upsert).some((w) => w.includes("CREATE only")));

  const plain = parseMapping(
    baseMapping([{ source: "Email", target: "emailaddress1", kind: "string" }])
  ) as Mapping;
  assert.equal(mappingWarnings(plain).length, 0);
});

test("validateMapping flags a non-GUID constant on a guid-resolution lookup", () => {
  const m = parseMapping(
    baseMapping([
      {
        constant: "not-a-guid",
        target: "ownerid",
        kind: "lookup",
        bindEntitySet: "systemusers",
        lookupResolution: "guid",
      },
    ])
  ) as Mapping;
  const errors = validateMapping(m);
  assert.ok(errors.some((e) => e.includes("must be a GUID")));
});

test("writeRowsToBuffer round-trips through readTableFromBuffer", async () => {
  // The add-in's failed-rows download must be re-importable, which means the
  // bytes have to parse back with the same headers and values.
  const headers = ["Name", "Email", "Age"];
  const rows = [
    { Name: "Ada", Email: "ada@example.com", Age: 36 },
    { Name: "Grace", Email: "grace@example.com", Age: 45 },
  ];
  const buf = await writeRowsToBuffer(headers, rows, "FailedRows");
  const back = await readTableFromBuffer(buf, { tableName: "FailedRows" });
  assert.deepEqual(back.headers, headers);
  assert.equal(back.rows.length, 2);
  assert.equal(back.rows[0].Name, "Ada");
  assert.equal(back.rows[1].Age, 45);
});

test("readTableFromBuffer without a table name falls back to the first table", async () => {
  // A file picked by a user may not have a named table, and they shouldn't
  // have to know its name if it does.
  const buf = await writeRowsToBuffer(["A", "B"], [{ A: 1, B: 2 }], "SomeOtherName");
  const back = await readTableFromBuffer(buf, {});
  assert.deepEqual(back.headers, ["A", "B"]);
  assert.equal(back.rows.length, 1);
  assert.equal(back.rows[0].A, 1);
});

test("validateColumn reports the same lookup errors validateMapping does", () => {
  // The add-in renders these per row, so the two must not drift apart.
  const bad = {
    constant: "",
    target: "ownerid",
    kind: "lookup" as const,
    bindEntitySet: "systemusers",
    treatEmptyAsNull: true,
  };
  const perColumn = validateColumn(bad);
  assert.ok(perColumn.some((e) => e.includes("missing lookupResolution")));

  const whole = validateMapping(parseMapping(baseMapping([bad])) as Mapping);
  for (const e of perColumn) assert.ok(whole.includes(e), `validateMapping lost: ${e}`);
});

test("validateColumn accepts a text lookup and rejects its options elsewhere", () => {
  assert.deepEqual(
    validateColumn({
      source: "Owner",
      target: "ownerid",
      kind: "lookup",
      bindEntitySet: "systemusers",
      lookupResolution: "text",
      keyAttribute: "domainname",
      createIfMissing: false,
      duplicateBehavior: "first",
      treatEmptyAsNull: true,
    }),
    []
  );

  // keyAttribute is mandatory for text/alternateKey…
  assert.ok(
    validateColumn({
      source: "Owner",
      target: "ownerid",
      kind: "lookup",
      bindEntitySet: "systemusers",
      lookupResolution: "text",
      treatEmptyAsNull: true,
    }).some((e) => e.includes("no keyAttribute set"))
  );

  // …and duplicateBehavior is meaningless outside text resolution.
  assert.ok(
    validateColumn({
      source: "Owner",
      target: "ownerid",
      kind: "lookup",
      bindEntitySet: "systemusers",
      lookupResolution: "guid",
      duplicateBehavior: "first",
      treatEmptyAsNull: true,
    }).some((e) => e.includes("duplicateBehavior requires lookupResolution=text"))
  );
});
