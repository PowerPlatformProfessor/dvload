// Tests for the $batch codec, input validation, and load accounting.

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  DataverseClient,
  assertLogicalName,
  assertGuid,
  formatKeyLiteral,
  type BatchOperation,
} from "./dataverse.js";
import { loadRows } from "./load.js";
import type { Mapping } from "./mapping.js";

const ENV = "https://unit.crm.dynamics.com";
const CRLF = "\r\n";

function clientWithFetch(fetchImpl: typeof fetch): DataverseClient {
  return new DataverseClient({
    environmentUrl: ENV,
    getToken: async () => "tok",
    fetch: fetchImpl,
  });
}

/** Build a Dataverse-style multipart batch response. */
function multipartResponse(parts: Array<{ contentId: number; status: string; headers?: string[]; body?: string }>): string {
  const lines: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    const cs = `changesetresponse_${i + 1}`;
    lines.push(
      `--batchresponse_x`,
      `Content-Type: multipart/mixed; boundary=${cs}`,
      ``,
      `--${cs}`,
      `Content-Type: application/http`,
      `Content-Transfer-Encoding: binary`,
      `Content-ID: ${p.contentId}`,
      ``,
      `HTTP/1.1 ${p.status}`,
      ...(p.headers ?? []),
      `Content-Type: application/json; odata.metadata=minimal`,
      ``,
      p.body ?? ``,
      `--${cs}--`
    );
  }
  lines.push(`--batchresponse_x--`, ``);
  return lines.join(CRLF);
}

/* -------------------------------------------------------------------------- */
/* create() — reading the new record's id back                                 */
/* -------------------------------------------------------------------------- */

// These go through the real client rather than a stubbed `create`, because
// that is exactly where the bug lived: load.test.ts asserted the contract
// ("createIfMissing uses the returned id") against a mock that always
// returned one, while the implementation could only read an OData-EntityId
// header — which Dataverse omits whenever `Prefer: return=representation` is
// sent, i.e. always. Every create silently produced an empty id.

const NEW_GUID = "11111111-2222-3333-4444-555555555555";

test("create reads the id from the OData-EntityId header (204 No Content)", async () => {
  const client = clientWithFetch(async () =>
    new Response(null, {
      status: 204,
      headers: { "OData-EntityId": `${ENV}/api/data/v9.2/accounts(${NEW_GUID})` },
    })
  );
  const r = await client.create("accounts", { name: "X" });
  assert.equal(r.id, NEW_GUID);
});

test("create reads the id from the body when the server returns a representation", async () => {
  // The real-world shape: 201 + the record, and NO OData-EntityId header.
  const client = clientWithFetch(async () =>
    new Response(JSON.stringify({ accountid: NEW_GUID, name: "X" }), {
      status: 201,
      headers: { "Content-Type": "application/json" },
    })
  );
  const r = await client.create("accounts", { name: "X" }, undefined, "accountid");
  assert.equal(r.id, NEW_GUID);
});

test("create falls back to the @odata.id annotation", async () => {
  const client = clientWithFetch(async () =>
    new Response(
      JSON.stringify({ "@odata.id": `${ENV}/api/data/v9.2/teams(${NEW_GUID})`, name: "X" }),
      { status: 201, headers: { "Content-Type": "application/json" } }
    )
  );
  // No primaryIdAttribute passed: the annotation is the only source.
  const r = await client.create("teams", { name: "X" });
  assert.equal(r.id, NEW_GUID);
});

test("create never guesses an id from a lookup field", async () => {
  // A created record is full of *id properties. Binding a lookup to
  // `ownerid` instead of the primary key would corrupt data silently, so an
  // unreadable id must stay empty and let the caller fail loudly.
  const client = clientWithFetch(async () =>
    new Response(
      JSON.stringify({ ownerid: NEW_GUID, _createdby_value: NEW_GUID, name: "X" }),
      { status: 201, headers: { "Content-Type": "application/json" } }
    )
  );
  const r = await client.create("accounts", { name: "X" }, undefined, "accountid");
  assert.equal(r.id, "");
});

test("create ignores a non-GUID value in the primary id attribute", async () => {
  const client = clientWithFetch(async () =>
    new Response(JSON.stringify({ accountid: "not-a-guid", name: "X" }), {
      status: 201,
      headers: { "Content-Type": "application/json" },
    })
  );
  const r = await client.create("accounts", { name: "X" }, undefined, "accountid");
  assert.equal(r.id, "");
});

/* -------------------------------------------------------------------------- */
/* Validation helpers                                                          */
/* -------------------------------------------------------------------------- */

test("assertLogicalName accepts valid names, rejects injection attempts", () => {
  assert.equal(assertLogicalName("contacts", "x"), "contacts");
  assert.equal(assertLogicalName("new_my_table2", "x"), "new_my_table2");
  for (const bad of ["", "conta cts", "contacts)", "a'b", "a\r\nb", "a/b", "a?b=c"]) {
    assert.throws(() => assertLogicalName(bad, "x"), /not a valid Dataverse logical name/);
  }
});

test("assertGuid rejects non-GUIDs", () => {
  assert.equal(
    assertGuid("11111111-2222-3333-4444-555555555555", "x"),
    "11111111-2222-3333-4444-555555555555"
  );
  for (const bad of ["", "not-a-guid", "11111111-2222-3333-4444-55555555555Z", "1)\r\nPATCH x"]) {
    assert.throws(() => assertGuid(bad, "x"), /not a valid GUID/);
  }
});

test("formatKeyLiteral escapes quotes and percent-encodes structure", () => {
  assert.equal(formatKeyLiteral(42), "42");
  assert.equal(formatKeyLiteral(true), "true");
  assert.equal(formatKeyLiteral("plain"), "'plain'");
  // Doubled quote, then percent-encoded.
  assert.equal(formatKeyLiteral("O'Brien"), "'O%27%27Brien'");
  // CRLF, spaces, parens, and quotes cannot survive into a request line:
  // everything between the delimiting quotes must be structure-free.
  const evil = formatKeyLiteral("x')\r\nDELETE https://evil HTTP/1.1");
  assert.match(evil, /^'[^'\r\n() ]*'$/, `unsafe characters left in: ${evil}`);
});

/* -------------------------------------------------------------------------- */
/* Batch request building                                                      */
/* -------------------------------------------------------------------------- */

test("batch wraps each operation in its own changeset", async () => {
  let sentBody = "";
  const client = clientWithFetch((async (_url: unknown, init?: RequestInit) => {
    sentBody = String(init?.body);
    return new Response(multipartResponse([{ contentId: 1, status: "204 No Content" }]), {
      status: 200,
    });
  }) as typeof fetch);

  const ops: BatchOperation[] = [
    { contentId: 1, method: "POST", url: "contacts", body: { a: 1 } },
    { contentId: 2, method: "PATCH", url: "contacts(emailaddress1='x')", body: { a: 2 } },
  ];
  await client.batch(ops);

  const changesetOpens = sentBody.match(/boundary=changeset_/g) ?? [];
  assert.equal(changesetOpens.length, 2, "expected one changeset per operation");
  assert.match(sentBody, /Content-ID: 1/);
  assert.match(sentBody, /Content-ID: 2/);
});

test("batch refuses URLs with CRLF or whitespace (request smuggling)", async () => {
  const client = clientWithFetch((async () => new Response("", { status: 200 })) as typeof fetch);
  const evil: BatchOperation[] = [
    {
      contentId: 1,
      method: "PATCH",
      url: "contacts(key='x')\r\nDELETE https://evil HTTP/1.1",
      body: {},
    },
  ];
  await assert.rejects(() => client.batch(evil), /URL contains whitespace\/CRLF/);
});

test("batch refuses headers with CRLF", async () => {
  const client = clientWithFetch((async () => new Response("", { status: 200 })) as typeof fetch);
  const evil: BatchOperation[] = [
    { contentId: 1, method: "POST", url: "contacts", body: {}, headers: { "If-Match": "*\r\nX: y" } },
  ];
  await assert.rejects(() => client.batch(evil), /header contains CRLF/);
});

/* -------------------------------------------------------------------------- */
/* Batch response parsing                                                      */
/* -------------------------------------------------------------------------- */

test("parses multi-changeset responses with mixed outcomes", async () => {
  const guid = "11111111-2222-3333-4444-555555555555";
  const responseText = multipartResponse([
    {
      contentId: 1,
      status: "201 Created",
      headers: [`OData-EntityId: ${ENV}/api/data/v9.2/contacts(${guid})`],
      body: JSON.stringify({ contactid: guid }),
    },
    {
      contentId: 2,
      status: "412 Precondition Failed",
      body: JSON.stringify({ error: { message: "A record with matching key values already exists." } }),
    },
    { contentId: 3, status: "200 OK", body: JSON.stringify({ contactid: guid }) },
  ]);
  const client = clientWithFetch(
    (async () => new Response(responseText, { status: 200 })) as typeof fetch
  );

  const results = await client.batch([
    { contentId: 1, method: "POST", url: "contacts", body: {} },
    { contentId: 2, method: "PATCH", url: "contacts(k='a')", body: {}, headers: { "If-None-Match": "*" } },
    { contentId: 3, method: "PATCH", url: "contacts(k='b')", body: {} },
  ]);

  assert.equal(results.length, 3);
  assert.deepEqual(
    results.map((r) => [r.contentId, r.status, r.ok]),
    [
      [1, 201, true],
      [2, 412, false],
      [3, 200, true],
    ]
  );
  assert.equal(results[0].id, guid);
  assert.match(results[1].errorMessage ?? "", /already exists/);
});

/* -------------------------------------------------------------------------- */
/* loadRows accounting                                                         */
/* -------------------------------------------------------------------------- */

function mapping(overrides: Partial<Mapping> = {}): Mapping {
  return {
    schemaVersion: 1,
    name: "t",
    environmentUrl: ENV,
    targetEntitySet: "contacts",
    sourceTable: "T",
    columns: [
      { source: "Email", target: "emailaddress1", kind: "string", treatEmptyAsNull: true },
    ],
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

test("loadRows flags operations the server returned no response for", async () => {
  const fakeClient = {
    batch: async (ops: BatchOperation[]) =>
      // Respond only to the first op — simulates a truncated/partial reply.
      [{ contentId: ops[0].contentId, status: 201, ok: true }],
  } as unknown as DataverseClient;

  const result = await loadRows({
    mapping: mapping(),
    rows: [{ Email: "a@x.com" }, { Email: "b@x.com" }],
    client: fakeClient,
  });

  assert.equal(result.total, 2);
  assert.equal(result.succeeded, 1);
  assert.equal(result.failed, 1);
  assert.match(result.errors[0].message, /no response returned/);
});

test("loadRows rejects non-GUID cell values for guid lookups before any request", async () => {
  let batchCalled = false;
  const fakeClient = {
    batch: async () => {
      batchCalled = true;
      return [];
    },
  } as unknown as DataverseClient;

  const result = await loadRows({
    mapping: mapping({
      columns: [
        {
          source: "Company",
          target: "parentcustomerid_account",
          kind: "lookup",
          bindEntitySet: "accounts",
          lookupResolution: "guid",
          treatEmptyAsNull: true,
        },
      ],
    }),
    rows: [{ Company: "not-a-guid)\r\nDELETE /accounts HTTP/1.1" }],
    client: fakeClient,
  });

  assert.equal(batchCalled, false, "no request should be sent");
  assert.equal(result.failed, 1);
  assert.match(result.errors[0].message, /not a valid GUID/);
});

test("loadRows counts 412 as skipped under skip-if-exists", async () => {
  const fakeClient = {
    batch: async (ops: BatchOperation[]) =>
      ops.map((o, i) => ({
        contentId: o.contentId,
        status: i === 0 ? 201 : 412,
        ok: i === 0,
        errorMessage: i === 0 ? undefined : "exists",
      })),
  } as unknown as DataverseClient;

  const result = await loadRows({
    mapping: mapping({ conflictMode: "skip-if-exists", upsertKey: ["emailaddress1"] }),
    rows: [{ Email: "a@x.com" }, { Email: "b@x.com" }],
    client: fakeClient,
  });

  assert.equal(result.created, 1);
  assert.equal(result.skipped, 1);
  assert.equal(result.failed, 0);
});
