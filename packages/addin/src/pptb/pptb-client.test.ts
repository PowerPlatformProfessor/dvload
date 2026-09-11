import { describe, it } from "vitest";
import assert from "node:assert/strict";
import { formatKeyLiteral } from "@dvload/core";
import { PptbDataverseClient, parseKeyExpression, parseOperationUrl } from "./pptb-client.js";
import type { PptbDataverseApi } from "./pptb-bridge.js";

const GUID_A = "11111111-2222-3333-4444-555555555555";
const GUID_B = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";

/**
 * A fake window.dataverseAPI: every method records its calls, and each test
 * overrides only what it needs. Unstubbed methods fail loudly rather than
 * returning undefined into the adapter.
 */
function fakeBridge(overrides: Partial<PptbDataverseApi> = {}): PptbDataverseApi & {
  calls: Array<{ method: string; args: unknown[] }>;
} {
  const calls: Array<{ method: string; args: unknown[] }> = [];
  const notStubbed =
    (name: string) =>
    async (...args: unknown[]) => {
      calls.push({ method: name, args });
      throw new Error(`bridge method ${name} not stubbed for this test`);
    };
  const record = <F extends (...args: never[]) => unknown>(name: string, fn: F): F =>
    ((...args: never[]) => {
      calls.push({ method: name, args });
      return fn(...args);
    }) as F;

  const base: Record<string, unknown> = {};
  for (const name of [
    "create",
    "update",
    "delete",
    "queryData",
    "getEntityMetadata",
    "getAllEntitiesMetadata",
    "getEntityRelatedMetadata",
    "createEntityDefinition",
    "createAttribute",
  ]) {
    const override = (overrides as Record<string, unknown>)[name];
    base[name] = override ? record(name, override as (...args: never[]) => unknown) : notStubbed(name);
  }
  return { ...(base as unknown as PptbDataverseApi), calls };
}

/** getAllEntitiesMetadata stub with one "contact" entity. */
function contactMetadata(): Pick<PptbDataverseApi, "getAllEntitiesMetadata"> {
  return {
    getAllEntitiesMetadata: async () => ({
      value: [
        {
          LogicalName: "contact",
          EntitySetName: "contacts",
          PrimaryIdAttribute: "contactid",
          PrimaryNameAttribute: "fullname",
          DisplayName: { UserLocalizedLabel: { Label: "Contact" } },
          MetadataId: GUID_B,
        },
      ],
    }),
  };
}

/* ─── key-expression parsing ─────────────────────────────────────────────── */

describe("parseKeyExpression", () => {
  it("recognises a bare GUID", () => {
    assert.deepEqual(parseKeyExpression(GUID_A.toUpperCase()), {
      kind: "guid",
      guid: GUID_A,
    });
  });

  it("parses alternate-key pairs and decodes formatKeyLiteral output", () => {
    // Round-trip what the engine actually produces, including the characters
    // formatKeyLiteral escapes beyond encodeURIComponent.
    const nasty = "O'Brien & Sons (50%), ~novel!";
    const expr = `accountnumber=${formatKeyLiteral(nasty)},priority=3,active=true`;
    assert.deepEqual(parseKeyExpression(expr), {
      kind: "alternate",
      pairs: [
        { attribute: "accountnumber", value: nasty },
        { attribute: "priority", value: 3 },
        { attribute: "active", value: true },
      ],
    });
  });

  it("rejects garbage", () => {
    assert.throws(() => parseKeyExpression("no-equals-here"));
  });
});

describe("parseOperationUrl", () => {
  it("splits entity set and key", () => {
    assert.deepEqual(parseOperationUrl("contacts"), {
      entitySet: "contacts",
      key: { kind: "none" },
    });
    assert.deepEqual(parseOperationUrl(`contacts(${GUID_A})`), {
      entitySet: "contacts",
      key: { kind: "guid", guid: GUID_A },
    });
  });

  it("rejects URLs that aren't entity-set-relative", () => {
    assert.throws(() => parseOperationUrl("https://evil.example/x"));
  });
});

/* ─── batch emulation ────────────────────────────────────────────────────── */

describe("PptbDataverseClient.batch", () => {
  it("maps POST creates to 201 with the new id", async () => {
    const bridge = fakeBridge({
      ...contactMetadata(),
      create: async () => ({ id: GUID_A }),
    });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const results = await client.batch([
      { contentId: 1, method: "POST", url: "contacts", body: { lastname: "A" } },
      { contentId: 2, method: "POST", url: "contacts", body: { lastname: "B" } },
    ]);
    assert.deepEqual(
      results.map((r) => [r.contentId, r.status, r.ok, r.id]),
      [
        [1, 201, true, GUID_A],
        [2, 201, true, GUID_A],
      ]
    );
  });

  it("updates when a PATCH-by-guid target exists (204), creates when it doesn't (201, id back in body)", async () => {
    const existing = new Set([GUID_A]);
    let createdBody: Record<string, unknown> | null = null;
    const bridge = fakeBridge({
      ...contactMetadata(),
      queryData: async (q: string) => ({
        value: [...existing].filter((g) => q.includes(g)).map((g) => ({ contactid: g })),
      }),
      update: async () => undefined,
      create: async (_entity, record) => {
        createdBody = record;
        return { id: GUID_B };
      },
    });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const results = await client.batch([
      { contentId: 1, method: "PATCH", url: `contacts(${GUID_A})`, body: { lastname: "X" } },
      { contentId: 2, method: "PATCH", url: `contacts(${GUID_B})`, body: { lastname: "Y" } },
    ]);
    assert.deepEqual(
      results.map((r) => [r.contentId, r.status, r.ok]),
      [
        [1, 204, true],
        [2, 201, true],
      ]
    );
    // The engine strips the primary id from an id-upsert body; the create
    // fallback must restore it or the record gets a random id.
    assert.equal(createdBody!.contactid, GUID_B);
  });

  it("resolves alternate keys via $filter and keeps key attributes in the create body", async () => {
    const queries: string[] = [];
    const bridge = fakeBridge({
      ...contactMetadata(),
      queryData: async (q: string) => {
        queries.push(q);
        return { value: [] };
      },
      create: async () => ({ id: GUID_A }),
    });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const keyExpr = `emailaddress1=${formatKeyLiteral("a'b@x.io")}`;
    const results = await client.batch([
      {
        contentId: 1,
        method: "PATCH",
        url: `contacts(${keyExpr})`,
        body: { emailaddress1: "a'b@x.io", lastname: "Z" },
      },
    ]);
    assert.deepEqual(
      results.map((r) => [r.status, r.ok, r.id]),
      [[201, true, GUID_A]]
    );
    // The probe re-encodes the decoded key value with formatKeyLiteral.
    assert.ok(queries.some((q) => q.includes(`emailaddress1 eq ${formatKeyLiteral("a'b@x.io")}`)));
  });

  it("synthesises 412 for If-None-Match when the record exists, creates when it doesn't", async () => {
    const bridge = fakeBridge({
      ...contactMetadata(),
      queryData: async (q: string) => ({
        value: q.includes(GUID_A) ? [{ contactid: GUID_A }] : [],
      }),
      create: async () => ({ id: GUID_B }),
    });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const results = await client.batch([
      {
        contentId: 1,
        method: "PATCH",
        url: `contacts(${GUID_A})`,
        body: {},
        headers: { "If-None-Match": "*" },
      },
      {
        contentId: 2,
        method: "PATCH",
        url: `contacts(${GUID_B})`,
        body: {},
        headers: { "If-None-Match": "*" },
      },
    ]);
    assert.deepEqual(
      results.map((r) => [r.contentId, r.status, r.ok]),
      [
        [1, 412, false],
        [2, 201, true],
      ]
    );
  });

  it("drops lookup-clear nulls from an emulated-upsert create, and only those", async () => {
    let createdBody: Record<string, unknown> | null = null;
    const bridge = fakeBridge({
      ...contactMetadata(),
      queryData: async () => ({ value: [] }),
      create: async (_entity, record) => {
        createdBody = record;
        return { id: GUID_A };
      },
    });
    const client = new PptbDataverseClient({ dataverse: bridge });
    await client.batch([
      {
        contentId: 1,
        method: "PATCH",
        url: `contacts(${GUID_B})`,
        body: {
          lastname: "K",
          parentcustomerid_account: null, // lookup clear — meaningless on create
          "ownerid@odata.bind": "/systemusers(" + GUID_A + ")", // bind survives
        },
      },
    ]);
    assert.equal(createdBody!.lastname, "K");
    assert.ok(!("parentcustomerid_account" in createdBody!));
    assert.equal(createdBody!["ownerid@odata.bind"], `/systemusers(${GUID_A})`);
  });

  it("maps DELETE to 204 and bridge failures to error items without aborting the batch", async () => {
    const bridge = fakeBridge({
      ...contactMetadata(),
      delete: async (_entity, id: string) => {
        if (id === GUID_B) throw new Error("404 Not Found: contact does not exist");
      },
    });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const results = await client.batch([
      { contentId: 1, method: "DELETE", url: `contacts(${GUID_A})` },
      { contentId: 2, method: "DELETE", url: `contacts(${GUID_B})` },
    ]);
    assert.deepEqual(
      results.map((r) => [r.contentId, r.status, r.ok]),
      [
        [1, 204, true],
        [2, 404, false],
      ]
    );
    assert.match(results[1].errorMessage ?? "", /does not exist/);
  });

  it("refuses per-operation headers the bridge cannot send", async () => {
    const bridge = fakeBridge({ ...contactMetadata() });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const results = await client.batch([
      {
        contentId: 1,
        method: "POST",
        url: "contacts",
        body: {},
        headers: { "MSCRM.BypassCustomPluginExecution": "true" },
      },
    ]);
    assert.equal(results[0].ok, false);
    assert.match(results[0].errorMessage ?? "", /isn't supported in the Power Platform ToolBox/);
    // Nothing must have reached the bridge for that operation.
    assert.ok(!bridge.calls.some((c) => c.method === "create"));
  });

  it("returns one result per operation, in contentId order", async () => {
    const bridge = fakeBridge({
      ...contactMetadata(),
      create: async () => ({ id: GUID_A }),
    });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const ops = Array.from({ length: 9 }, (_, i) => ({
      contentId: i + 1,
      method: "POST" as const,
      url: "contacts",
      body: {},
    }));
    const results = await client.batch(ops);
    assert.deepEqual(
      results.map((r) => r.contentId),
      ops.map((o) => o.contentId)
    );
  });
});

/* ─── queries and metadata ───────────────────────────────────────────────── */

describe("PptbDataverseClient queries", () => {
  it("queryAll follows @odata.nextLink relativised for the bridge", async () => {
    const queries: string[] = [];
    const bridge = fakeBridge({
      queryData: async (q: string) => {
        queries.push(q);
        if (!q.includes("skiptoken")) {
          return {
            value: [{ n: 1 }],
            "@odata.nextLink": "https://org.crm.dynamics.com/api/data/v9.2/contacts?$select=n&$skiptoken=abc",
          };
        }
        return { value: [{ n: 2 }] };
      },
    });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const rows = await client.queryAll("contacts?$select=n");
    assert.deepEqual(rows, [{ n: 1 }, { n: 2 }]);
    assert.deepEqual(queries, ["contacts?$select=n", "contacts?$select=n&$skiptoken=abc"]);
  });

  it("getEntitySetInfo answers from one cached metadata listing", async () => {
    const bridge = fakeBridge({ ...contactMetadata() });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const info = await client.getEntitySetInfo("contacts");
    assert.deepEqual(info, {
      logicalName: "contact",
      primaryIdAttribute: "contactid",
      primaryNameAttribute: "fullname",
    });
    await client.listEntities();
    await client.getEntitySetInfo("contacts");
    assert.equal(bridge.calls.filter((c) => c.method === "getAllEntitiesMetadata").length, 1);
    await assert.rejects(() => client.getEntitySetInfo("nosuchset"), /No entity found/);
  });

  it("resolveByKey returns the single match, null on none, and throws on many", async () => {
    let rows: Array<Record<string, unknown>> = [];
    const bridge = fakeBridge({
      ...contactMetadata(),
      queryData: async () => ({ value: rows }),
    });
    const client = new PptbDataverseClient({ dataverse: bridge });

    rows = [{ contactid: GUID_A }];
    assert.equal(await client.resolveByKey("contacts", "emailaddress1", "a@x.io"), GUID_A);
    rows = [];
    assert.equal(await client.resolveByKey("contacts", "emailaddress1", "a@x.io"), null);
    rows = [{ contactid: GUID_A }, { contactid: GUID_B }];
    await assert.rejects(
      () => client.resolveByKey("contacts", "emailaddress1", "a@x.io"),
      /not a unique key/
    );
  });

  it("getRecord builds a $filter from the key expression and treats 0 or 2+ matches as null", async () => {
    let rows: Array<Record<string, unknown>> = [{ lastname: "K" }];
    const queries: string[] = [];
    const bridge = fakeBridge({
      ...contactMetadata(),
      queryData: async (q: string) => {
        queries.push(q);
        return { value: rows };
      },
    });
    const client = new PptbDataverseClient({ dataverse: bridge });

    assert.deepEqual(await client.getRecord("contacts", GUID_A, ["lastname"]), { lastname: "K" });
    assert.ok(queries[0].includes(`$filter=contactid eq ${GUID_A}`));

    rows = [];
    assert.equal(await client.getRecord("contacts", GUID_A, ["lastname"]), null);
    rows = [{ lastname: "K" }, { lastname: "L" }];
    assert.equal(await client.getRecord("contacts", GUID_A, ["lastname"]), null);
  });

  it("getEntityDefinition merges entity metadata with its attributes", async () => {
    const bridge = fakeBridge({
      getEntityMetadata: async () => ({ LogicalName: "contact", EntitySetName: "contacts" }),
      getEntityRelatedMetadata: async () => ({ value: [{ LogicalName: "lastname" }] }),
    });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const def = await client.getEntityDefinition("contact");
    assert.equal(def.EntitySetName, "contacts");
    assert.deepEqual(def.Attributes, [{ LogicalName: "lastname" }]);
  });

  it("getOptionSetLabels reads the local option set, then falls back to the global one", async () => {
    const asked: string[] = [];
    const bridge = fakeBridge({
      getEntityRelatedMetadata: async (_entity, path: string) => {
        asked.push(path);
        if (path.endsWith("/OptionSet")) throw new Error("404 no local option set");
        return {
          Options: [
            { Value: 1, Label: { UserLocalizedLabel: { Label: "Hot" } } },
            { Value: 2, Label: { UserLocalizedLabel: { Label: "Cold" } } },
          ],
        };
      },
    });
    const client = new PptbDataverseClient({ dataverse: bridge });
    const map = await client.getOptionSetLabels("contact", "temperature");
    assert.deepEqual(map, { Hot: 1, Cold: 2 });
    assert.deepEqual(asked, [
      "Attributes(LogicalName='temperature')/OptionSet",
      "Attributes(LogicalName='temperature')/GlobalOptionSet",
    ]);
  });

  it("createEntityKey is refused with an actionable message", async () => {
    const client = new PptbDataverseClient({ dataverse: fakeBridge() });
    await assert.rejects(() => client.createEntityKey(), /alternate keys/i);
  });
});
