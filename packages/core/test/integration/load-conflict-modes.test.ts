/**
 * End-to-end conflict-mode behaviour, driven through the REAL DataverseClient
 * over the fake Web API. These are the automated counterpart to TEST-PROTOCOL
 * §6 (insert/dry-run) and §7 (conflict modes) — the sections that previously
 * only existed as a manual checklist.
 *
 * If a change breaks batch serialization, upsert status accounting, or the
 * sync safety guard, it fails here rather than in someone's production load.
 */

import { describe, it, expect, beforeEach } from "vitest";
import { createFakeDataverse, contactsFixture, type FakeDataverse } from "../support/fake-dataverse.js";
import { clientFor, contactMapping, contactRows } from "../support/builders.js";
import { loadRows } from "../../src/load.js";

let fake: FakeDataverse;

beforeEach(() => {
  fake = createFakeDataverse({ entities: contactsFixture() });
});

describe("insert", () => {
  it("creates one record per row and reports them as created", async () => {
    const rows = contactRows(250);
    const result = await loadRows({
      mapping: contactMapping({ batchSize: 100 }),
      rows,
      client: clientFor(fake),
    });

    expect(result.created).toBe(250);
    expect(result.updated).toBe(0);
    expect(result.failed).toBe(0);
    expect(fake.records("contacts")).toHaveLength(250);
  });

  it("sends each row as its own changeset so one failure cannot roll back its neighbours", async () => {
    // A single shared changeset is atomic in Dataverse: the first failure
    // aborts every other operation in the batch. The client deliberately
    // wraps each op in its own changeset — assert on the wire, because this
    // is invisible from the result object.
    fake.setOperationFault((op) => (op.seq === 3 ? { status: 400, message: "plugin rejected row" } : null));

    const result = await loadRows({
      mapping: contactMapping({ batchSize: 10, maxErrors: 0 }),
      rows: contactRows(10),
      client: clientFor(fake),
    });

    expect(result.failed).toBe(1);
    expect(result.created).toBe(9);
    expect(fake.records("contacts")).toHaveLength(9);
  });

  it("asks Dataverse to continue past failed changesets", async () => {
    await loadRows({ mapping: contactMapping(), rows: contactRows(3), client: clientFor(fake) });

    const batch = fake.requests.find((r) => r.path === "$batch");
    expect(batch?.headers["prefer"]).toContain("odata.continue-on-error");
  });

  it("writes nothing on a dry run but still reports what it would have done", async () => {
    const result = await loadRows({
      mapping: contactMapping(),
      rows: contactRows(50),
      client: clientFor(fake),
      dryRun: true,
    });

    expect(result.created).toBe(50);
    expect(fake.records("contacts")).toHaveLength(0);
    expect(fake.requests.filter((r) => r.path === "$batch")).toHaveLength(0);
  });
});

describe("upsert on an alternate key", () => {
  const upsertMapping = () =>
    contactMapping({ conflictMode: "upsert", upsertKey: ["emailaddress1"], batchSize: 50 });

  it("creates on first run and updates on the second", async () => {
    const rows = contactRows(120);

    const first = await loadRows({ mapping: upsertMapping(), rows, client: clientFor(fake) });
    expect(first.created).toBe(120);
    expect(first.updated).toBe(0);

    const second = await loadRows({ mapping: upsertMapping(), rows, client: clientFor(fake) });
    expect(second.created).toBe(0);
    expect(second.updated).toBe(120);
    // The decisive assertion: an upsert must not duplicate.
    expect(fake.records("contacts")).toHaveLength(120);
  });

  it("mixes creates and updates in a single run", async () => {
    await loadRows({ mapping: upsertMapping(), rows: contactRows(50), client: clientFor(fake) });

    const result = await loadRows({
      mapping: upsertMapping(),
      rows: contactRows(80), // 50 existing + 30 new
      client: clientFor(fake),
    });

    expect(result.updated).toBe(50);
    expect(result.created).toBe(30);
    expect(fake.records("contacts")).toHaveLength(80);
  });

  it("round-trips values that would break a naively built request URL", async () => {
    // Key values come from spreadsheet cells. Quotes, slashes, plus signs and
    // non-ASCII all have to survive formatKeyLiteral → URL → the server's
    // decoder without changing identity.
    const nasty = [
      "o'brien@example.test",
      "a+b@example.test",
      "user/name@example.test",
      "ünïcøde@example.test",
      "spaces in key@example.test",
    ];
    const rows = nasty.map((Email) => ({ Email, First: "X", Last: "DVLT-Load" }));

    await loadRows({ mapping: upsertMapping(), rows, client: clientFor(fake) });
    const again = await loadRows({ mapping: upsertMapping(), rows, client: clientFor(fake) });

    expect(again.updated).toBe(nasty.length);
    expect(fake.records("contacts")).toHaveLength(nasty.length);
    expect(
      fake
        .records("contacts")
        .map((r) => r.emailaddress1)
        .sort()
    ).toEqual([...nasty].sort());
  });
});

describe("skip-if-exists", () => {
  it("counts an existing record as skipped, not as an error", async () => {
    const mapping = contactMapping({
      conflictMode: "skip-if-exists",
      upsertKey: ["emailaddress1"],
      batchSize: 25,
    });
    const rows = contactRows(40);

    const first = await loadRows({ mapping, rows, client: clientFor(fake) });
    expect(first.created).toBe(40);

    const second = await loadRows({ mapping, rows, client: clientFor(fake) });
    expect(second.skipped).toBe(40);
    expect(second.failed).toBe(0);
    expect(second.created).toBe(0);
  });

  it("sends If-None-Match so the server, not the client, decides existence", async () => {
    await loadRows({
      mapping: contactMapping({ conflictMode: "skip-if-exists", upsertKey: ["emailaddress1"] }),
      rows: contactRows(2),
      client: clientFor(fake),
    });

    const ops = fake.requests.find((r) => r.path === "$batch")?.batchOps ?? [];
    expect(ops).toHaveLength(2);
    for (const op of ops) expect(op.headers["if-none-match"]).toBe("*");
  });
});

describe("sync", () => {
  const syncMapping = (over = {}) =>
    contactMapping({
      conflictMode: "sync",
      upsertKey: ["emailaddress1"],
      batchSize: 50,
      ...over,
    });

  it("deactivates target records absent from the source", async () => {
    await loadRows({ mapping: syncMapping(), rows: contactRows(30), client: clientFor(fake) });

    // Source shrinks to the first 20 rows.
    const result = await loadRows({
      mapping: syncMapping(),
      rows: contactRows(20),
      client: clientFor(fake),
    });

    expect(result.removed).toBe(10);
    const inactive = fake.records("contacts").filter((r) => r.statecode === 1);
    expect(inactive).toHaveLength(10);
    // Deactivate must not delete.
    expect(fake.records("contacts")).toHaveLength(30);
  });

  it("deletes target records when syncAction is delete", async () => {
    await loadRows({
      mapping: syncMapping({ syncAction: "delete" }),
      rows: contactRows(30),
      client: clientFor(fake),
    });

    const result = await loadRows({
      mapping: syncMapping({ syncAction: "delete" }),
      rows: contactRows(20),
      client: clientFor(fake),
    });

    expect(result.removed).toBe(10);
    expect(fake.records("contacts")).toHaveLength(20);
  });

  it("refuses to run against an empty source instead of wiping the table", async () => {
    await loadRows({ mapping: syncMapping(), rows: contactRows(10), client: clientFor(fake) });

    await expect(loadRows({ mapping: syncMapping(), rows: [], client: clientFor(fake) })).rejects.toThrow(
      /refuses to run with 0 source rows/i
    );

    expect(fake.records("contacts").filter((r) => r.statecode === 1)).toHaveLength(0);
  });

  it("does not run the removal pass when the load was cancelled", async () => {
    // Removing based on a partially loaded source would deactivate records
    // that simply were never reached.
    await loadRows({ mapping: syncMapping(), rows: contactRows(30), client: clientFor(fake) });

    const controller = new AbortController();
    const result = await loadRows({
      mapping: syncMapping({ batchSize: 5 }),
      rows: contactRows(10),
      client: clientFor(fake),
      signal: controller.signal,
      onProgress: (e) => {
        if (e.type === "batch") controller.abort();
      },
    });

    expect(result.cancelled).toBe(true);
    expect(result.removed).toBe(0);
    expect(fake.records("contacts").filter((r) => r.statecode === 1)).toHaveLength(0);
  });

  it("pages through more target records than fit in one response", async () => {
    // Exercises @odata.nextLink following in queryAll; a bug here silently
    // truncates the "what should be removed" set.
    const paged = createFakeDataverse({ entities: contactsFixture(), pageSize: 7 });
    await loadRows({ mapping: syncMapping(), rows: contactRows(40), client: clientFor(paged) });

    const result = await loadRows({
      mapping: syncMapping(),
      rows: contactRows(10),
      client: clientFor(paged),
    });

    expect(result.removed).toBe(30);
  });
});

describe("skipUnchanged", () => {
  it("reports rows whose values already match as unchanged and sends nothing for them", async () => {
    const mapping = contactMapping({
      conflictMode: "upsert",
      upsertKey: ["emailaddress1"],
      skipUnchanged: true,
      batchSize: 20,
    });
    const rows = contactRows(20);

    await loadRows({ mapping, rows, client: clientFor(fake) });
    fake.resetFaults(); // clears the recorded request log

    const second = await loadRows({ mapping, rows, client: clientFor(fake) });

    expect(second.unchanged).toBe(20);
    expect(second.updated).toBe(0);
    expect(fake.requests.filter((r) => r.path === "$batch")).toHaveLength(0);
  });

  it("sends only the attributes that actually differ", async () => {
    const mapping = contactMapping({
      conflictMode: "upsert",
      upsertKey: ["emailaddress1"],
      skipUnchanged: true,
      batchSize: 10,
    });
    await loadRows({ mapping, rows: contactRows(1), client: clientFor(fake) });
    fake.resetFaults();

    const changed = [{ Email: "contact0@example.test", First: "Renamed", Last: "DVLT-Load" }];
    const result = await loadRows({ mapping, rows: changed, client: clientFor(fake) });

    expect(result.updated).toBe(1);
    const batch = fake.requests.find((r) => r.path === "$batch");
    const body = batch?.body ?? "";
    expect(body).toContain("Renamed");
    // lastname is identical to the stored value, so it must not be resent.
    expect(body).not.toContain("DVLT-Load");
  });
});

describe("error budget", () => {
  it("stops scheduling batches once maxErrors is reached and accounts for the remainder", async () => {
    fake.setOperationFault(() => ({ status: 400, message: "always fails" }));

    const result = await loadRows({
      mapping: contactMapping({ batchSize: 5, maxErrors: 5 }),
      rows: contactRows(50),
      client: clientFor(fake),
    });

    expect(result.failed).toBeGreaterThanOrEqual(5);
    expect(result.failed).toBeLessThan(50);
    // Every row is accounted for exactly once — the counters must add up or
    // the CLI's summary lies to the user.
    expect(result.succeeded + result.failed + result.skipped).toBe(result.total);
  });

  it("surfaces the server's error message rather than a generic status", async () => {
    fake.setOperationFault((op) =>
      op.seq === 0
        ? { status: 400, code: "0x80040333", message: "Business Process Error: tier is required" }
        : null
    );

    const result = await loadRows({
      mapping: contactMapping({ batchSize: 5, maxErrors: 0 }),
      rows: contactRows(5),
      client: clientFor(fake),
    });

    expect(result.errors[0].message).toMatch(/tier is required/);
    expect(result.errors[0].httpStatus).toBe(400);
  });
});

describe("resume", () => {
  it("skips rows before the checkpoint offset without re-sending them", async () => {
    const result = await loadRows({
      mapping: contactMapping({ batchSize: 10 }),
      rows: contactRows(50),
      client: clientFor(fake),
      startOffset: 20,
    });

    expect(result.skipped).toBe(20);
    expect(result.created).toBe(30);
    expect(fake.records("contacts")).toHaveLength(30);
  });
});

describe("concurrency", () => {
  it("produces the same totals with a parallel pool as with a sequential one", async () => {
    const rows = contactRows(200);

    const sequential = await loadRows({
      mapping: contactMapping({ batchSize: 20, concurrency: 1 }),
      rows,
      client: clientFor(fake),
    });

    const parallelFake = createFakeDataverse({ entities: contactsFixture() });
    const parallel = await loadRows({
      mapping: contactMapping({ batchSize: 20, concurrency: 4 }),
      rows,
      client: clientFor(parallelFake),
    });

    expect(parallel.created).toBe(sequential.created);
    expect(parallel.failed).toBe(sequential.failed);
    expect(parallelFake.records("contacts")).toHaveLength(200);
  });
});
