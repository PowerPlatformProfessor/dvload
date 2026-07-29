/**
 * Throttling, dropped connections, and the "did it execute?" question.
 *
 * These are the failures that actually bite a long-running production load,
 * and the ones hardest to reason about from the code: a 503 that arrives
 * *after* the server committed the batch looks identical to one that arrives
 * before. Getting the replay decision wrong duplicates records.
 */

import { describe, it, expect, beforeEach } from "vitest";
import { createFakeDataverse, contactsFixture, type FakeDataverse } from "../support/fake-dataverse.js";
import { clientFor, contactMapping, contactRows } from "../support/builders.js";
import { loadRows } from "../../src/load.js";
import type { RetryInfo } from "../../src/dataverse.js";

let fake: FakeDataverse;

beforeEach(() => {
  fake = createFakeDataverse({ entities: contactsFixture() });
});

describe("throttling", () => {
  it("retries a 429 and completes the load", async () => {
    fake.throttleNext(2);
    const retries: RetryInfo[] = [];

    const result = await loadRows({
      mapping: contactMapping({ batchSize: 10 }),
      rows: contactRows(10),
      client: clientFor(fake, { retry: { baseDelayMs: 1, maxDelayMs: 5, onRetry: (i) => retries.push(i) } }),
    });

    expect(result.created).toBe(10);
    expect(retries).toHaveLength(2);
    expect(retries.every((r) => r.status === 429)).toBe(true);
  });

  it("honours Retry-After given in seconds", async () => {
    fake.throttleNext(1, { retryAfter: 3 });
    const retries: RetryInfo[] = [];

    await loadRows({
      mapping: contactMapping(),
      rows: contactRows(1),
      // maxDelayMs caps the actual sleep so the test stays fast, but the
      // reported source must still say the delay came from the header.
      client: clientFor(fake, { retry: { baseDelayMs: 1, maxDelayMs: 5, onRetry: (i) => retries.push(i) } }),
    });

    expect(retries[0].source).toBe("retry-after-seconds");
  });

  it("honours Retry-After given as an HTTP date", async () => {
    fake.throttleNext(1, { retryAfter: new Date(Date.now() + 2000).toUTCString() });
    const retries: RetryInfo[] = [];

    await loadRows({
      mapping: contactMapping(),
      rows: contactRows(1),
      client: clientFor(fake, { retry: { baseDelayMs: 1, maxDelayMs: 5, onRetry: (i) => retries.push(i) } }),
    });

    expect(retries[0].source).toBe("retry-after-date");
  });

  it("caps a server-supplied Retry-After at maxDelayMs", async () => {
    // Regression: Retry-After used to bypass maxDelayMs entirely, so a
    // Retry-After of 3600 slept for an hour with no output — indistinguishable
    // from a hung process, and impossible to interrupt cleanly on a schedule.
    fake.throttleNext(1, { retryAfter: 3600 });
    const retries: RetryInfo[] = [];
    const started = Date.now();

    await loadRows({
      mapping: contactMapping(),
      rows: contactRows(1),
      client: clientFor(fake, { retry: { baseDelayMs: 1, maxDelayMs: 50, onRetry: (i) => retries.push(i) } }),
    });

    expect(retries[0].delayMs).toBeLessThanOrEqual(50);
    expect(Date.now() - started).toBeLessThan(2000);
  });

  it("gives up after maxAttempts and fails the rows rather than hanging", async () => {
    fake.throttleNext(99);

    const result = await loadRows({
      mapping: contactMapping({ batchSize: 5, maxErrors: 0 }),
      rows: contactRows(5),
      client: clientFor(fake, { retry: { maxAttempts: 3, baseDelayMs: 1, maxDelayMs: 2 } }),
    });

    expect(result.created).toBe(0);
    expect(result.failed).toBe(5);
  });
});

describe("dropped connections", () => {
  it("does not replay a plain insert batch after an ambiguous drop", async () => {
    // ECONNRESET mid-flight: the server may already have committed. For
    // POST-based inserts a replay would create duplicates, so the client must
    // surface the failure instead.
    fake.failNextNetwork(1, "ECONNRESET");

    const result = await loadRows({
      mapping: contactMapping({ conflictMode: "insert", batchSize: 5, maxErrors: 0 }),
      rows: contactRows(5),
      client: clientFor(fake, { retry: { maxAttempts: 5, baseDelayMs: 1, maxDelayMs: 2 } }),
    });

    expect(result.failed).toBe(5);
    expect(fake.records("contacts")).toHaveLength(0);
  });

  it("does replay an insert batch when the connection provably never opened", async () => {
    // ECONNREFUSED means the request never reached Dataverse, so replaying is
    // safe even for non-idempotent creates.
    fake.failNextNetwork(1, "ECONNREFUSED");

    const result = await loadRows({
      mapping: contactMapping({ conflictMode: "insert", batchSize: 5 }),
      rows: contactRows(5),
      client: clientFor(fake, { retry: { maxAttempts: 5, baseDelayMs: 1, maxDelayMs: 2 } }),
    });

    expect(result.created).toBe(5);
  });

  it("replays an upsert batch after an ambiguous drop, because upserts are idempotent", async () => {
    fake.failNextNetwork(1, "ECONNRESET");

    const result = await loadRows({
      mapping: contactMapping({ conflictMode: "upsert", upsertKey: ["emailaddress1"], batchSize: 5 }),
      rows: contactRows(5),
      client: clientFor(fake, { retry: { maxAttempts: 5, baseDelayMs: 1, maxDelayMs: 2 } }),
    });

    expect(result.created).toBe(5);
    expect(fake.records("contacts")).toHaveLength(5);
  });
});

describe("real sockets", () => {
  it("works against an actual HTTP server, not just the fetch shim", async () => {
    // The shim can't catch bugs that only appear once undici is in the path:
    // header casing, chunked bodies, content-length on multipart replies.
    const server = createFakeDataverse({ entities: contactsFixture() });
    await server.listen();
    try {
      const result = await loadRows({
        mapping: contactMapping({ conflictMode: "upsert", upsertKey: ["emailaddress1"], batchSize: 25 }),
        rows: contactRows(60),
        client: clientFor(server),
      });

      expect(result.created).toBe(60);
      expect(result.failed).toBe(0);
      expect(server.records("contacts")).toHaveLength(60);
    } finally {
      await server.close();
    }
  });
});
