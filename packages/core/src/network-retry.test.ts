// Retrying requests that THROW rather than returning a status.
//
// A laptop sleeping mid-run, a VPN reconnecting, or a flaky link makes fetch
// reject with `TypeError: fetch failed` instead of producing a 503. That used
// to escape the retry loop entirely and fail a whole batch of rows.

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  DataverseClient,
  classifyNetworkError,
  describeNetworkError,
  type BatchOperation,
  type RetryInfo,
} from "./dataverse.js";

/** What undici throws when a connection dies: a wrapper around a coded cause. */
function fetchFailed(code: string): Error {
  const err = new TypeError("fetch failed");
  (err as { cause?: unknown }).cause = Object.assign(new Error(code), { code });
  return err;
}

const BATCH_BODY =
  "--batchresponse_1\r\n" +
  "Content-Type: application/http\r\n" +
  "Content-Transfer-Encoding: binary\r\n" +
  "Content-ID: 1\r\n\r\n" +
  "HTTP/1.1 204 No Content\r\n\r\n" +
  "--batchresponse_1--\r\n";

function okBatchResponse(): Response {
  return new Response(BATCH_BODY, {
    status: 200,
    headers: { "Content-Type": "multipart/mixed; boundary=batchresponse_1" },
  });
}

const OP: BatchOperation = { contentId: 1, method: "PATCH", url: "contacts(x)", body: {} };

function makeClient(
  fetchFn: (...args: unknown[]) => Promise<Response>,
  retries: RetryInfo[] = []
): DataverseClient {
  return new DataverseClient({
    environmentUrl: "https://contoso.crm.dynamics.com",
    getToken: async () => "token",
    fetch: fetchFn as unknown as typeof fetch,
    retry: { baseDelayMs: 1, maxDelayMs: 2, onRetry: (i) => retries.push(i) },
  });
}

test("classifyNetworkError separates never-sent from possibly-executed", () => {
  assert.equal(classifyNetworkError(fetchFailed("ENOTFOUND")), "not-executed");
  assert.equal(classifyNetworkError(fetchFailed("ECONNREFUSED")), "not-executed");
  assert.equal(classifyNetworkError(fetchFailed("UND_ERR_CONNECT_TIMEOUT")), "not-executed");
  // Connection died mid-flight — the server may already have committed.
  assert.equal(classifyNetworkError(fetchFailed("ECONNRESET")), "ambiguous");
  assert.equal(classifyNetworkError(fetchFailed("ETIMEDOUT")), "ambiguous");
  assert.equal(classifyNetworkError(new TypeError("fetch failed")), "ambiguous");
  // Deliberate cancellation must never be replayed.
  const abort = new Error("This operation was aborted");
  abort.name = "AbortError";
  assert.equal(classifyNetworkError(abort), "abort");
});

test("describeNetworkError surfaces the underlying code, not just 'fetch failed'", () => {
  assert.equal(describeNetworkError(fetchFailed("ECONNRESET")), "fetch failed (ECONNRESET)");
});

test("a transient drop is retried and the batch succeeds", async () => {
  const retries: RetryInfo[] = [];
  let calls = 0;
  const client = makeClient(async () => {
    calls++;
    if (calls < 3) throw fetchFailed("ECONNRESET");
    return okBatchResponse();
  }, retries);

  const results = await client.batch([OP], { idempotent: true });
  assert.equal(calls, 3);
  assert.equal(results[0].ok, true);
  assert.equal(retries.length, 2);
  // status 0 marks "never produced a response" so logs can say what happened.
  assert.equal(retries[0].status, 0);
  assert.match(retries[0].error ?? "", /ECONNRESET/);
});

test("a non-idempotent batch is NOT replayed after an ambiguous drop", async () => {
  // insert mode = plain POSTs; the changeset may already have committed, so a
  // replay would duplicate records. Better to fail the batch.
  let calls = 0;
  const client = makeClient(async () => {
    calls++;
    throw fetchFailed("ECONNRESET");
  });

  await assert.rejects(() => client.batch([OP], { idempotent: false }), /fetch failed/);
  assert.equal(calls, 1);
});

test("a non-idempotent batch IS replayed when the request never reached the server", async () => {
  let calls = 0;
  const client = makeClient(async () => {
    calls++;
    if (calls < 2) throw fetchFailed("ENOTFOUND");
    return okBatchResponse();
  });

  const results = await client.batch([OP], { idempotent: false });
  assert.equal(calls, 2);
  assert.equal(results[0].ok, true);
});

test("cancellation is never retried", async () => {
  let calls = 0;
  const client = makeClient(async () => {
    calls++;
    const e = new Error("This operation was aborted");
    e.name = "AbortError";
    throw e;
  });

  await assert.rejects(() => client.batch([OP], { idempotent: true }), /aborted/);
  assert.equal(calls, 1);
});

test("attempts are bounded and the last error propagates", async () => {
  const retries: RetryInfo[] = [];
  let calls = 0;
  const client = new DataverseClient({
    environmentUrl: "https://contoso.crm.dynamics.com",
    getToken: async () => "token",
    fetch: (async () => {
      calls++;
      throw fetchFailed("ECONNRESET");
    }) as unknown as typeof fetch,
    retry: { maxAttempts: 3, baseDelayMs: 1, maxDelayMs: 2, onRetry: (i) => retries.push(i) },
  });

  await assert.rejects(() => client.batch([OP], { idempotent: true }), /fetch failed/);
  assert.equal(calls, 3);
  assert.equal(retries.length, 2); // slept between attempts, not after the last
});

test("a dropped request is recorded in the request log", async () => {
  const entries: Array<{ status: number; ok: boolean; errorBody?: unknown }> = [];
  let calls = 0;
  const client = new DataverseClient({
    environmentUrl: "https://contoso.crm.dynamics.com",
    getToken: async () => "token",
    fetch: (async () => {
      calls++;
      if (calls < 2) throw fetchFailed("ECONNRESET");
      return okBatchResponse();
    }) as unknown as typeof fetch,
    retry: { baseDelayMs: 1, maxDelayMs: 2 },
    onRequest: (e) => entries.push(e),
  });

  await client.batch([OP], { idempotent: true });
  assert.equal(entries.length, 2);
  assert.equal(entries[0].status, 0);
  assert.equal(entries[0].ok, false);
  assert.match(String(entries[0].errorBody), /ECONNRESET/);
  assert.equal(entries[1].status, 200);
});

test("status-code retries still work and report a real status", async () => {
  const retries: RetryInfo[] = [];
  let calls = 0;
  const client = makeClient(async () => {
    calls++;
    if (calls < 2) return new Response("busy", { status: 429 });
    return okBatchResponse();
  }, retries);

  const results = await client.batch([OP], { idempotent: true });
  assert.equal(results[0].ok, true);
  assert.equal(retries[0].status, 429);
  assert.equal(retries[0].error, undefined);
});
