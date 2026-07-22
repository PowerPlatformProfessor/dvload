// Regression test for the aborted-batch bug: Dataverse (without
// continue-on-error, or via older endpoints) can return an OUTER 400 whose
// multipart body still reports earlier changesets as 201 Created. Those
// records exist — treating the whole batch as failed misreports them and
// makes a re-run duplicate them. batch() must parse the body instead of
// throwing, and must send Prefer: odata.continue-on-error.

import { test } from "node:test";
import assert from "node:assert/strict";
import { DataverseClient, type BatchOperation } from "./dataverse.js";

const CRLF = "\r\n";

function multipartBody(): string {
  const part = (contentId: number, statusLine: string, body: string): string =>
    [
      `--batchresponse_outer`,
      `Content-Type: multipart/mixed; boundary=changesetresponse_${contentId}`,
      ``,
      `--changesetresponse_${contentId}`,
      `Content-Type: application/http`,
      `Content-Transfer-Encoding: binary`,
      `Content-ID: ${contentId}`,
      ``,
      statusLine,
      `OData-EntityId: https://unit.crm.dynamics.com/api/data/v9.2/contacts(11111111-1111-1111-1111-11111111111${contentId})`,
      `Content-Type: application/json`,
      ``,
      body,
      `--changesetresponse_${contentId}--`,
    ].join(CRLF);

  return [
    part(1, "HTTP/1.1 201 Created", `{"contactid":"x"}`),
    part(2, "HTTP/1.1 400 Bad Request", `{"error":{"code":"0x80044330","message":"money out of range"}}`),
    `--batchresponse_outer--`,
    ``,
  ].join(CRLF);
}

function makeClient(onFetch: (url: string, init: RequestInit) => Response): {
  client: DataverseClient;
  headers: () => Record<string, string>;
} {
  let captured: Record<string, string> = {};
  const client = new DataverseClient({
    environmentUrl: "https://unit.crm.dynamics.com",
    getToken: async () => "token",
    fetch: (async (url: string | URL | Request, init?: RequestInit) => {
      captured = Object.fromEntries(
        Object.entries((init?.headers ?? {}) as Record<string, string>)
      );
      return onFetch(String(url), init ?? {});
    }) as typeof fetch,
  });
  return { client, headers: () => captured };
}

const OPS: BatchOperation[] = [
  { contentId: 1, method: "POST", url: "contacts", body: { firstname: "a" } },
  { contentId: 2, method: "POST", url: "contacts", body: { firstname: "b" } },
];

test("outer 400 with a multipart body yields per-operation results, not a throw", async () => {
  const { client } = makeClient(
    () => new Response(multipartBody(), { status: 400, statusText: "Bad Request" })
  );

  const results = await client.batch(OPS);
  assert.equal(results.length, 2);

  const [r1, r2] = results;
  assert.equal(r1.contentId, 1);
  assert.equal(r1.ok, true);
  assert.equal(r1.status, 201);
  assert.equal(r2.contentId, 2);
  assert.equal(r2.ok, false);
  assert.equal(r2.status, 400);
  assert.match(r2.errorMessage ?? "", /money out of range/);
});

test("outer 400 with a NON-multipart body still throws", async () => {
  const { client } = makeClient(
    () =>
      new Response(`{"error":{"code":"0x80072560","message":"user not licensed"}}`, {
        status: 400,
        statusText: "Bad Request",
      })
  );
  await assert.rejects(() => client.batch(OPS), /user not licensed/);
});

test("$batch requests send Prefer: odata.continue-on-error", async () => {
  const { client, headers } = makeClient(
    () => new Response(multipartBody(), { status: 200 })
  );
  await client.batch(OPS);
  assert.equal(headers()["Prefer"], "odata.continue-on-error");
});
