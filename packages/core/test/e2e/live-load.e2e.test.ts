/**
 * Live end-to-end load against a real Dataverse sandbox.
 *
 * This is the automated core of TEST-PROTOCOL.md §§6-7. It does NOT replace
 * the manual protocol — the add-in UI, Excel COM, Windows scheduling and
 * conditional-access flows still need a human — but it removes the part that
 * is both the most tedious to do by hand and the most likely to regress.
 *
 * SAFETY RULES, in order of importance:
 *   1. Every record created here carries lastname = "DVLT-E2E" plus a
 *      per-run marker, so cleanup can never touch anything else.
 *   2. Cleanup runs in afterAll AND is idempotent, so a crashed run leaves at
 *      most one generation of junk behind.
 *   3. The suite refuses to run against an environment URL that doesn't look
 *      like a sandbox unless DVLOAD_E2E_I_KNOW_WHAT_IM_DOING=1 is set.
 *   4. No test uses conflictMode "sync": a sync pass deactivates everything
 *      not in the source, and a mistake there is unrecoverable.
 */

import { describe, it, expect, beforeAll, afterAll } from "vitest";
import { DataverseClient } from "../../src/dataverse.js";
import { parseMapping } from "../../src/mapping.js";
import { loadRows } from "../../src/load.js";
import type { SourceRow } from "../../src/types.js";

const ENV_URL = process.env.DVLOAD_E2E_ENV_URL;
const CLIENT_ID = process.env.DVLOAD_E2E_CLIENT_ID;
const TENANT_ID = process.env.DVLOAD_E2E_TENANT_ID;
const SECRET = process.env.DVLOAD_E2E_SECRET;

const configured = Boolean(ENV_URL && CLIENT_ID && TENANT_ID && SECRET);

/**
 * Unique per run so two overlapping runs (a nightly and a manual dispatch)
 * cannot delete each other's records.
 */
const RUN_MARKER = `E2E-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
const SENTINEL_LASTNAME = "DVLT-E2E";

let client: DataverseClient;

/** Client-credentials token. Kept inline so the suite has no CLI dependency. */
async function getToken(): Promise<string> {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: CLIENT_ID!,
      client_secret: SECRET!,
      scope: `${new URL(ENV_URL!).origin}/.default`,
      grant_type: "client_credentials",
    }),
  });
  if (!res.ok) throw new Error(`token request failed: ${res.status} ${await res.text()}`);
  return ((await res.json()) as { access_token: string }).access_token;
}

function rows(n: number): SourceRow[] {
  return Array.from({ length: n }, (_, i) => ({
    Email: `${RUN_MARKER.toLowerCase()}-${i}@dvload-e2e.invalid`,
    First: `E2E${i}`,
    Last: SENTINEL_LASTNAME,
  }));
}

function mapping(over: Record<string, unknown> = {}) {
  return parseMapping({
    schemaVersion: 1,
    name: "e2e",
    environmentUrl: ENV_URL,
    targetEntitySet: "contacts",
    sourceTable: "tblE2E",
    columns: [
      { source: "Email", target: "emailaddress1", kind: "string", treatEmptyAsNull: true },
      { source: "First", target: "firstname", kind: "string", treatEmptyAsNull: true },
      { source: "Last", target: "lastname", kind: "string", treatEmptyAsNull: true },
    ],
    conflictMode: "insert",
    batchSize: 200,
    maxErrors: 0,
    logDir: "./logs",
    concurrency: 2,
    bypassCustomLogic: false,
    skipUnchanged: false,
    ...over,
  });
}

/** Delete every contact this run created. Safe to call more than once. */
async function cleanup(): Promise<number> {
  const found = await client.queryAll(
    `contacts?$select=contactid&$filter=lastname eq '${SENTINEL_LASTNAME}' and startswith(firstname,'E2E')` +
      ` and contains(emailaddress1,'${RUN_MARKER.toLowerCase()}')`
  );
  let deleted = 0;
  for (let i = 0; i < found.length; i += 200) {
    const slice = found.slice(i, i + 200);
    const results = await client.batch(
      slice.map((r, idx) => ({
        contentId: idx + 1,
        method: "DELETE" as const,
        url: `contacts(${String(r.contactid)})`,
      }))
    );
    deleted += results.filter((r) => r.ok).length;
  }
  return deleted;
}

describe.skipIf(!configured)("live Dataverse load", () => {
  beforeAll(async () => {
    const url = new URL(ENV_URL!);
    const looksLikeSandbox =
      /(^|[.-])(dev|test|sandbox|uat|sit|qa)([.-]|$)/i.test(url.hostname) ||
      process.env.DVLOAD_E2E_I_KNOW_WHAT_IM_DOING === "1";
    if (!looksLikeSandbox) {
      throw new Error(
        `Refusing to run E2E against "${url.hostname}": the hostname does not look like a ` +
          `sandbox. Rename the environment, or set DVLOAD_E2E_I_KNOW_WHAT_IM_DOING=1 if you ` +
          `are certain this is not production.`
      );
    }

    client = new DataverseClient({
      environmentUrl: ENV_URL!,
      getToken,
      retry: { maxAttempts: 6, baseDelayMs: 1000, maxDelayMs: 30_000 },
    });

    // Fail fast and legibly if the app user lacks permissions, rather than
    // halfway through a load with a confusing per-row error.
    const info = await client.getEntitySetInfo("contacts");
    expect(info.primaryIdAttribute).toBe("contactid");
  });

  afterAll(async () => {
    if (!configured || !client) return;
    const deleted = await cleanup();
    console.log(`e2e cleanup: deleted ${deleted} record(s) for run ${RUN_MARKER}`);
  });

  it("inserts a batch of records and reports accurate counts", async () => {
    const data = rows(500);
    const result = await loadRows({ mapping: mapping(), rows: data, client });

    expect(result.failed).toBe(0);
    expect(result.created).toBe(500);
    expect(result.succeeded + result.failed + result.skipped).toBe(result.total);

    const found = await client.queryAll(
      `contacts?$select=contactid&$filter=contains(emailaddress1,'${RUN_MARKER.toLowerCase()}')`
    );
    expect(found).toHaveLength(500);
  });

  it("upserts the same rows without creating duplicates", async () => {
    // Requires an active alternate key on contact.emailaddress1 —
    // TEST-PROTOCOL.md §1.2. Without it Dataverse returns 400 and this test
    // tells you exactly what is missing.
    const data = rows(500);
    const result = await loadRows({
      mapping: mapping({ conflictMode: "upsert", upsertKey: ["emailaddress1"] }),
      rows: data,
      client,
    });

    expect(result.failed).toBe(0);
    expect(result.created).toBe(0);
    expect(result.updated).toBe(500);

    const found = await client.queryAll(
      `contacts?$select=contactid&$filter=contains(emailaddress1,'${RUN_MARKER.toLowerCase()}')`
    );
    expect(found).toHaveLength(500);
  });

  it("skips rows whose values already match when skipUnchanged is on", async () => {
    const data = rows(200);
    const result = await loadRows({
      mapping: mapping({ conflictMode: "upsert", upsertKey: ["emailaddress1"], skipUnchanged: true }),
      rows: data,
      client,
    });

    expect(result.failed).toBe(0);
    expect(result.unchanged).toBe(200);
  });

  it("survives real throttling on a larger load", async () => {
    // 5k rows at concurrency 4 reliably trips Dataverse service protection
    // limits on a sandbox — which is the point: the retry path is exercised
    // against the real Retry-After values, not a fake.
    const retries: number[] = [];
    const throttleClient = new DataverseClient({
      environmentUrl: ENV_URL!,
      getToken,
      retry: {
        maxAttempts: 8,
        baseDelayMs: 1000,
        maxDelayMs: 60_000,
        onRetry: (i) => retries.push(i.status),
      },
    });

    const data = rows(5000);
    const result = await loadRows({
      mapping: mapping({ conflictMode: "upsert", upsertKey: ["emailaddress1"], concurrency: 4 }),
      rows: data,
      client: throttleClient,
    });

    expect(result.failed).toBe(0);
    expect(result.succeeded).toBe(5000);
    if (retries.length > 0) {
      console.log(`observed ${retries.length} retries, statuses: ${[...new Set(retries)].join(", ")}`);
    }
  });
});

describe.skipIf(configured)("live Dataverse load (not configured)", () => {
  it.skip("skipped: DVLOAD_E2E_* environment variables are not set", () => {
    // Present so the reason shows up in the report instead of an empty file.
  });
});
