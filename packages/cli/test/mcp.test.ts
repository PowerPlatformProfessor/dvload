/**
 * The MCP server (`dvload mcp`), driven through a real MCP client over an
 * in-memory transport.
 *
 * What matters here is the contract an agent relies on: failures come back as
 * readable tool errors rather than protocol faults, and nothing is written to
 * Dataverse unless dryRun=false was asked for AND the server allows it.
 */

import { describe, it, expect, beforeAll, afterAll } from "vitest";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { withFakeHome, type FakeHome } from "./support/fake-home.js";

type McpModule = typeof import("../src/commands/mcp.js");

const mapping = (extra: Record<string, unknown> = {}) =>
  JSON.stringify({
    schemaVersion: 1,
    name: "mcp test",
    environmentUrl: "https://org.crm.dynamics.com",
    targetEntitySet: "contacts",
    sourceTable: "tblContacts",
    conflictMode: "insert",
    columns: [{ source: "Email", target: "emailaddress1", kind: "string" }],
    ...extra,
  });

let ctx: { mod: McpModule } & FakeHome;
let tmpDir: string;
let validMapping: string;
let upsertNoKey: string;
let csv: string;

async function connect(opts: { readOnly?: boolean } = {}): Promise<Client> {
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  await ctx.mod.buildMcpServer(opts).connect(serverSide);
  const client = new Client({ name: "test", version: "0" });
  await client.connect(clientSide);
  return client;
}

async function call(client: Client, name: string, args: Record<string, unknown>) {
  const res = await client.callTool({ name, arguments: args });
  const text = (res.content as Array<{ text: string }>)[0].text;
  return { isError: res.isError === true, text, json: () => JSON.parse(text) };
}

beforeAll(async () => {
  // A dry run still consults the secure store and telemetry config; keep
  // both away from the developer's real ~/.dvload.
  ctx = await withFakeHome<McpModule>(() => import("../src/commands/mcp.js"));
  tmpDir = await mkdtemp(path.join(os.tmpdir(), "dvload-mcp-"));
  validMapping = path.join(tmpDir, "valid.dvmap.json");
  upsertNoKey = path.join(tmpDir, "upsert-no-key.dvmap.json");
  csv = path.join(tmpDir, "contacts.csv");
  await writeFile(validMapping, mapping());
  await writeFile(upsertNoKey, mapping({ conflictMode: "upsert" }));
  await writeFile(csv, "Email,Name\na@example.invalid,A\nb@example.invalid,B\nc@example.invalid,C\n");
});

afterAll(async () => {
  ctx.cleanup();
  await rm(tmpDir, { recursive: true, force: true });
});

describe("dvload mcp", () => {
  it("advertises its tools, with run_mapping marked as not read-only", async () => {
    const client = await connect();
    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name).sort()).toEqual([
      "auth_status",
      "inspect_source",
      "list_dataflows",
      "list_profiles",
      "run_mapping",
      "validate_mapping",
    ]);
    const byName = Object.fromEntries(tools.map((t) => [t.name, t]));
    expect(byName.run_mapping.annotations?.readOnlyHint).toBe(false);
    expect(byName.validate_mapping.annotations?.readOnlyHint).toBe(true);
  });

  it("inspect_source returns columns and a bounded sample", async () => {
    const res = await call(await connect(), "inspect_source", { file: csv, sampleRows: 2 });
    expect(res.isError).toBe(false);
    const out = res.json();
    expect(out.columns).toEqual(["Email", "Name"]);
    expect(out.rowCount).toBe(3);
    expect(out.sample).toHaveLength(2);
  });

  it("validate_mapping reports schema errors as data, not as a failure", async () => {
    const client = await connect();
    const good = (await call(client, "validate_mapping", { mapping: validMapping, remote: false })).json();
    expect(good).toMatchObject({ ok: true, schemaErrors: [], remote: null });

    const bad = await call(client, "validate_mapping", { mapping: upsertNoKey, remote: false });
    expect(bad.isError).toBe(false);
    expect(bad.json().ok).toBe(false);
    expect(bad.json().schemaErrors.join(" ")).toMatch(/upsertKey/);
  });

  it("run_mapping is a dry run unless told otherwise", async () => {
    const res = await call(await connect(), "run_mapping", { mapping: validMapping, workbook: csv });
    expect(res.isError).toBe(false);
    expect(res.json()).toMatchObject({ dryRun: true, total: 3, failed: 0, errorsOmitted: 0 });
  });

  it("run_mapping surfaces schema errors in the tool result", async () => {
    const res = await call(await connect(), "run_mapping", { mapping: upsertNoKey, workbook: csv });
    expect(res.isError).toBe(true);
    expect(res.text).toMatch(/upsertKey/);
  });

  it("--read-only refuses a real run but still allows a dry run", async () => {
    const client = await connect({ readOnly: true });
    const real = await call(client, "run_mapping", { mapping: validMapping, workbook: csv, dryRun: false });
    expect(real.isError).toBe(true);
    expect(real.text).toMatch(/read-only/);

    const dry = await call(client, "run_mapping", { mapping: validMapping, workbook: csv });
    expect(dry.isError).toBe(false);
  });

  it("a real run with no session fails with a login hint instead of prompting", async () => {
    const res = await call(await connect(), "run_mapping", {
      mapping: validMapping,
      workbook: csv,
      dryRun: false,
    });
    expect(res.isError).toBe(true);
    expect(res.text).toMatch(/dvload login/);
  });

  it("a missing file is a tool error, and leaves no exit code behind", async () => {
    const res = await call(await connect(), "run_mapping", {
      mapping: path.join(tmpDir, "nope.dvmap.json"),
      workbook: csv,
    });
    expect(res.isError).toBe(true);
    expect(res.text).toMatch(/ENOENT/);
    expect(process.exitCode).toBeUndefined();
  });

  it("auth_status reports an unconfigured environment without throwing", async () => {
    const res = await call(await connect(), "auth_status", { env: "https://org.crm.dynamics.com" });
    expect(res.isError).toBe(false);
    expect(res.json()).toMatchObject({ mode: "none", tokenOk: false });
  });

  it("list_profiles is empty on a fresh home", async () => {
    const res = await call(await connect(), "list_profiles", {});
    expect(res.json()).toEqual({ profiles: {} });
  });
});
