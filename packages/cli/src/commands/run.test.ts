import { describe, it, before, after } from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Path to the built CLI entry point.
const CLI = path.resolve(__dirname, "../../dist/index.js");

function dvload(args: string[], opts: { cwd?: string } = {}) {
  return spawnSync(process.execPath, [CLI, ...args], {
    encoding: "utf8",
    cwd: opts.cwd,
    env: process.env,
  });
}

// Passes both parseMapping and validateMapping.
const validMapping = () =>
  JSON.stringify({
    schemaVersion: 1,
    name: "test run",
    environmentUrl: "https://org1f722cc4.crm.dynamics.com",
    targetEntitySet: "contacts",
    sourceTable: "tblContacts",
    conflictMode: "insert",
    columns: [{ source: "Email", target: "emailaddress1", kind: "string" }],
  });

// Passes parseMapping but fails validateMapping (upsert without key).
const upsertNoKeyMapping = () =>
  JSON.stringify({
    schemaVersion: 1,
    name: "bad upsert",
    environmentUrl: "https://org1f722cc4.crm.dynamics.com",
    targetEntitySet: "contacts",
    sourceTable: "tblContacts",
    conflictMode: "upsert",
    columns: [{ source: "Email", target: "emailaddress1", kind: "string" }],
  });

let tmpDir: string;

describe("dvload run", () => {
  before(async () => {
    tmpDir = await mkdtemp(path.join(os.tmpdir(), "dvload-run-"));
  });

  after(async () => {
    await rm(tmpDir, { recursive: true, force: true });
  });

  it("exits 2 (schema error) when upsert mapping is missing upsertKey", async () => {
    const mappingFile = path.join(tmpDir, "bad-upsert.dvmap.json");
    await writeFile(mappingFile, upsertNoKeyMapping());
    const result = dvload(["run", mappingFile, "-w", path.join(tmpDir, "dummy.xlsx")]);
    assert.equal(result.status, 2, `stderr: ${result.stderr}`);
    assert.ok(result.stderr.includes("upsertKey"), "should mention upsertKey");
  });

  it("exits non-zero when mapping file does not exist", async () => {
    const result = dvload([
      "run",
      path.join(tmpDir, "nonexistent.dvmap.json"),
      "-w",
      path.join(tmpDir, "dummy.xlsx"),
    ]);
    assert.notEqual(result.status, 0);
  });

  it("exits non-zero when workbook file does not exist (after schema check)", async () => {
    const mappingFile = path.join(tmpDir, "valid.dvmap.json");
    await writeFile(mappingFile, validMapping());
    // Workbook is missing — fails after schema validation passes.
    const result = dvload(["run", mappingFile, "-w", path.join(tmpDir, "no-book.xlsx")]);
    assert.notEqual(result.status, 0);
  });
});
