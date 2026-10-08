import { describe, it, beforeAll as before, afterAll as after } from "vitest";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtemp, writeFile, rm, readdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Path to the built CLI entry point.
const CLI = path.resolve(__dirname, "../../dist/index.js");

function dvload(args: string[], opts: { cwd?: string; env?: NodeJS.ProcessEnv } = {}) {
  return spawnSync(process.execPath, [CLI, ...args], {
    encoding: "utf8",
    cwd: opts.cwd,
    env: opts.env ?? process.env,
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

  describe("with no signed-in session", () => {
    let home: string;
    let mappingFile: string;
    let source: string;
    // An empty home: no ~/.dvload, so no credentials of either kind.
    const env = () => ({ ...process.env, HOME: home, USERPROFILE: home, DVLOAD_TELEMETRY: "0" });

    before(async () => {
      home = await mkdtemp(path.join(os.tmpdir(), "dvload-run-home-"));
      mappingFile = path.join(home, "contacts.dvmap.json");
      source = path.join(home, "contacts.csv");
      await writeFile(mappingFile, validMapping());
      await writeFile(source, "Email\na@example.invalid\nb@example.invalid\n");
    });

    after(async () => {
      await rm(home, { recursive: true, force: true });
    });

    it("fails once with the login hint instead of failing every row", async () => {
      const result = dvload(["run", mappingFile, "-w", source, "--non-interactive", "--json"], {
        env: env(),
      });
      assert.equal(result.status, 1, `stderr: ${result.stderr}`);
      assert.match(result.stderr, /dvload login/);
      // No result object: the load never started, so there are no row errors
      // and no failed-rows file holding the whole source.
      assert.equal(result.stdout.trim(), "");
      const logs = await readdir(path.join(home, "logs")).catch(() => [] as string[]);
      assert.deepEqual(logs.filter((f) => f.startsWith("failed_")), []);
    });

    it("a dry run still works, since it never calls Dataverse", async () => {
      const result = dvload(["run", mappingFile, "-w", source, "--dry-run", "--json"], { env: env() });
      assert.equal(result.status, 0, `stderr: ${result.stderr}`);
      assert.equal(JSON.parse(result.stdout).total, 2);
    });
  });
});
