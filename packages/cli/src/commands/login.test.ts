// Hermetic CLI tests: point HOME/USERPROFILE at a temp dir so nothing on
// the machine (secure store, MSAL caches, profiles) leaks into assertions.
import { describe, it, beforeAll as before, afterAll as after } from "vitest";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CLI = path.resolve(__dirname, "../../dist/index.js");
const ENV = "https://example.crm.dynamics.com";

let home: string;
before(() => {
  home = mkdtempSync(path.join(os.tmpdir(), "dvload-test-"));
});
after(() => {
  rmSync(home, { recursive: true, force: true });
});

function dvload(args: string[]) {
  return spawnSync(process.execPath, [CLI, ...args], {
    encoding: "utf8",
    env: { ...process.env, HOME: home, USERPROFILE: home },
  });
}

describe("dvload whoami (fresh environment)", () => {
  it("reports auth mode none and suggests logging in", () => {
    const result = dvload(["whoami", "--env", ENV]);
    assert.equal(result.status, 0, `stderr: ${result.stderr}`);
    assert.ok(result.stdout.includes("none"), "should report mode none");
    assert.ok(result.stdout.includes("Not configured"), "should say not configured");
  });
});

describe("dvload login", () => {
  it("exits non-zero for an invalid environment URL", () => {
    const result = dvload(["login", "--env", "not a url"]);
    assert.notEqual(result.status, 0);
  });

  it("--help exits 0 and documents --env", () => {
    const result = dvload(["login", "--help"]);
    assert.equal(result.status, 0);
    assert.ok(result.stdout.includes("--env"));
  });
});

describe("dvload run flags", () => {
  it("--help documents --json and --non-interactive", () => {
    const result = dvload(["run", "--help"]);
    assert.equal(result.status, 0);
    assert.ok(result.stdout.includes("--json"));
    assert.ok(result.stdout.includes("--non-interactive"));
  });
});

describe("dvload profile", () => {
  it("add + list round-trips", () => {
    const add = dvload(["profile", "add", "test", ENV]);
    assert.equal(add.status, 0, `stderr: ${add.stderr}`);
    const list = dvload(["profile", "list"]);
    assert.equal(list.status, 0);
    assert.ok(list.stdout.includes("test"));
    assert.ok(list.stdout.includes(ENV));
  });
});
