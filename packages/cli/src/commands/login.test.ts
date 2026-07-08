import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CLI = path.resolve(__dirname, "../../dist/index.js");
const ENV = "https://org1f722cc4.crm.dynamics.com";

function dvload(args: string[]) {
  return spawnSync(process.execPath, [CLI, ...args], {
    encoding: "utf8",
    env: process.env,
  });
}

describe("dvload whoami", () => {
  it("exits 0 when delegated token is cached", () => {
    const result = dvload(["whoami", "--env", ENV]);
    assert.equal(result.status, 0, `stderr: ${result.stderr}`);
    assert.ok(result.stdout.includes("delegated"), "should report delegated mode");
  });

  it("exits 0 and reports token OK", () => {
    const result = dvload(["whoami", "--env", ENV]);
    assert.equal(result.status, 0);
    assert.ok(result.stdout.includes("Token acquisition: OK"));
  });
});

describe("dvload login", () => {
  it("exits non-zero for an obviously invalid environment URL", () => {
    // A completely bogus URL causes an error before any interactive prompt.
    const result = dvload(["login", "--env", "https://not-a-real-env.invalid"]);
    assert.notEqual(result.status, 0);
  });
});

describe("dvload logout / login flags", () => {
  it("--help exits 0", () => {
    const result = dvload(["login", "--help"]);
    assert.equal(result.status, 0);
    assert.ok(result.stdout.includes("--env"));
  });
});
