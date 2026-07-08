import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { isWellKnownDevClient, dataverseScope, warnIfWellKnown } from "./auth.js";

describe("isWellKnownDevClient", () => {
  it("returns true for the PowerApps client id", () => {
    assert.equal(isWellKnownDevClient("2ad88395-b77d-4561-9441-d0e40824f9bc"), true);
  });

  it("returns true for the Power Query client id", () => {
    assert.equal(isWellKnownDevClient("51f81489-12ee-4a9e-aaae-a2591f45987d"), true);
  });

  it("returns false for an unknown client id", () => {
    assert.equal(isWellKnownDevClient("00000000-0000-0000-0000-000000000000"), false);
  });

  it("returns false for an empty string", () => {
    assert.equal(isWellKnownDevClient(""), false);
  });
});

describe("dataverseScope", () => {
  it("appends /.default to the origin", () => {
    assert.equal(
      dataverseScope("https://org.crm.dynamics.com"),
      "https://org.crm.dynamics.com/.default"
    );
  });

  it("uses origin only, ignoring any path segment", () => {
    assert.equal(
      dataverseScope("https://org.crm.dynamics.com/some/path"),
      "https://org.crm.dynamics.com/.default"
    );
  });
});

describe("warnIfWellKnown", () => {
  it("writes to stderr for a known client id", () => {
    const captured: string[] = [];
    const origWrite = process.stderr.write.bind(process.stderr);
    (process.stderr as any).write = (
      msg: string | Uint8Array
    ) => {
      captured.push(String(msg));
      return true;
    };
    try {
      warnIfWellKnown("2ad88395-b77d-4561-9441-d0e40824f9bc");
    } finally {
      process.stderr.write = origWrite;
    }
    const output = captured.join("");
    assert.ok(output.includes("dev mode"), "should mention dev mode");
    assert.ok(output.includes("Microsoft PowerApps"), "should name the borrowed client");
  });

  it("is silent for an unknown client id", () => {
    const captured: string[] = [];
    const origWrite = process.stderr.write.bind(process.stderr);
    (process.stderr as any).write = (msg: string | Uint8Array) => {
      captured.push(String(msg));
      return true;
    };
    try {
      warnIfWellKnown("00000000-0000-0000-0000-000000000000");
    } finally {
      process.stderr.write = origWrite;
    }
    assert.equal(captured.length, 0);
  });
});
