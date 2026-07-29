import { describe, it, beforeAll as before, afterAll as after } from "vitest";
import assert from "node:assert/strict";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { validateCommand } from "./validate.js";

// Passes parseMapping AND validateMapping.
const VALID_MAPPING = JSON.stringify({
  schemaVersion: 1,
  name: "test import",
  environmentUrl: "https://org.crm.dynamics.com",
  targetEntitySet: "contacts",
  sourceTable: "tblContacts",
  conflictMode: "insert",
  columns: [{ source: "Email", target: "emailaddress1", kind: "string" }],
});

// Passes parseMapping but fails validateMapping:
// conflictMode=upsert with no upsertKey.
const UPSERT_WITHOUT_KEY = JSON.stringify({
  schemaVersion: 1,
  name: "upsert without key",
  environmentUrl: "https://org.crm.dynamics.com",
  targetEntitySet: "contacts",
  sourceTable: "tblContacts",
  conflictMode: "upsert",
  columns: [{ source: "Email", target: "emailaddress1", kind: "string" }],
});

// Fails parseMapping: missing environmentUrl.
const MISSING_REQUIRED_FIELD = JSON.stringify({
  schemaVersion: 1,
  name: "broken",
  columns: [{ source: "Email", target: "emailaddress1", kind: "string" }],
});

let tmpDir: string;

describe("validateCommand", () => {
  before(async () => {
    tmpDir = await mkdtemp(path.join(os.tmpdir(), "dvload-validate-"));
    await writeFile(path.join(tmpDir, "valid.dvmap.json"), VALID_MAPPING);
    await writeFile(path.join(tmpDir, "upsert-no-key.dvmap.json"), UPSERT_WITHOUT_KEY);
    await writeFile(path.join(tmpDir, "missing-field.dvmap.json"), MISSING_REQUIRED_FIELD);
  });

  after(async () => {
    await rm(tmpDir, { recursive: true, force: true });
  });

  describe("--no-remote (local schema only)", () => {
    it("valid mapping → no error", async () => {
      const prev = process.exitCode;
      process.exitCode = undefined;
      try {
        await validateCommand(path.join(tmpDir, "valid.dvmap.json"), { remote: false });
        assert.equal(process.exitCode, undefined);
      } finally {
        process.exitCode = prev;
      }
    });

    it("upsert without upsertKey → exitCode 1", async () => {
      const prev = process.exitCode;
      process.exitCode = undefined;
      try {
        await validateCommand(path.join(tmpDir, "upsert-no-key.dvmap.json"), { remote: false });
        assert.equal(process.exitCode, 1);
      } finally {
        process.exitCode = prev;
      }
    });

    it("missing required field → rejects (parseMapping throws)", async () => {
      await assert.rejects(
        () => validateCommand(path.join(tmpDir, "missing-field.dvmap.json"), { remote: false }),
        /environmentUrl/
      );
    });

    it("missing file → rejects", async () => {
      await assert.rejects(() =>
        validateCommand(path.join(tmpDir, "nonexistent.dvmap.json"), { remote: false })
      );
    });
  });
});
