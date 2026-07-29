import { describe, it, beforeAll as before, afterAll as after } from "vitest";
import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { parseRunPlan } from "@dvload/core";
import { buildExecutionBatches, validatePlanMappings } from "./run-all.js";

let tmpDir: string;

describe("run-all plan orchestration", () => {
  before(async () => {
    tmpDir = await mkdtemp(path.join(os.tmpdir(), "dvload-run-all-"));
  });

  after(async () => {
    await rm(tmpDir, { recursive: true, force: true });
  });

  it("builds parallel batches from stage/dependency data", () => {
    const plan = parseRunPlan({
      schemaVersion: 1,
      name: "nightly",
      steps: [
        { id: "accounts", mapping: "./a.dvmap.json", workbook: "./x.xlsx", stage: 1 },
        { id: "owners", mapping: "./o.dvmap.json", workbook: "./x.xlsx", stage: 1 },
        { id: "contacts", mapping: "./c.dvmap.json", workbook: "./x.xlsx", dependsOn: ["accounts"] },
      ],
    });
    const batches = buildExecutionBatches(plan);
    assert.equal(batches.length, 2);
    assert.deepEqual(batches[0].map((s) => s.id).sort(), ["accounts", "owners"]);
    assert.deepEqual(
      batches[1].map((s) => s.id),
      ["contacts"]
    );
  });

  it("treats alternateKeyLinks as execution dependencies", () => {
    const plan = parseRunPlan({
      schemaVersion: 1,
      name: "nightly",
      steps: [
        { id: "accounts", mapping: "./a.dvmap.json", workbook: "./x.xlsx", stage: 1 },
        {
          id: "contacts",
          mapping: "./c.dvmap.json",
          workbook: "./x.xlsx",
          alternateKeyLinks: [
            {
              fromStep: "accounts",
              lookupTarget: "parentcustomerid_account",
              keyAttribute: "accountnumber",
            },
          ],
        },
      ],
    });
    const batches = buildExecutionBatches(plan);
    assert.equal(batches.length, 2);
    assert.deepEqual(
      batches[0].map((s) => s.id),
      ["accounts"]
    );
    assert.deepEqual(
      batches[1].map((s) => s.id),
      ["contacts"]
    );
  });

  it("handles mixed stage and non-stage steps", () => {
    const plan = parseRunPlan({
      schemaVersion: 1,
      name: "nightly",
      steps: [
        { id: "accounts", mapping: "./a.dvmap.json", workbook: "./x.xlsx", stage: 1 },
        { id: "contacts", mapping: "./c.dvmap.json", workbook: "./x.xlsx" },
        { id: "leads", mapping: "./l.dvmap.json", workbook: "./x.xlsx", stage: 2 },
      ],
    });
    const batches = buildExecutionBatches(plan);
    assert.equal(batches.length, 2);
    assert.deepEqual(batches[0].map((s) => s.id).sort(), ["accounts", "contacts"]);
    assert.deepEqual(
      batches[1].map((s) => s.id),
      ["leads"]
    );
  });

  it("validates alternate-key cross-step wiring", async () => {
    await writeFile(path.join(tmpDir, "accounts.dvmap.json"), validAccountsMap(), "utf8");
    await writeFile(path.join(tmpDir, "contacts.dvmap.json"), validContactsMap(), "utf8");
    const plan = parseRunPlan({
      schemaVersion: 1,
      name: "nightly",
      steps: [
        { id: "accounts", mapping: "./accounts.dvmap.json", workbook: "./data.xlsx", stage: 1 },
        {
          id: "contacts",
          mapping: "./contacts.dvmap.json",
          workbook: "./data.xlsx",
          stage: 2,
          dependsOn: ["accounts"],
          alternateKeyLinks: [
            {
              fromStep: "accounts",
              lookupTarget: "parentcustomerid_account",
              keyAttribute: "accountnumber",
            },
          ],
        },
      ],
    });
    const result = await validatePlanMappings(plan, tmpDir);
    assert.equal(result.errors.length, 0, result.errors.join("\n"));
  });

  it("fails when alternate-key lookup config is missing", async () => {
    await writeFile(path.join(tmpDir, "accounts.dvmap.json"), validAccountsMap(), "utf8");
    await writeFile(path.join(tmpDir, "contacts.dvmap.json"), invalidContactsMap(), "utf8");
    const plan = parseRunPlan({
      schemaVersion: 1,
      name: "nightly",
      steps: [
        { id: "accounts", mapping: "./accounts.dvmap.json", workbook: "./data.xlsx", stage: 1 },
        {
          id: "contacts",
          mapping: "./contacts.dvmap.json",
          workbook: "./data.xlsx",
          stage: 2,
          dependsOn: ["accounts"],
          alternateKeyLinks: [
            {
              fromStep: "accounts",
              lookupTarget: "parentcustomerid_account",
              keyAttribute: "accountnumber",
            },
          ],
        },
      ],
    });
    const result = await validatePlanMappings(plan, tmpDir);
    assert.ok(
      result.errors.some((e) => e.includes("lookupResolution=alternateKey")),
      result.errors.join("\n")
    );
  });
});

function validAccountsMap(): string {
  return JSON.stringify(
    {
      schemaVersion: 1,
      name: "accounts",
      environmentUrl: "https://org.crm.dynamics.com",
      targetEntitySet: "accounts",
      sourceTable: "tblAccounts",
      conflictMode: "upsert",
      upsertKey: ["accountnumber"],
      columns: [{ source: "AccountNumber", target: "accountnumber", kind: "string" }],
    },
    null,
    2
  );
}

function validContactsMap(): string {
  return JSON.stringify(
    {
      schemaVersion: 1,
      name: "contacts",
      environmentUrl: "https://org.crm.dynamics.com",
      targetEntitySet: "contacts",
      sourceTable: "tblContacts",
      conflictMode: "insert",
      columns: [
        { source: "Email", target: "emailaddress1", kind: "string" },
        {
          source: "ParentAccount",
          target: "parentcustomerid_account",
          kind: "lookup",
          bindEntitySet: "accounts",
          lookupResolution: "alternateKey",
          keyAttribute: "accountnumber",
        },
      ],
    },
    null,
    2
  );
}

function invalidContactsMap(): string {
  return JSON.stringify(
    {
      schemaVersion: 1,
      name: "contacts",
      environmentUrl: "https://org.crm.dynamics.com",
      targetEntitySet: "contacts",
      sourceTable: "tblContacts",
      conflictMode: "insert",
      columns: [
        { source: "Email", target: "emailaddress1", kind: "string" },
        {
          source: "ParentAccount",
          target: "parentcustomerid_account",
          kind: "lookup",
          bindEntitySet: "accounts",
          lookupResolution: "guid",
        },
      ],
    },
    null,
    2
  );
}
