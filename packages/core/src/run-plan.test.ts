import { describe, it } from "vitest";
import assert from "node:assert/strict";
import { parseRunPlan, RunPlanParseError, validateRunPlan } from "./run-plan.js";

describe("run-plan", () => {
  it("parses the new .dvplan.json shape", () => {
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
    assert.equal(plan.steps.length, 2);
    assert.equal(plan.steps[1].alternateKeyLinks?.[0]?.fromStep, "accounts");
  });

  it("parses legacy runs[] manifests for backward compatibility", () => {
    const plan = parseRunPlan({
      stopOnError: true,
      runs: [
        { mapping: "./a.dvmap.json", workbook: "./book.xlsx" },
        { mapping: "./b.dvmap.json", workbook: "./book.xlsx", refresh: true },
      ],
    });
    assert.equal(plan.schemaVersion, 1);
    assert.equal(plan.steps[0].id, "run-1");
    assert.equal(plan.steps[0].stage, 1);
    assert.equal(plan.steps[1].stage, 2);
    assert.equal(plan.steps[1].refresh, true);
  });

  it("rejects unknown dependencies", () => {
    assert.throws(
      () =>
        parseRunPlan({
          schemaVersion: 1,
          name: "bad",
          steps: [
            {
              id: "contacts",
              mapping: "./contacts.dvmap.json",
              workbook: "./data.xlsx",
              dependsOn: ["accounts"],
            },
          ],
        }),
      RunPlanParseError
    );
  });

  it("reports dependency cycles", () => {
    const errs = validateRunPlan({
      schemaVersion: 1,
      name: "cycle",
      stopOnError: true,
      steps: [
        { id: "a", mapping: "./a.dvmap.json", workbook: "./x.xlsx", dependsOn: ["b"] },
        { id: "b", mapping: "./b.dvmap.json", workbook: "./x.xlsx", dependsOn: ["a"] },
      ],
    });
    assert.ok(errs.some((e) => e.includes("cycle")));
  });
});
