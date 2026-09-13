import { describe, it } from "vitest";
import assert from "node:assert/strict";
import {
  buildExecutionBatches,
  crossValidatePlanMappings,
  parseRunPlan,
  RunPlanParseError,
  validateRunPlan,
  type RunPlan,
  type RunPlanStep,
} from "./run-plan.js";
import type { ColumnMapping, Mapping } from "./mapping.js";

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

describe("crossValidatePlanMappings", () => {
  const plan = (steps: RunPlanStep[]): RunPlan => ({
    schemaVersion: 1,
    name: "p",
    stopOnError: true,
    steps,
  });
  const child: RunPlanStep = {
    id: "lines",
    mapping: "./lines.dvmap.json",
    workbook: "./x.xlsx",
    alternateKeyLinks: [{ fromStep: "orders", lookupTarget: "orderid", keyAttribute: "ordernumber" }],
  };
  const parent: RunPlanStep = { id: "orders", mapping: "./orders.dvmap.json", workbook: "./x.xlsx" };
  const mapping = (over: Partial<Mapping>): Mapping =>
    ({
      name: "m",
      environmentUrl: "https://x.crm.dynamics.com",
      targetEntitySet: "orders",
      sourceTable: "t",
      conflictMode: "insert",
      columns: [],
      ...over,
    }) as Mapping;

  it("accepts a correctly linked parent/child pair", () => {
    const errs = crossValidatePlanMappings(
      plan([parent, child]),
      new Map([
        ["orders", mapping({ upsertKey: ["ordernumber"] })],
        [
          "lines",
          mapping({
            columns: [
              {
                source: "order",
                target: "orderid",
                kind: "lookup",
                lookupResolution: "alternateKey",
                keyAttribute: "ordernumber",
              } as ColumnMapping,
            ],
          }),
        ],
      ])
    );
    assert.deepEqual(errs, []);
  });

  it("flags an upstream step whose upsertKey lacks the link attribute", () => {
    const errs = crossValidatePlanMappings(
      plan([parent, child]),
      new Map([
        ["orders", mapping({})],
        [
          "lines",
          mapping({
            columns: [
              {
                source: "order",
                target: "orderid",
                kind: "lookup",
                lookupResolution: "alternateKey",
                keyAttribute: "ordernumber",
              } as ColumnMapping,
            ],
          }),
        ],
      ])
    );
    assert.ok(errs.some((e) => e.includes("upsertKey")));
  });

  it("skips steps whose mapping the caller could not load", () => {
    const errs = crossValidatePlanMappings(plan([parent, child]), new Map());
    assert.deepEqual(errs, []);
  });
});

describe("buildExecutionBatches", () => {
  it("orders alternate-key-linked steps parent-first without explicit stages", () => {
    const batches = buildExecutionBatches({
      schemaVersion: 1,
      name: "p",
      stopOnError: true,
      steps: [
        {
          id: "lines",
          mapping: "./l.dvmap.json",
          workbook: "./x.xlsx",
          alternateKeyLinks: [{ fromStep: "orders", lookupTarget: "orderid", keyAttribute: "no" }],
        },
        { id: "orders", mapping: "./o.dvmap.json", workbook: "./x.xlsx" },
      ],
    });
    assert.deepEqual(
      batches.map((b) => b.map((s) => s.id)),
      [["orders"], ["lines"]]
    );
  });
});
