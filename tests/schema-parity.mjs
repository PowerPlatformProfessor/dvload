// Differential check: docs/schema/*.schema.json vs the hand-rolled validators
// in @dvload/core. The schemas are documentation for editors and CI, not the
// enforcement point (see docs/ARCHITECTURE.md), so they WILL drift unless
// something compares them.
//
// The contract this asserts: the schema must never reject a document the
// parser accepts. That direction matters — a schema stricter than the parser
// would red-underline valid mapping files in the user's editor. The reverse
// (schema accepts, parser rejects) is expected for the referential rules
// listed in KNOWN_SCHEMA_GAPS below, which draft-07 cannot express.
//
// Usage:
//   npm run build                       # the core dist/ must exist
//   npm i --no-save ajv@8               # not a repo dependency
//   node tests/schema-parity.mjs
//
// Exits non-zero on any unexpected divergence.

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/**
 * Cases where the parser is stricter than the schema, by necessity: JSON
 * Schema draft-07 cannot express uniqueness of a nested property, references
 * between array items, or cycle detection. Each entry must stay covered by a
 * unit test in packages/core instead.
 */
const KNOWN_SCHEMA_GAPS = new Set([
  "dvmap: upsertKey attribute not present in columns",
  "dvmap: duplicate column targets",
  "dvplan: duplicate step ids",
  "dvplan: dependsOn names an unknown step",
  "dvplan: step depends on itself",
  "dvplan: dependency cycle",
  "dvplan: alternateKeyLinks names an unknown step",
  "dvplan: alternateKeyLinks references its own step",
]);

let Ajv;
try {
  ({ default: Ajv } = await import("ajv"));
} catch {
  console.error("ajv is not installed. Run: npm i --no-save ajv@8");
  process.exit(2);
}

let core;
try {
  core = await import(path.join(ROOT, "packages/core/dist/index.js"));
} catch {
  console.error("packages/core/dist not found. Run: npm run build");
  process.exit(2);
}

const ajv = new Ajv({ allErrors: true, strict: false });
const load = (p) => JSON.parse(readFileSync(path.join(ROOT, p), "utf8"));
const validateMapSchema = ajv.compile(load("docs/schema/dvmap.schema.json"));
const validatePlanSchema = ajv.compile(load("docs/schema/dvplan.schema.json"));

const mappingParserAccepts = (doc) => {
  try {
    return core.validateMapping(core.parseMapping(structuredClone(doc))).length === 0;
  } catch {
    return false;
  }
};

const planParserAccepts = (doc) => {
  try {
    core.parseRunPlan(structuredClone(doc), "parity");
    return true;
  } catch {
    return false;
  }
};

/* ---------------------------------------------------------------------- */
/* Cases                                                                   */
/* ---------------------------------------------------------------------- */

const M = {
  schemaVersion: 1,
  name: "parity",
  environmentUrl: "https://contoso.crm.dynamics.com",
  targetEntitySet: "contacts",
  sourceTable: "tblContacts",
  columns: [{ source: "Surname", target: "lastname", kind: "string", treatEmptyAsNull: true }],
};
const withCol = (o) => ({ ...M, columns: [{ ...M.columns[0], ...o }] });
const GUID = "11111111-2222-3333-4444-555555555555";

const mappingCases = [
  ["baseline", M],
  ["wrong schemaVersion", { ...M, schemaVersion: 2 }],
  ["missing name", { ...M, name: undefined }],
  ["empty columns", { ...M, columns: [] }],
  ["batchSize 0", { ...M, batchSize: 0 }],
  ["batchSize 1000", { ...M, batchSize: 1000 }],
  ["batchSize 1001", { ...M, batchSize: 1001 }],
  ["concurrency 8", { ...M, concurrency: 8 }],
  ["concurrency 9", { ...M, concurrency: 9 }],
  ["negative maxErrors", { ...M, maxErrors: -1 }],
  ["unknown conflictMode", { ...M, conflictMode: "merge" }],
  ["upsert without upsertKey", { ...M, conflictMode: "upsert" }],
  ["upsert with empty upsertKey", { ...M, conflictMode: "upsert", upsertKey: [] }],
  ["upsert valid", { ...M, conflictMode: "upsert", upsertKey: ["lastname"] }],
  ["sync without upsertKey", { ...M, conflictMode: "sync" }],
  ["syncAction without sync", { ...M, syncAction: "delete" }],
  ["skipUnchanged with insert", { ...M, skipUnchanged: true }],
  ["skipUnchanged with upsert", { ...M, conflictMode: "upsert", upsertKey: ["lastname"], skipUnchanged: true }],
  ["malformed impersonateUserId", { ...M, impersonateUserId: "nope" }],
  ["valid impersonateUserId", { ...M, impersonateUserId: GUID }],
  ["plain http notifyUrl", { ...M, notifyUrl: "http://example.invalid/hook" }],
  ["https notifyUrl", { ...M, notifyUrl: "https://hooks.example/x" }],
  ["localhost notifyUrl", { ...M, notifyUrl: "http://localhost:3000/x" }],
  ["unknown kind", withCol({ kind: "guidish" })],
  ["source and constant together", withCol({ constant: "z" })],
  ["neither source nor constant", { ...M, columns: [{ target: "lastname", kind: "string", treatEmptyAsNull: true }] }],
  ["constant only", { ...M, columns: [{ constant: "z", target: "lastname", kind: "string", treatEmptyAsNull: true }] }],
  ["constant of wrong type", { ...M, columns: [{ constant: {}, target: "lastname", kind: "string", treatEmptyAsNull: true }] }],
  ["lookup missing bindEntitySet", withCol({ kind: "lookup", lookupResolution: "guid" })],
  ["lookup missing lookupResolution", withCol({ kind: "lookup", bindEntitySet: "accounts" })],
  ["lookup by guid", withCol({ kind: "lookup", bindEntitySet: "accounts", lookupResolution: "guid" })],
  ["lookup altKey missing keyAttribute", withCol({ kind: "lookup", bindEntitySet: "accounts", lookupResolution: "alternateKey" })],
  ["lookup by altKey", withCol({ kind: "lookup", bindEntitySet: "accounts", lookupResolution: "alternateKey", keyAttribute: "accountnumber" })],
  ["lookup by text, full", withCol({ kind: "lookup", bindEntitySet: "accounts", lookupResolution: "text", keyAttribute: "name", createIfMissing: true, duplicateBehavior: "first" })],
  ["createIfMissing on non-lookup", withCol({ createIfMissing: true })],
  ["duplicateBehavior on non-lookup", withCol({ duplicateBehavior: "first" })],
  ["createIfMissing with guid resolution", withCol({ kind: "lookup", bindEntitySet: "accounts", lookupResolution: "guid", createIfMissing: true })],
  ["unknown duplicateBehavior", withCol({ kind: "lookup", bindEntitySet: "accounts", lookupResolution: "text", keyAttribute: "name", duplicateBehavior: "last" })],
  ["optionMap with string value", withCol({ kind: "choice", target: "tier", optionMap: { Gold: "3" } })],
  ["optionMap valid", withCol({ kind: "choice", target: "tier", optionMap: { Gold: 3 } })],
  ["upsertKey attribute not present in columns", { ...M, conflictMode: "upsert", upsertKey: ["nosuchattr"] }],
  ["duplicate column targets", { ...M, columns: [M.columns[0], { ...M.columns[0], source: "Other" }] }],
];

const P = {
  schemaVersion: 1,
  name: "parity plan",
  steps: [{ id: "a", mapping: "./a.dvmap.json", workbook: "./data.xlsx" }],
};
const withStep = (o) => ({ ...P, steps: [P.steps[0], { id: "b", mapping: "./b.dvmap.json", workbook: "./data.xlsx", ...o }] });

const planCases = [
  ["baseline", P],
  ["wrong schemaVersion", { ...P, schemaVersion: 2 }],
  ["no steps", { ...P, steps: [] }],
  ["missing name", { ...P, name: undefined }],
  ["whitespace-only name", { ...P, name: "   " }],
  ["step missing workbook", { ...P, steps: [{ id: "a", mapping: "./a.dvmap.json" }] }],
  ["stage 0", withStep({ stage: 0 })],
  ["stage 2", withStep({ stage: 2 })],
  ["dependsOn valid", withStep({ dependsOn: ["a"] })],
  ["overrides valid", withStep({ overrides: { maxErrors: 5, concurrency: 4, dryRun: true, user: false, failedRows: false, notifyUrl: "https://hooks.example/x" } })],
  ["overrides concurrency 9", withStep({ overrides: { concurrency: 9 } })],
  ["overrides plain http notifyUrl", withStep({ overrides: { notifyUrl: "http://example.invalid/x" } })],
  ["alternateKeyLinks valid", withStep({ alternateKeyLinks: [{ fromStep: "a", lookupTarget: "parentcustomerid_account", keyAttribute: "accountnumber" }] })],
  ["alternateKeyLinks missing field", withStep({ alternateKeyLinks: [{ fromStep: "a", lookupTarget: "x" }] })],
  ["legacy runs manifest", { runs: [{ mapping: "./a.dvmap.json", workbook: "./data.xlsx", refresh: true }] }],
  ["legacy runs empty", { runs: [] }],
  ["legacy runs with stopOnError", { stopOnError: false, runs: [{ mapping: "./a.dvmap.json", workbook: "./data.xlsx" }] }],
  ["duplicate step ids", { ...P, steps: [P.steps[0], P.steps[0]] }],
  ["dependsOn names an unknown step", withStep({ dependsOn: ["nosuchstep"] })],
  ["step depends on itself", withStep({ id: "b", dependsOn: ["b"] })],
  ["dependency cycle", { ...P, steps: [
    { id: "a", mapping: "./a.dvmap.json", workbook: "./d.xlsx", dependsOn: ["b"] },
    { id: "b", mapping: "./b.dvmap.json", workbook: "./d.xlsx", dependsOn: ["a"] },
  ] }],
  ["alternateKeyLinks names an unknown step", withStep({ alternateKeyLinks: [{ fromStep: "nosuchstep", lookupTarget: "x", keyAttribute: "k" }] })],
  ["alternateKeyLinks references its own step", withStep({ id: "b", alternateKeyLinks: [{ fromStep: "b", lookupTarget: "x", keyAttribute: "k" }] })],
];

/* ---------------------------------------------------------------------- */
/* Run                                                                     */
/* ---------------------------------------------------------------------- */

let failures = 0;
let expectedGaps = 0;
let checked = 0;

function run(kind, cases, parserAccepts, schemaValidate) {
  for (const [label, doc] of cases) {
    checked++;
    const byParser = parserAccepts(doc);
    const bySchema = schemaValidate(structuredClone(doc));
    if (byParser === bySchema) continue;

    const id = `${kind}: ${label}`;
    if (!byParser && bySchema && KNOWN_SCHEMA_GAPS.has(id)) {
      expectedGaps++;
      continue;
    }
    failures++;
    if (byParser && !bySchema) {
      console.error(`FAIL  ${id}\n      parser accepts but the SCHEMA REJECTS it — the schema is too strict.`);
      for (const e of schemaValidate.errors ?? []) {
        console.error(`      ${e.instancePath || "/"} ${e.message}`);
      }
    } else {
      console.error(
        `FAIL  ${id}\n      schema accepts but the parser rejects it. Either tighten the schema, ` +
          `or add "${id}" to KNOWN_SCHEMA_GAPS with a note on why draft-07 can't express it.`
      );
    }
  }
}

run("dvmap", mappingCases, mappingParserAccepts, validateMapSchema);
run("dvplan", planCases, planParserAccepts, validatePlanSchema);

// Every committed fixture must validate.
for (const rel of [
  "packages/core/examples/contacts.dvmap.json",
  "tests/dummy-data/contacts.dvmap.json",
  "tests/dummy-data/contacts-load-insert.dvmap.json",
  "tests/dummy-data/contacts-load-upsert.dvmap.json",
]) {
  checked++;
  const doc = load(rel);
  if (!validateMapSchema(doc)) {
    failures++;
    console.error(`FAIL  fixture ${rel} does not satisfy dvmap.schema.json`);
    for (const e of validateMapSchema.errors ?? []) {
      console.error(`      ${e.instancePath || "/"} ${e.message}`);
    }
  }
}

// Unused entries mean the parser was relaxed (or a case was renamed) and the
// exemption list is now lying about what the schema can't do.
const staleGaps = [...KNOWN_SCHEMA_GAPS].filter((id) => {
  const [kind, label] = [id.slice(0, id.indexOf(":")), id.slice(id.indexOf(": ") + 2)];
  const cases = kind === "dvmap" ? mappingCases : planCases;
  return !cases.some(([l]) => l === label);
});
if (staleGaps.length > 0) {
  failures++;
  console.error(`FAIL  KNOWN_SCHEMA_GAPS has entries with no matching case: ${staleGaps.join(", ")}`);
}

console.log(
  `\n${checked} check(s): ${failures} failure(s), ` +
    `${expectedGaps} known draft-07 gap(s) exercised.`
);
process.exit(failures > 0 ? 1 : 0);
