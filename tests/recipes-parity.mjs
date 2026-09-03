// Every JSON example in docs/RECIPES.md must parse and validate.
//
// Documentation that used to work is worse than none: someone copies a
// mapping, hits a validation error, and concludes the tool is broken. The
// recipes are the first thing a new user touches, so they are checked
// against the real parser rather than reviewed by eye.
//
// Fragments are supported deliberately — most recipes show only the lines
// that matter, and spelling out a whole mapping every time would bury the
// point. A fragment is spliced into a minimal valid mapping before checking,
// so what's verified is "these lines are correct in context".
//
// Usage: node tests/recipes-parity.mjs   (run from the repo root)

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, "..");
const recipes = path.join(repoRoot, "docs", "RECIPES.md");

// pathToFileURL, not a bare path: Windows absolute paths ("C:\…") are not
// valid ESM specifiers, so a plain import() fails on the platform most
// users run this on.
const { parseMapping, validateMapping } = await import(
  pathToFileURL(path.join(repoRoot, "packages", "core", "dist", "index.js")).href
).catch(() => {
  console.error(
    "Could not load @dvload/core. Build it first:\n" +
      "  npm run build --workspace=@dvload/core"
  );
  process.exit(2);
});

/**
 * Minimal valid mapping that fragments are merged into. It maps the
 * attributes the upsert and sync recipes name as keys, so the only failures
 * reported are ones a reader would actually hit.
 */
const base = {
  schemaVersion: 1,
  name: "recipe check",
  environmentUrl: "https://example.crm.dynamics.com",
  targetEntitySet: "contacts",
  sourceTable: "t",
  conflictMode: "insert",
  batchSize: 100,
  maxErrors: 0,
  logDir: "./logs",
  concurrency: 1,
  bypassCustomLogic: false,
  skipUnchanged: false,
  columns: [
    { source: "A", target: "lastname", kind: "string", treatEmptyAsNull: true },
    { source: "E", target: "emailaddress1", kind: "string", treatEmptyAsNull: true },
    { source: "N", target: "accountnumber", kind: "string", treatEmptyAsNull: true },
  ],
};

const md = readFileSync(recipes, "utf8");
const blocks = [...md.matchAll(/```json\n([\s\S]*?)```/g)].map((m) => m[1].trim());

if (blocks.length === 0) {
  console.error("No ```json blocks found in docs/RECIPES.md — has it moved?");
  process.exit(2);
}

let failures = 0;

for (const [i, raw] of blocks.entries()) {
  // `"columns": [...]` is a valid object body, not a whole document.
  const text = raw.startsWith("{") ? raw : `{${raw}}`;

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    report(i, raw, `not valid JSON — ${e.message}`);
    continue;
  }

  let candidate;
  if (parsed.schemaVersion) {
    candidate = parsed; // a complete mapping
  } else if (Array.isArray(parsed.columns)) {
    candidate = { ...base, ...parsed, columns: [...base.columns, ...parsed.columns] };
  } else if (parsed.target) {
    candidate = { ...base, columns: [...base.columns, parsed] }; // one column
  } else {
    candidate = { ...base, ...parsed }; // mapping-level settings
  }

  try {
    const errors = validateMapping(parseMapping(candidate));
    if (errors.length) report(i, raw, errors.join("; "));
  } catch (e) {
    report(i, raw, e.message);
  }
}

function report(index, raw, message) {
  failures++;
  const firstLine = raw.split("\n").find((l) => l.trim()) ?? "";
  console.error(`\nBlock ${index + 1} (${firstLine.trim().slice(0, 60)}…)\n  ${message}`);
}

if (failures > 0) {
  console.error(`\n${failures} of ${blocks.length} examples in docs/RECIPES.md are broken.`);
  process.exit(1);
}

console.log(`docs/RECIPES.md: all ${blocks.length} JSON examples parse and validate.`);
