#!/usr/bin/env node
/**
 * Coverage ratchet.
 *
 * A fixed coverage threshold has two failure modes. Set it low and it never
 * catches anything; set it high and every PR that touches a hard-to-test file
 * gets blocked until someone lowers it "temporarily". Both end with a number
 * nobody believes.
 *
 * A ratchet avoids that: the gate is whatever coverage the repo achieved last
 * time, minus a small tolerance for measurement noise. Coverage can go up
 * freely and is recorded automatically. It can only go DOWN by an explicit,
 * reviewable commit to .github/coverage-baseline.json.
 *
 * Usage:
 *   node scripts/coverage-gate.mjs            # check (exits 1 on regression)
 *   node scripts/coverage-gate.mjs --accept   # rewrite the baseline upward
 *
 * Reads coverage/coverage-summary.json, produced by `vitest run --coverage`
 * via the json-summary reporter.
 */

import { readFileSync, writeFileSync, existsSync, mkdirSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const REPO_ROOT = path.resolve(import.meta.dirname, "..");
const SUMMARY = path.join(REPO_ROOT, "coverage", "coverage-summary.json");
const BASELINE = path.join(REPO_ROOT, ".github", "coverage-baseline.json");

/**
 * Tolerance, in percentage points. V8 coverage can wobble by a hair between
 * Node patch releases and across OSes, and a gate that fails on noise trains
 * people to rerun CI until it passes.
 */
const TOLERANCE = 0.5;

const METRICS = ["lines", "statements", "functions", "branches"];

function die(message) {
  console.error(`\n  coverage-gate: ${message}\n`);
  process.exit(1);
}

if (!existsSync(SUMMARY)) {
  die(
    `no coverage summary at ${path.relative(REPO_ROOT, SUMMARY)}.\n` +
      `  Run \`npm run test:coverage\` first.`
  );
}

const summary = JSON.parse(readFileSync(SUMMARY, "utf8"));
const total = summary.total;
if (!total) die("coverage summary has no `total` key — is the json-summary reporter enabled?");

const current = Object.fromEntries(METRICS.map((m) => [m, round(total[m].pct)]));

function round(n) {
  return Math.round(n * 100) / 100;
}

const accept = process.argv.includes("--accept");

if (!existsSync(BASELINE)) {
  mkdirSync(path.dirname(BASELINE), { recursive: true });
  writeFileSync(BASELINE, `${JSON.stringify({ tolerance: TOLERANCE, ...current }, null, 2)}\n`);
  console.log("coverage-gate: no baseline found — created one from this run:");
  console.table(current);
  process.exit(0);
}

const baseline = JSON.parse(readFileSync(BASELINE, "utf8"));
const tolerance = typeof baseline.tolerance === "number" ? baseline.tolerance : TOLERANCE;

const rows = METRICS.map((m) => {
  const was = Number(baseline[m] ?? 0);
  const now = current[m];
  const delta = round(now - was);
  return { metric: m, baseline: was, current: now, delta, regressed: now < was - tolerance };
});

console.table(
  Object.fromEntries(
    rows.map((r) => [
      r.metric,
      { baseline: `${r.baseline}%`, current: `${r.current}%`, delta: `${r.delta >= 0 ? "+" : ""}${r.delta}` },
    ])
  )
);

const regressions = rows.filter((r) => r.regressed);

if (accept) {
  // Record the CURRENT numbers verbatim, including a decrease. That is the
  // entire point of --accept: it is the deliberate, reviewable act of saying
  // "this drop is intended". Clamping to Math.max here would make the flag a
  // no-op on exactly the case it exists for.
  const next = { tolerance, ...Object.fromEntries(rows.map((r) => [r.metric, r.current])) };
  writeFileSync(BASELINE, `${JSON.stringify(next, null, 2)}\n`);
  const drops = rows.filter((r) => r.current < r.baseline);
  console.log(
    drops.length > 0
      ? `coverage-gate: baseline LOWERED (${drops
          .map((r) => `${r.metric} ${r.baseline}% → ${r.current}%`)
          .join(", ")}). Explain why in the PR description.`
      : "coverage-gate: baseline updated."
  );
  process.exit(0);
}

if (regressions.length > 0) {
  die(
    `coverage regressed:\n` +
      regressions.map((r) => `    ${r.metric}: ${r.current}% < ${r.baseline}% (baseline)`).join("\n") +
      `\n\n  Add tests for the code you changed. If the drop is genuinely\n` +
      `  unavoidable, run \`npm run coverage:accept\` and explain why in the PR\n` +
      `  description — the baseline change will show up in review.`
  );
}

// Ratchet up automatically so the next PR is held to the new standard. In CI
// this write is not committed (the workflow diffs it and comments instead);
// locally it keeps your baseline honest.
const improvements = rows.filter((r) => r.current > r.baseline);
if (improvements.length > 0) {
  const next = {
    tolerance,
    ...Object.fromEntries(rows.map((r) => [r.metric, Math.max(r.baseline, r.current)])),
  };
  writeFileSync(BASELINE, `${JSON.stringify(next, null, 2)}\n`);
  console.log(
    `coverage-gate: coverage improved (${improvements
      .map((r) => `${r.metric} +${r.delta}`)
      .join(", ")}) — baseline raised. Commit .github/coverage-baseline.json.`
  );
} else {
  console.log("coverage-gate: OK");
}
