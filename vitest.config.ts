/**
 * Root Vitest configuration.
 *
 * The suite is split into named projects so CI can run the cheap, portable
 * parts everywhere and the expensive/platform-bound parts only where they
 * make sense:
 *
 *   core          — pure engine unit tests. Fast, no I/O, runs on every OS.
 *   core-int      — integration tests driving the real DataverseClient against
 *                   the in-process fake Web API (test/support/fake-dataverse).
 *   core-property — fast-check property tests over coercion/mapping. Slower,
 *                   deliberately separated so a flaky counterexample is
 *                   obvious in the report.
 *   cli           — CLI unit tests. Some touch DPAPI/schtasks and self-skip
 *                   off Windows (see test/setup/platform.ts).
 *   addin         — Office.js add-in logic. Browser-ish globals, jsdom-free
 *                   because the tested code is pure.
 *
 * Run everything:            npm test
 * Run one project:           npx vitest --project core
 * Coverage (gated in CI):    npm run test:coverage
 */
import { defineConfig } from "vitest/config";

/** Attribute coverage to the package that owns the file, not the test file. */
const COVERAGE_INCLUDE = [
  "packages/core/src/**/*.ts",
  "packages/cli/src/**/*.ts",
  "packages/addin/src/**/*.ts",
];

/**
 * Exclusions are the part of a coverage setup most likely to be abused, so
 * every entry below has to justify itself. The rule used here: exclude a file
 * only when covering it in-process would require faking the very thing the
 * file exists to talk to (Office.js, a TTY, a child process). Anything that
 * is merely *inconvenient* to test stays in.
 */
const COVERAGE_EXCLUDE = [
  "**/*.test.ts",
  "**/*.d.ts",
  "**/dist/**",
  "**/build/**",
  "**/node_modules/**",
  "**/test/**",

  // Type-only and barrel modules: no executable statements to measure.
  "packages/core/src/types.ts",
  "packages/core/src/index.ts",
  "packages/core/src/exceljs-shim.d.ts",

  // Commander wiring and interactive prompts. Reachable only by running the
  // binary, which the CLI tests do — as a child process, so V8 in-process
  // coverage cannot see it. Counting them would report 0% for code that is
  // in fact tested, and pressure contributors into writing fake in-process
  // tests that assert nothing. See TESTING.md "Known gaps".
  "packages/cli/src/index.ts",
  "packages/cli/src/prompt.ts",
  "packages/cli/src/commands/addin.ts",
  "packages/cli/src/commands/login.ts",
  "packages/cli/src/commands/pqt.ts",
  "packages/cli/src/commands/profile.ts",
  "packages/cli/src/commands/run.ts",
  "packages/cli/src/commands/schedule.ts",

  // Office.js host integration: DOM + a live Excel host. Covered by the
  // manual protocol (TEST-PROTOCOL.md §§12-16), not by unit tests.
  "packages/addin/src/taskpane/**",
  "packages/addin/src/commands/**",
  "packages/addin/src/auth.ts",
  "packages/addin/src/excel.ts",
  "packages/addin/src/host.ts",
  "packages/addin/src/telemetry.ts",
];

export default defineConfig({
  test: {
    // Vitest resolves `./foo.js` specifiers to `./foo.ts` for TS sources, so
    // the existing NodeNext-style imports keep working unchanged.
    projects: [
      {
        extends: true,
        test: {
          name: "core",
          root: "packages/core",
          include: ["src/**/*.test.ts"],
          environment: "node",
        },
      },
      {
        extends: true,
        test: {
          name: "core-int",
          root: "packages/core",
          include: ["test/integration/**/*.test.ts"],
          environment: "node",
          // Integration tests spin up an in-process server per file; give
          // them a little more room than a unit test but still fail loudly
          // if something hangs waiting on a socket.
          testTimeout: 30_000,
        },
      },
      {
        extends: true,
        test: {
          name: "core-property",
          root: "packages/core",
          include: ["test/property/**/*.test.ts"],
          environment: "node",
          // Property runs are slower by construction; fast-check's own
          // per-property budget is set in test/property/setup.ts.
          testTimeout: 60_000,
        },
      },
      {
        extends: true,
        test: {
          name: "cli",
          root: "packages/cli",
          include: ["src/**/*.test.ts", "test/**/*.test.ts"],
          environment: "node",
          // Several CLI tests spawn the compiled binary to assert on real
          // exit codes and stderr. Build it first so `git clone && npm test`
          // works without a separate build step.
          globalSetup: ["./test/global-setup.ts"],
          // Child processes are slower than in-process assertions, and
          // Windows process spawn is slower still.
          testTimeout: 30_000,
        },
      },
      {
        extends: true,
        test: {
          name: "addin",
          root: "packages/addin",
          include: ["src/**/*.test.ts"],
          environment: "node",
        },
      },
    ],

    // Deterministic, CI-friendly defaults.
    clearMocks: true,
    mockReset: true,
    restoreMocks: true,
    unstubEnvs: true,
    unstubGlobals: true,

    // A test that logs nothing is a test that debugs badly; a test that logs
    // everything drowns CI. Keep stdout but attribute it to its test.
    printConsoleTrace: false,
    disableConsoleIntercept: false,

    reporters: process.env.CI
      ? [["default", { summary: false }], ["junit", { outputFile: "./reports/junit.xml" }], "github-actions"]
      : ["default"],

    coverage: {
      provider: "v8",
      include: COVERAGE_INCLUDE,
      exclude: COVERAGE_EXCLUDE,
      reporter: ["text-summary", "json-summary", "json", "lcov", "html"],
      reportsDirectory: "./coverage",
      // Report files that no test ever imports — otherwise an untested module
      // silently improves the average by being invisible.
      all: true,
      /**
       * Hard floors, set per area rather than globally — a single number
       * would either be too soft for the engine or too harsh for the CLI
       * shell, and the average would hide both.
       *
       * These are the "never below this, ever" net. The real gate is the
       * ratchet in scripts/coverage-gate.mjs, which compares against
       * .github/coverage-baseline.json and refuses to let a PR lower it.
       */
      thresholds: {
        // The mapping/coercion/load engine. Every bug here writes wrong data
        // to somebody's production Dataverse, so it carries the strict bar.
        "packages/core/src/**/*.ts": {
          lines: 78,
          statements: 75,
          functions: 82,
          branches: 68,
        },
        // CLI: auth, profiles, secure storage, scheduling. Lower because a
        // meaningful chunk is only reachable through a spawned process.
        "packages/cli/src/**/*.ts": {
          lines: 45,
          statements: 44,
          functions: 43,
          branches: 40,
        },
        // Add-in: only the pure suggestion/combobox logic is in scope.
        "packages/addin/src/**/*.ts": {
          lines: 34,
          statements: 37,
          functions: 32,
          branches: 44,
        },
      },
    },
  },
});
