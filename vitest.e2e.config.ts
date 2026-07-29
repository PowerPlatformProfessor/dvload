/**
 * Live end-to-end tests against a REAL Dataverse environment.
 *
 * Separate config, separate command, never part of `npm test`. Running these
 * writes records to an actual environment, so they must never fire by
 * accident — including from a fork PR that edits vitest.config.ts.
 *
 * Required environment (see .github/workflows/e2e-live.yml):
 *   DVLOAD_E2E_ENV_URL    https://<org>.crm.dynamics.com  (a SANDBOX)
 *   DVLOAD_E2E_CLIENT_ID  app registration client id
 *   DVLOAD_E2E_TENANT_ID  tenant id
 *   DVLOAD_E2E_SECRET     client secret
 *
 * With any of those missing, every test self-skips with a reason rather than
 * failing — so a contributor running `npm run test:e2e` locally gets a clear
 * "not configured" instead of a wall of auth errors.
 */
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    name: "e2e-live",
    root: "packages/core",
    include: ["test/e2e/**/*.e2e.test.ts"],
    environment: "node",
    // Real network, real throttling, real batches of a few thousand rows.
    testTimeout: 15 * 60 * 1000,
    hookTimeout: 5 * 60 * 1000,
    // Never parallelise: concurrent runs against one environment interfere
    // with each other's cleanup and produce failures nobody can reproduce.
    fileParallelism: false,
    pool: "forks",
    poolOptions: { forks: { singleFork: true } },
    retry: 0,
    reporters: process.env.CI
      ? ["default", ["junit", { outputFile: "./reports/junit-e2e.xml" }]]
      : ["default"],
  },
});
