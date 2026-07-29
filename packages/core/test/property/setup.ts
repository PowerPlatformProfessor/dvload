/**
 * Shared fast-check configuration.
 *
 * Property tests are only useful if a failure is reproducible. Seeding from
 * the environment (with a fixed fallback) means:
 *   - CI reruns are deterministic, so a red build stays red until fixed;
 *   - the nightly mutation/fuzz job can pass FC_SEED=$RANDOM to explore new
 *     ground without making every PR flaky;
 *   - a reported counterexample can be replayed locally with
 *     `FC_SEED=<seed> npm run test:property`.
 */

import fc from "fast-check";

const seed = process.env.FC_SEED ? Number(process.env.FC_SEED) : 0x64766c64;

fc.configureGlobal({
  seed,
  numRuns: process.env.CI ? 500 : 100,
  // Print the full counterexample rather than a truncated one — these values
  // are the whole point of the failure report.
  verbose: fc.VerbosityLevel.Verbose,
  // Interrupt rather than time out the whole Vitest file if a property is
  // pathologically slow; the report then names the property.
  interruptAfterTimeLimit: 20_000,
  markInterruptAsFailure: true,
});

export { fc, seed };
