/**
 * Test data builders.
 *
 * Integration tests read best when the setup states only what the test is
 * about; everything else should be a sane default that never changes the
 * meaning of the assertion. These builders exist so a test that cares about
 * `conflictMode` doesn't also have to spell out `logDir`.
 */

import { parseMapping, type Mapping, type ColumnMapping } from "../../src/mapping.js";
import { DataverseClient } from "../../src/dataverse.js";
import type { SourceRow } from "../../src/types.js";
import type { FakeDataverse } from "./fake-dataverse.js";

export const TEST_ENV = "https://fake.crm.dynamics.com";

/** A minimal, valid contact mapping. Overrides are shallow-merged. */
export function contactMapping(overrides: Partial<Mapping> = {}): Mapping {
  return parseMapping({
    schemaVersion: 1,
    name: "integration",
    environmentUrl: TEST_ENV,
    targetEntitySet: "contacts",
    sourceTable: "tblContacts",
    columns: [
      col({ source: "Email", target: "emailaddress1", kind: "string" }),
      col({ source: "First", target: "firstname", kind: "string" }),
      col({ source: "Last", target: "lastname", kind: "string" }),
    ],
    conflictMode: "insert",
    batchSize: 100,
    maxErrors: 0,
    logDir: "./logs",
    concurrency: 1,
    bypassCustomLogic: false,
    skipUnchanged: false,
    ...overrides,
  });
}

export function col(c: Partial<ColumnMapping> & Pick<ColumnMapping, "target" | "kind">): ColumnMapping {
  return { treatEmptyAsNull: true, ...c } as ColumnMapping;
}

/** N deterministic contact rows: contact0@example.test, contact1@… */
export function contactRows(n: number, start = 0): SourceRow[] {
  return Array.from({ length: n }, (_, i) => ({
    Email: `contact${start + i}@example.test`,
    First: `First${start + i}`,
    Last: "DVLT-Load",
  }));
}

/** A real DataverseClient wired to the fake's transport. */
export function clientFor(
  fake: FakeDataverse,
  opts: Partial<ConstructorParameters<typeof DataverseClient>[0]> = {}
): DataverseClient {
  return new DataverseClient({
    environmentUrl: fake.environmentUrl,
    getToken: async () => "fake-token",
    fetch: fake.fetch,
    // Keep retries fast: the point of a retry test is the decision, not the
    // wall-clock sleep. Individual tests override this when they assert on
    // Retry-After handling specifically.
    retry: { baseDelayMs: 1, maxDelayMs: 5 },
    ...opts,
  });
}
