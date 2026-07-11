// Tests for the DeleteExistingDataOnLoad / conflictMode semantics fix.

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  injectMappingIntoPqt,
  mappingFromPqt,
  STANDARD_CONTENT_TYPES,
  type PqtArchive,
} from "./pqt.js";
import type { Mapping } from "./mapping.js";

function archive(deleteExisting: boolean): PqtArchive {
  return {
    mashupDocument: 'section Section1;\nshared Contacts = let x = 1 in x;',
    mashupMetadata: {
      QueryGroups: [],
      DocumentLocale: "en-US",
      FastCombine: false,
      AllowNativeQueries: false,
      QueriesMetadata: {
        Contacts: {
          QueryId: "q1",
          QueryName: "Contacts",
          QueryGroupId: null,
          LoadEnabled: true,
          DeleteExistingDataOnLoad: deleteExisting,
          FieldsMetadata: {
            emailaddress1: {
              SourceColumnName: "Email",
              DestinationFieldType: "String",
              AutoNumberSettings: null,
            },
          },
        },
      },
    },
    metadata: { Name: "Contacts import", Description: "", Version: "1.0.0.0" },
    contentTypes: STANDARD_CONTENT_TYPES,
  };
}

function upsertMapping(): Mapping {
  return {
    schemaVersion: 1,
    name: "t",
    environmentUrl: "https://unit.crm.dynamics.com",
    targetEntitySet: "contacts",
    sourceTable: "Contacts",
    columns: [{ source: "Email", target: "emailaddress1", kind: "string", treatEmptyAsNull: true }],
    conflictMode: "upsert",
    upsertKey: ["emailaddress1"],
    batchSize: 100,
    maxErrors: 0,
    logDir: "./logs",
  };
}

test("injectMappingIntoPqt never sets DeleteExistingDataOnLoad (truncate != upsert)", () => {
  const a = archive(false);
  injectMappingIntoPqt(a, upsertMapping());
  assert.equal(a.mashupMetadata.QueriesMetadata["Contacts"].DeleteExistingDataOnLoad, false);
  assert.equal(a.mashupMetadata.QueriesMetadata["Contacts"].LoadEnabled, true);
});

test("mappingFromPqt emits a valid insert mapping for truncate-and-reload Dataflows", () => {
  const m = mappingFromPqt(archive(true), "Contacts", {
    environmentUrl: "https://unit.crm.dynamics.com",
  });
  // Previously: conflictMode="upsert" with no upsertKey → failed validation on first run.
  assert.equal(m.conflictMode, "insert");
  assert.equal(m.upsertKey, undefined);
  assert.match(m.description ?? "", /DeleteExistingDataOnLoad/);
  assert.equal(m.columns.length, 1);
  assert.equal(m.columns[0].source, "Email");
});

test("mappingFromPqt adds no note for plain Dataflows", () => {
  const m = mappingFromPqt(archive(false), "Contacts", {
    environmentUrl: "https://unit.crm.dynamics.com",
  });
  assert.equal(m.conflictMode, "insert");
  assert.equal(m.description, undefined);
});
