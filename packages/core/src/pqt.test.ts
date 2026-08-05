// Tests for the DeleteExistingDataOnLoad / conflictMode semantics fix.

import { test } from "vitest";
import assert from "node:assert/strict";
import {
  buildWorkbookWithQueries,
  extractPqtFromXlsx,
  injectMappingIntoPqt,
  mappingFromPqt,
  mappingsFromPqtAll,
  normalizeQueriesMetadata,
  readPqt,
  writePqt,
  STANDARD_CONTENT_TYPES,
  type PqtArchive,
} from "./pqt.js";
import JSZip from "jszip";
import type { Mapping } from "./mapping.js";

function archive(deleteExisting: boolean): PqtArchive {
  return {
    mashupDocument: "section Section1;\nshared Contacts = let x = 1 in x;",
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
    concurrency: 1,
    bypassCustomLogic: false,
    skipUnchanged: false,
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

test("mappingsFromPqtAll returns one mapping per query", () => {
  const a = archive(false);
  a.mashupMetadata.QueriesMetadata["Accounts"] = {
    QueryId: "q2",
    QueryName: "Accounts",
    QueryGroupId: null,
    LoadEnabled: false,
    FieldsMetadata: {},
  };
  const all = mappingsFromPqtAll(a, { environmentUrl: "https://unit.crm.dynamics.com" });
  assert.deepEqual(Object.keys(all).sort(), ["Accounts", "Contacts"]);
  assert.equal(all["Contacts"].columns.length, 1);
  assert.equal(all["Accounts"].columns.length, 0);
});

/* -------------------------------------------------------------------------- */
/* MashupMetadata.QueriesMetadata: array (Power Query Online) vs object        */
/* (Dataverse Dataflows). Reading the array form as an object used to name     */
/* every query "0", because Object.keys([{…}]) === ["0"].                      */
/* -------------------------------------------------------------------------- */

const PQ_ONLINE_M =
  "section Section1;\nshared Full_Match_Table = let\n" +
  '  Source = Excel.Workbook(Web.Contents("https://example/x.xlsx"), null, true)\n' +
  "in\n  Source;\n";

/** A .pqt as Power Query Online writes it: QueriesMetadata is an array. */
async function pqOnlinePqtBytes(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file("MashupDocument.pq", PQ_ONLINE_M);
  zip.file(
    "MashupMetadata.json",
    JSON.stringify({
      DocumentLocale: "en-US",
      EngineVersion: "2.156.778.0",
      QueriesMetadata: [
        {
          QueryName: "Full_Match_Table",
          QueryGroupId: null,
          LastKnownIsParameter: false,
          LastKnownResultTypeName: null,
          LoadEnabled: true,
          IsHidden: null,
        },
      ],
      QueryGroups: [],
    })
  );
  zip.file(
    "Metadata.json",
    JSON.stringify({ Name: "1. Import AI match Accounts", Description: "", Version: "1.0.0.0" })
  );
  zip.file("[Content_Types].xml", STANDARD_CONTENT_TYPES);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

test("normalizeQueriesMetadata keys an array by QueryName, not by index", () => {
  const { queriesMetadata, shape } = normalizeQueriesMetadata([
    { QueryName: "Full_Match_Table", LoadEnabled: true },
    { QueryName: "Staging" },
  ]);
  assert.equal(shape, "array");
  assert.deepEqual(Object.keys(queriesMetadata).sort(), ["Full_Match_Table", "Staging"]);
  assert.equal(queriesMetadata["Full_Match_Table"].LoadEnabled, true);
});

test("normalizeQueriesMetadata passes the keyed-object form through", () => {
  const { queriesMetadata, shape } = normalizeQueriesMetadata({
    Contacts: { QueryName: "Contacts" },
  });
  assert.equal(shape, "object");
  assert.deepEqual(Object.keys(queriesMetadata), ["Contacts"]);
});

test("normalizeQueriesMetadata tolerates a missing or null QueriesMetadata", () => {
  for (const raw of [null, undefined, 42]) {
    const { queriesMetadata, shape } = normalizeQueriesMetadata(raw);
    assert.deepEqual(queriesMetadata, {});
    assert.equal(shape, "object");
  }
});

test("normalizeQueriesMetadata keeps unnamed array entries instead of dropping them", () => {
  const { queriesMetadata } = normalizeQueriesMetadata([{ LoadEnabled: true }, { QueryName: "" }]);
  assert.deepEqual(Object.keys(queriesMetadata).sort(), ["Query1", "Query2"]);
});

test("readPqt names a Power Query Online query correctly (regression: was '0')", async () => {
  const a = await readPqt(await pqOnlinePqtBytes());
  assert.deepEqual(Object.keys(a.mashupMetadata.QueriesMetadata), ["Full_Match_Table"]);
  assert.equal(a.queriesMetadataShape, "array");
  // The old bug surfaced here: sourceTable came out as "0".
  const m = mappingFromPqt(a, "Full_Match_Table", {
    environmentUrl: "https://unit.crm.dynamics.com",
  });
  assert.equal(m.sourceTable, "Full_Match_Table");
  // No FieldsMetadata in a PQ Online export — there is nothing to map yet.
  assert.equal(m.columns.length, 0);
});

test("writePqt preserves the array shape it was read in", async () => {
  const a = await readPqt(await pqOnlinePqtBytes());
  const zip = await JSZip.loadAsync(await writePqt(a));
  const meta = JSON.parse(await zip.file("MashupMetadata.json")!.async("string"));
  assert.ok(Array.isArray(meta.QueriesMetadata), "expected QueriesMetadata to stay an array");
  assert.equal(meta.QueriesMetadata[0].QueryName, "Full_Match_Table");
  // Keys we don't model must survive the round-trip.
  assert.equal(meta.EngineVersion, "2.156.778.0");
});

test("injectMappingIntoPqt promotes an array-shaped .pqt to the Dataflows form", async () => {
  const a = await readPqt(await pqOnlinePqtBytes());
  injectMappingIntoPqt(a, { ...upsertMapping(), sourceTable: "Full_Match_Table" });
  assert.equal(a.queriesMetadataShape, "object");
  const q = a.mashupMetadata.QueriesMetadata["Full_Match_Table"];
  assert.equal(q.FieldsMetadata?.emailaddress1.SourceColumnName, "Email");
  assert.ok(q.QueryId, "a Dataflows-bound query needs a QueryId");

  const zip = await JSZip.loadAsync(await writePqt(a));
  const meta = JSON.parse(await zip.file("MashupMetadata.json")!.async("string"));
  assert.ok(!Array.isArray(meta.QueriesMetadata));
  assert.equal(meta.QueriesMetadata.Full_Match_Table.QueryName, "Full_Match_Table");
});

test("mappingsFromPqtAll uses real query names for array-shaped .pqt", async () => {
  const a = await readPqt(await pqOnlinePqtBytes());
  const all = mappingsFromPqtAll(a, { environmentUrl: "https://unit.crm.dynamics.com" });
  assert.deepEqual(Object.keys(all), ["Full_Match_Table"]);
});

test("workbook DataMashup Config/Package.xml stays in the shape Excel's QDEFF reader accepts", async () => {
  // Verified against Excel via COM (2026-08): Excel deserializes this part
  // strictly and rejects the entire mashup — workbook opens with an empty
  // Queries pane — if the <Package> root carries a default xmlns (even the
  // DataMashup namespace) or contains elements it doesn't model, such as
  // <SafeCombine>. Both regressions shipped once; keep this pinned.
  const xlsxBytes = await buildWorkbookWithQueries(archive(false));
  const xlsx = await JSZip.loadAsync(xlsxBytes);
  const item1 = await xlsx.file("customXml/item1.xml")!.async("string");
  const b64 = /<DataMashup[^>]*>([\s\S]*?)<\/DataMashup>/.exec(item1)![1];
  const blob = Buffer.from(b64, "base64");
  const packageLen = blob.readUInt32LE(4);
  const inner = await JSZip.loadAsync(blob.subarray(8, 8 + packageLen));
  const pkgXml = await inner.file("Config/Package.xml")!.async("string");
  assert.doesNotMatch(pkgXml, /xmlns="/, "Package root must be in the empty namespace");
  assert.doesNotMatch(pkgXml, /SafeCombine/, "Excel rejects unknown elements like <SafeCombine>");
  assert.match(pkgXml, /<Version>.*<\/Version>/);
  assert.match(pkgXml, /<MinVersion>.*<\/MinVersion>/);
  assert.match(pkgXml, /<Culture>.*<\/Culture>/);
});

test("EXPERIMENTAL: pqt → workbook → extract round-trips the M document", async () => {
  const a = archive(false);
  a.mashupDocument =
    'section Section1;\r\nshared Contacts = let Source = Csv.Document("x") in Source;\r\nshared #"My Query" = 1 + 1;';
  const xlsxBytes = await buildWorkbookWithQueries(a);

  // The workbook must contain a DataMashup part our own parser accepts.
  const roundTripped = await extractPqtFromXlsx(xlsxBytes);
  assert.equal(roundTripped.mashupDocument, a.mashupDocument);
  assert.deepEqual(Object.keys(roundTripped.mashupMetadata.QueriesMetadata).sort(), ["Contacts", "My Query"]);
});
