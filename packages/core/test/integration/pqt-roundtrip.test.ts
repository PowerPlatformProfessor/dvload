/**
 * .pqt file format: read, write, and round-trip.
 *
 * WHY THIS FILE EXISTS
 * --------------------
 * `writePqt` and `readPqt` had zero test coverage. They are the entire I/O
 * boundary for the Power Query Template format:
 *
 *   dvload extract-pqt   xlsx  → PqtArchive → writePqt → .pqt   (export)
 *   dvload import-pqt    .pqt  → readPqt    → mappingFromPqt    (import)
 *
 * The existing pqt.test.ts covers the *semantics* (DeleteExistingDataOnLoad,
 * field-type translation) but always with a hand-built `PqtArchive` object.
 * Nothing ever checked that an archive survives being written to a zip and
 * read back — so a mistake in an entry name, an encoding, or the JSON
 * serialisation would produce a .pqt that Dataverse Dataflows rejects, and
 * every test would still pass.
 *
 * These tests use the real JSZip path in both directions.
 *
 * SCOPE NOTE: a .pqt is exported from the Dataverse Dataflow UI by hand and
 * handed to `dvload import-pqt`. There is no Web API call to fetch one, so
 * there is nothing here for the fake Dataverse server to serve — this is a
 * pure file-format concern.
 */

import { describe, it, expect } from "vitest";
import {
  writePqt,
  readPqt,
  extractPqtFromXlsx,
  injectMappingIntoPqt,
  mappingFromPqt,
  mappingsFromPqtAll,
  buildWorkbookWithQueries,
  parseQueryNames,
  STANDARD_CONTENT_TYPES,
  type PqtArchive,
} from "../../src/pqt.js";
import { parseMapping, type Mapping } from "../../src/mapping.js";

const ENV = "https://fake.crm.dynamics.com";

/** A realistic two-query archive, shaped like a Dataverse Dataflow export. */
function archive(over: Partial<PqtArchive> = {}): PqtArchive {
  return {
    mashupDocument:
      "section Section1;\r\n" +
      'shared Contacts = let Source = Excel.Workbook(File.Contents("C:\\data.xlsx")) in Source;\r\n' +
      'shared Accounts = let Source = Csv.Document(File.Contents("C:\\a.csv")) in Source;\r\n',
    mashupMetadata: {
      QueryGroups: [],
      DocumentLocale: "en-GB",
      FastCombine: false,
      AllowNativeQueries: false,
      QueriesMetadata: {
        Contacts: {
          QueryId: "11111111-1111-1111-1111-111111111111",
          QueryName: "Contacts",
          EntityName: "contacts",
          DeleteExistingDataOnLoad: false,
          FieldsMetadata: {
            emailaddress1: { SourceColumnName: "Email", DestinationFieldType: "String" },
            // "DateAndTime" is the Dataflow token, not "DateTime" — see
            // DataflowFieldType in src/pqt.ts.
            birthdate: { SourceColumnName: "DOB", DestinationFieldType: "DateAndTime" },
            creditlimit: { SourceColumnName: "Limit", DestinationFieldType: "Money" },
            // A type dvload doesn't know. Must degrade to string rather than
            // throwing: a .pqt comes from Microsoft's UI and may gain new
            // field types at any time.
            dvlt_note: { SourceColumnName: "Note", DestinationFieldType: "SomethingNew" },
          },
        },
        Accounts: {
          QueryId: "22222222-2222-2222-2222-222222222222",
          QueryName: "Accounts",
          EntityName: "accounts",
          DeleteExistingDataOnLoad: true,
          FieldsMetadata: {
            name: { SourceColumnName: "Name", DestinationFieldType: "String" },
          },
        },
      },
    },
    metadata: {
      Name: "Nightly contacts",
      Description: "Exported from a Dataverse Dataflow",
      Version: "1.0.0.0",
    },
    contentTypes: STANDARD_CONTENT_TYPES,
    ...over,
  } as PqtArchive;
}

function contactsMapping(): Mapping {
  return parseMapping({
    schemaVersion: 1,
    name: "contacts",
    environmentUrl: ENV,
    targetEntitySet: "contacts",
    sourceTable: "Contacts",
    columns: [
      { source: "Email", target: "emailaddress1", kind: "string", treatEmptyAsNull: true },
      { source: "DOB", target: "birthdate", kind: "dateonly", treatEmptyAsNull: true },
    ],
    conflictMode: "insert",
  });
}

/* -------------------------------------------------------------------------- */
/* Export: writePqt                                                            */
/* -------------------------------------------------------------------------- */

describe("writePqt (export)", () => {
  it("produces a zip containing exactly the four parts a .pqt must have", async () => {
    // Dataverse Dataflows import is strict about these names. Getting one
    // wrong yields a file that only fails once a user tries to import it.
    const { default: JSZip } = await import("jszip");
    const zip = await JSZip.loadAsync(await writePqt(archive()));

    expect(Object.keys(zip.files).sort()).toEqual(
      ["[Content_Types].xml", "MashupDocument.pq", "MashupMetadata.json", "Metadata.json"].sort()
    );
  });

  it("writes a real ZIP, not a bare buffer", async () => {
    const bytes = await writePqt(archive());
    // PK\x03\x04 — if compression or the output type regressed, this catches
    // it before anyone tries to open the file.
    expect([bytes[0], bytes[1], bytes[2], bytes[3]]).toEqual([0x50, 0x4b, 0x03, 0x04]);
    expect(bytes.byteLength).toBeGreaterThan(100);
  });

  it("stores the M document verbatim, including CRLF line endings", async () => {
    // Power Query's Advanced Editor is whitespace-sensitive, and a zip layer
    // is a classic place for line endings to get normalised.
    const a = archive();
    const back = await readPqt(await writePqt(a));
    expect(back.mashupDocument).toBe(a.mashupDocument);
    expect(back.mashupDocument).toContain("\r\n");
  });
});

/* -------------------------------------------------------------------------- */
/* Import: readPqt                                                             */
/* -------------------------------------------------------------------------- */

describe("readPqt (import)", () => {
  it("round-trips an archive without losing or reshaping anything", async () => {
    const original = archive();
    const back = await readPqt(await writePqt(original));

    expect(back).toEqual(original);
  });

  it("survives a second write/read cycle unchanged", async () => {
    // `dvload import-pqt` → edit → `extract-pqt` is a real workflow; drift
    // across cycles would accumulate silently.
    const once = await readPqt(await writePqt(archive()));
    const twice = await readPqt(await writePqt(once));

    expect(twice).toEqual(once);
  });

  it("preserves non-ASCII in query names, metadata and the M document", async () => {
    const a = archive({
      mashupDocument: 'section Section1;\r\nshared Küñdën = let x = "café ☕" in x;\r\n',
      metadata: { Name: "Kundendaten — täglich", Description: "Ürünler", Version: "1.0.0.0" },
    });

    const back = await readPqt(await writePqt(a));
    expect(back.mashupDocument).toBe(a.mashupDocument);
    expect(back.metadata.Name).toBe("Kundendaten — täglich");
    expect(back.metadata.Description).toBe("Ürünler");
  });

  it("names the missing part when the archive is incomplete", async () => {
    // A truncated or hand-assembled .pqt should say which part is missing,
    // not throw a JSZip internal error.
    const { default: JSZip } = await import("jszip");
    const zip = new JSZip();
    zip.file("MashupDocument.pq", "section Section1;");
    zip.file("Metadata.json", "{}");
    const bytes = await zip.generateAsync({ type: "uint8array" });

    await expect(readPqt(bytes)).rejects.toThrow(/MashupMetadata\.json/);
  });

  it("rejects a file that is not a zip at all", async () => {
    const notAZip = new TextEncoder().encode("this is not a .pqt");
    await expect(readPqt(notAZip)).rejects.toThrow();
  });

  it("accepts an ArrayBuffer as well as a Uint8Array", async () => {
    // The add-in reads files through FileReader and hands over an
    // ArrayBuffer; the CLI passes a Node Buffer. Both must work.
    const bytes = await writePqt(archive());
    const ab = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);

    const back = await readPqt(ab as ArrayBuffer);
    expect(back.metadata.Name).toBe("Nightly contacts");
  });
});

/* -------------------------------------------------------------------------- */
/* The workflows these functions actually serve                                */
/* -------------------------------------------------------------------------- */

describe("import-pqt: .pqt on disk → mapping", () => {
  it("derives a usable mapping after a real read from bytes", async () => {
    // The existing unit tests call mappingFromPqt on an in-memory object.
    // This is the path the CLI takes: bytes → readPqt → mappingFromPqt.
    const bytes = await writePqt(archive());
    const loaded = await readPqt(bytes);

    const mapping = mappingFromPqt(loaded, "Contacts", { environmentUrl: ENV });

    expect(mapping.targetEntitySet).toBe("contacts");
    expect(mapping.sourceTable).toBe("Contacts");
    expect(mapping.columns.map((c) => [c.source, c.target, c.kind])).toEqual([
      ["Email", "emailaddress1", "string"],
      ["DOB", "birthdate", "datetime"],
      ["Limit", "creditlimit", "money"],
      // Unknown Dataflow type degrades to string instead of throwing.
      ["Note", "dvlt_note", "string"],
    ]);
    // And it must be a mapping the parser accepts, not just a shaped object.
    expect(() => parseMapping(mapping)).not.toThrow();
  });

  it("carries the truncate-and-reload warning through the file boundary", async () => {
    const loaded = await readPqt(await writePqt(archive()));
    const mapping = mappingFromPqt(loaded, "Accounts", { environmentUrl: ENV });

    expect(mapping.conflictMode).toBe("insert");
    expect(mapping.description).toMatch(/DeleteExistingDataOnLoad/);
  });

  it("produces one valid mapping per query in the file", async () => {
    const loaded = await readPqt(await writePqt(archive()));
    // Keyed by query name, not an array — the CLI writes one
    // <queryName>.dvmap.json per entry.
    const mappings = mappingsFromPqtAll(loaded, { environmentUrl: ENV });

    expect(Object.keys(mappings).sort()).toEqual(["Accounts", "Contacts"]);
    expect(mappings.Contacts.sourceTable).toBe("Contacts");
    expect(mappings.Accounts.targetEntitySet).toBe("accounts");
    for (const m of Object.values(mappings)) expect(() => parseMapping(m)).not.toThrow();
  });

  it("skips queries present in the M document but absent from the metadata", async () => {
    // Real exports carry helper queries — staging steps, parameter tables —
    // that have no Dataverse destination. They must be dropped quietly, not
    // turned into a broken mapping.
    const a = archive();
    a.mashupDocument += "shared StagingHelper = let x = 1 in x;\r\n";

    const loaded = await readPqt(await writePqt(a));
    const mappings = mappingsFromPqtAll(loaded, { environmentUrl: ENV });

    expect(Object.keys(mappings).sort()).toEqual(["Accounts", "Contacts"]);
    expect(parseQueryNames(loaded.mashupDocument)).toContain("StagingHelper");
  });

  it("reports an unknown query by name rather than returning an empty mapping", async () => {
    const loaded = await readPqt(await writePqt(archive()));
    expect(() => mappingFromPqt(loaded, "NoSuchQuery", { environmentUrl: ENV })).toThrow(/NoSuchQuery/);
  });
});

describe("extract-pqt: workbook → .pqt with an injected mapping", () => {
  it("round-trips FieldsMetadata injected from a .dvmap.json", async () => {
    // This is what `dvload extract-pqt --mapping contacts.dvmap.json` does:
    // the resulting .pqt should import into Dataflows with the column
    // mapping already populated.
    const a = archive();
    injectMappingIntoPqt(a, contactsMapping());

    const back = await readPqt(await writePqt(a));
    const fields = back.mashupMetadata.QueriesMetadata.Contacts.FieldsMetadata;

    expect(fields).toBeDefined();
    expect(fields!.emailaddress1.SourceColumnName).toBe("Email");
    expect(back.mashupMetadata.QueriesMetadata.Contacts.EntityName).toBe("contacts");
  });

  it("never turns on truncate-and-reload as a side effect of injection", async () => {
    // Pinned across the file boundary too: DeleteExistingDataOnLoad wipes the
    // target table on every Dataflow run. dvload must never set it.
    const a = archive();
    a.mashupMetadata.QueriesMetadata.Contacts.DeleteExistingDataOnLoad = false;
    injectMappingIntoPqt(a, contactsMapping());

    const back = await readPqt(await writePqt(a));
    expect(back.mashupMetadata.QueriesMetadata.Contacts.DeleteExistingDataOnLoad).toBe(false);
  });

  it("survives the full xlsx → pqt → xlsx → pqt loop", async () => {
    // buildWorkbookWithQueries is marked EXPERIMENTAL, so this asserts the
    // M document specifically rather than byte equality of the workbook.
    const a = archive();
    const xlsx = await buildWorkbookWithQueries(a);
    const extracted = await extractPqtFromXlsx(xlsx, { name: "Round-tripped" });

    const back = await readPqt(await writePqt(extracted));

    expect(back.mashupDocument).toBe(a.mashupDocument);
    expect(back.metadata.Name).toBe("Round-tripped");
    expect(parseQueryNames(back.mashupDocument).sort()).toEqual(["Accounts", "Contacts"]);
  });
});
