// Power Query Template (.pqt) codec for dvload.
//
// .pqt is the format Dataverse Dataflows and Power Query Online emit.
// It's a plain ZIP containing:
//
//   MashupDocument.pq      The M code, identical to Excel's Formulas/Section1.m
//   MashupMetadata.json    Per-query metadata + Dataverse FieldsMetadata
//   Metadata.json          Name, Description, Version
//   [Content_Types].xml    Standard OOXML content types
//
// Excel doesn't store this format directly — instead it stores Power
// Queries inside `customXml/item*.xml` as a base64 "DataMashup" blob
// (binary QDEFF wrapper around an inner ZIP that holds Section1.m). To
// produce a .pqt from an Excel workbook we:
//
//   1. Open the .xlsx as a ZIP
//   2. Find the customXml part with <DataMashup>
//   3. Base64-decode and parse the QDEFF header to slice out the inner ZIP
//   4. Read Formulas/Section1.m from the inner ZIP
//   5. Synthesize MashupMetadata.json, Metadata.json, [Content_Types].xml
//   6. Re-zip into a .pqt
//
// The QDEFF format is documented in Microsoft Open Specifications: MS-QDEFF.

import JSZip from "jszip";
import type { Mapping, ColumnMapping } from "./mapping.js";
import type { DataverseFieldKind } from "./types.js";

/* -------------------------------------------------------------------------- */
/* Types                                                                       */
/* -------------------------------------------------------------------------- */

export interface PqtArchive {
  mashupDocument: string;
  mashupMetadata: MashupMetadata;
  metadata: PqtTopMetadata;
  /** [Content_Types].xml. Stored verbatim so we can round-trip. */
  contentTypes: string;
}

export interface PqtTopMetadata {
  Name: string;
  Description: string;
  Version: string;
}

/**
 * The MashupMetadata.json shape. We model only the fields we read or write —
 * unknown keys are preserved through `extra` for round-trip fidelity.
 */
export interface MashupMetadata {
  QueryGroups: unknown[];
  DocumentLocale: string;
  QueriesMetadata: Record<string, QueryMetadata>;
  FastCombine: boolean;
  AllowNativeQueries: boolean;
  HostContext?: { Type: string; Details?: Record<string, unknown> };
  /** Bag for fields we don't model. Preserved on round-trip. */
  extra?: Record<string, unknown>;
}

export interface QueryMetadata {
  QueryId: string;
  QueryName: string;
  QueryGroupId: string | null;
  EntityName?: string | null;
  LastKnownIsCalculatedEntity?: boolean;
  LastKnownIsLinkedEntity?: boolean;
  LastKnownIsParameter?: boolean;
  IsHidden?: boolean | null;
  LastKnownResultTypeName?: string | null;
  LoadEnabled?: boolean;
  DeleteExistingDataOnLoad?: boolean;
  FieldsMetadata?: Record<string, FieldMetadata>;
  HostProperties?: Record<string, unknown>;
  JsonOutputDestinations?: unknown;
  BindToDefaultOutputDestination?: unknown;
  JsonIncrementalRefreshSettings?: unknown;
  JsonStagingDefinition?: unknown;
}

/**
 * Subset of Dataflow destination field types relevant to mappings.
 * Matches the literals seen in Dataverse Dataflows .pqt exports.
 */
export type DataflowFieldType =
  | "Memo"
  | "String"
  | "Integer"
  | "Decimal"
  | "Double"
  | "Boolean"
  | "DateAndTime"
  | "DateOnly"
  | "Money"
  | "UniqueIdentifier"
  | "Choice"
  | "Lookup";

export interface FieldMetadata {
  SourceColumnName: string;
  DestinationFieldType: DataflowFieldType | string;
  AutoNumberSettings: unknown | null;
}

/* -------------------------------------------------------------------------- */
/* Constants                                                                   */
/* -------------------------------------------------------------------------- */

export const STANDARD_CONTENT_TYPES =
  '<?xml version="1.0" encoding="utf-8"?>\n' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n' +
  '  <Default Extension="pq" ContentType="text/plain" />\n' +
  '  <Default Extension="json" ContentType="application/json" />\n' +
  '  <Default Extension="xml" ContentType="application/xml" />\n' +
  "</Types>";

/* -------------------------------------------------------------------------- */
/* Public API                                                                  */
/* -------------------------------------------------------------------------- */

export interface ExtractOptions {
  /** Override the Metadata.json Name field (default: "Extracted from Excel"). */
  name?: string;
  description?: string;
  /** Default "1.0.0.0". */
  version?: string;
  /** Document locale (default "en-US"). */
  locale?: string;
}

/** Extract Power Query M from an Excel workbook and package it as a .pqt archive. */
export async function extractPqtFromXlsx(
  xlsxBytes: Uint8Array | ArrayBuffer,
  opts: ExtractOptions = {}
): Promise<PqtArchive> {
  const xlsxZip = await JSZip.loadAsync(xlsxBytes);

  const dataMashupB64 = await findDataMashup(xlsxZip);
  if (!dataMashupB64) {
    throw new Error(
      "No Power Query found in this workbook (no customXml/<DataMashup> entry). " +
        "If you authored the queries entirely in Excel for the web, the binary " +
        "DataMashup blob may not have been written yet — open and save the file " +
        "in Excel desktop and try again."
    );
  }

  const innerPackageZip = parseQdeffPackage(base64Decode(dataMashupB64));
  const innerZip = await JSZip.loadAsync(innerPackageZip);

  const sectionFile =
    innerZip.file("Formulas/Section1.m") ?? innerZip.file("Formulas/Section1.pq");
  if (!sectionFile) {
    throw new Error(
      "DataMashup package has no Formulas/Section1.m. The workbook may be using " +
        "an unsupported variant of the format."
    );
  }
  const mashupDocument = await sectionFile.async("string");

  const queryNames = parseQueryNames(mashupDocument);

  return {
    mashupDocument,
    mashupMetadata: makeMinimalMashupMetadata(queryNames, opts.locale ?? "en-US"),
    metadata: {
      Name: opts.name ?? "Extracted from Excel",
      Description: opts.description ?? "",
      Version: opts.version ?? "1.0.0.0",
    },
    contentTypes: STANDARD_CONTENT_TYPES,
  };
}

/** Serialize a PqtArchive to a .pqt file (ZIP buffer). */
export async function writePqt(archive: PqtArchive): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file("MashupDocument.pq", archive.mashupDocument);
  zip.file("MashupMetadata.json", JSON.stringify(archive.mashupMetadata));
  zip.file("Metadata.json", JSON.stringify(archive.metadata));
  zip.file("[Content_Types].xml", archive.contentTypes);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

/** Read a .pqt file back into the in-memory representation. */
export async function readPqt(pqtBytes: Uint8Array | ArrayBuffer): Promise<PqtArchive> {
  const zip = await JSZip.loadAsync(pqtBytes);
  const must = (name: string) => {
    const f = zip.file(name);
    if (!f) throw new Error(`Missing ${name} in .pqt`);
    return f;
  };
  const [doc, meta, top, ct] = await Promise.all([
    must("MashupDocument.pq").async("string"),
    must("MashupMetadata.json").async("string"),
    must("Metadata.json").async("string"),
    must("[Content_Types].xml").async("string"),
  ]);
  return {
    mashupDocument: doc,
    mashupMetadata: JSON.parse(meta) as MashupMetadata,
    metadata: JSON.parse(top) as PqtTopMetadata,
    contentTypes: ct,
  };
}

/**
 * Inject FieldsMetadata into the .pqt's MashupMetadata using a Mapping. The
 * resulting .pqt imports into Dataverse Dataflows with the column-to-attribute
 * mapping pre-populated.
 */
export function injectMappingIntoPqt(archive: PqtArchive, mapping: Mapping): void {
  const queryName = mapping.sourceTable;
  const q = archive.mashupMetadata.QueriesMetadata[queryName];
  if (!q) {
    throw new Error(
      `Query "${queryName}" not found in this .pqt. ` +
        `Available: ${Object.keys(archive.mashupMetadata.QueriesMetadata).join(", ") || "(none)"}`
    );
  }
  q.LoadEnabled = true;
  // Never set DeleteExistingDataOnLoad from conflictMode: in Dataflows it
  // means truncate-and-reload, which is not upsert semantics. Setting it
  // for an upsert mapping would make the Dataflow wipe the table each run.
  q.DeleteExistingDataOnLoad = false;
  q.EntityName = q.EntityName ?? mapping.targetEntitySet;
  q.FieldsMetadata = q.FieldsMetadata ?? {};
  for (const col of mapping.columns) {
    // Constant columns have no Power Query source column to map.
    if (col.source === undefined) continue;
    q.FieldsMetadata[col.target] = {
      SourceColumnName: col.source,
      DestinationFieldType: dataverseKindToDataflowType(col.kind),
      AutoNumberSettings: null,
    };
  }
}

/**
 * Build a Mapping from a .pqt by reading FieldsMetadata for one query.
 * Useful as a migration path from an existing Dataverse Dataflow.
 */
export function mappingFromPqt(
  archive: PqtArchive,
  queryName: string,
  opts: { environmentUrl: string; targetEntitySet?: string }
): Mapping {
  const q = archive.mashupMetadata.QueriesMetadata[queryName];
  if (!q) throw new Error(`Query "${queryName}" not found in .pqt`);

  const columns: ColumnMapping[] = Object.entries(q.FieldsMetadata ?? {}).map(
    ([target, fm]) => ({
      source: fm.SourceColumnName,
      target,
      kind: dataflowTypeToDataverseKind(String(fm.DestinationFieldType)),
      treatEmptyAsNull: true,
    })
  );

  // DeleteExistingDataOnLoad (truncate-and-reload) has no dvload equivalent
  // and is NOT upsert — mapping it to "upsert" also produced an invalid
  // mapping (upsert requires upsertKey, which a .pqt can't supply). Always
  // emit insert and note the source behavior so the user can decide.
  const description = [
    archive.metadata.Description || undefined,
    q.DeleteExistingDataOnLoad
      ? "NOTE: the source Dataflow used DeleteExistingDataOnLoad (truncate-and-reload). " +
        "dvload has no equivalent; this mapping was set to conflictMode=insert. " +
        "If you need update semantics, set conflictMode=upsert and define upsertKey."
      : undefined,
  ]
    .filter(Boolean)
    .join(" ");

  return {
    schemaVersion: 1,
    name: archive.metadata.Name || queryName,
    description: description || undefined,
    environmentUrl: opts.environmentUrl,
    targetEntitySet: opts.targetEntitySet ?? q.EntityName ?? queryName,
    sourceTable: queryName,
    columns,
    conflictMode: "insert",
    batchSize: 100,
    maxErrors: 0,
    logDir: "./logs",
    concurrency: 1,
    bypassCustomLogic: false,
    skipUnchanged: false,
  };
}

/* -------------------------------------------------------------------------- */
/* QDEFF / DataMashup writing (EXPERIMENTAL)                                   */
/*                                                                             */
/* The reverse of the parser below: build a binary DataMashup blob per        */
/* MS-QDEFF and embed it in a fresh .xlsx so Excel's Power Query editor       */
/* shows the queries natively (as connection-only queries; the user picks     */
/* "Load To…" per query). Excel is strict about these parts — treat output    */
/* as experimental and keep the paste-into-Advanced-Editor fallback in mind.  */
/* -------------------------------------------------------------------------- */

const QDEFF_PACKAGE_XML =
  '<?xml version="1.0" encoding="utf-8"?>' +
  '<Package xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
  'xmlns="http://schemas.microsoft.com/DataMashup">' +
  "<Version>2.72.5556.181</Version><MinVersion>2.21.0.0</MinVersion>" +
  "<Culture>en-US</Culture><SafeCombine>true</SafeCombine></Package>";

const QDEFF_PERMISSIONS_XML =
  '<?xml version="1.0" encoding="utf-8"?>' +
  '<PermissionList xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
  "<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>" +
  "<FirewallEnabled>true</FirewallEnabled>" +
  '<WorkbookGroupType xsi:nil="true" /></PermissionList>';

const QDEFF_PACKAGE_CONTENT_TYPES =
  '<?xml version="1.0" encoding="utf-8"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="xml" ContentType="text/xml" />' +
  '<Default Extension="m" ContentType="application/x-ms-m" /></Types>';

function xmlEscape(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function u32le(n: number): Uint8Array {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, n, true);
  return b;
}

function concatBytes(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((n, p) => n + p.byteLength, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.byteLength;
  }
  return out;
}

const utf8 = {
  encode: (s: string): Uint8Array => new TextEncoder().encode(s),
};

/** Build the LocalPackageMetadataFile XML: one AllFormulas item + one per query. */
function buildMetadataXml(queryNames: string[]): string {
  const items: string[] = [
    "<Item><ItemLocation><ItemType>AllFormulas</ItemType><ItemPath /></ItemLocation><StableEntries /></Item>",
  ];
  for (const name of queryNames) {
    const itemPath = xmlEscape("Section1/" + encodeURIComponent(name));
    items.push(
      "<Item><ItemLocation><ItemType>Formula</ItemType>" +
        `<ItemPath>${itemPath}</ItemPath></ItemLocation>` +
        "<StableEntries>" +
        '<Entry Type="IsPrivate" Value="l0" />' +
        '<Entry Type="FillEnabled" Value="l0" />' +
        '<Entry Type="ResultType" Value="sTable" />' +
        "</StableEntries></Item>"
    );
  }
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<LocalPackageMetadataFile xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<Items>${items.join("")}</Items></LocalPackageMetadataFile>`
  );
}

/**
 * Serialize a PqtArchive into a binary DataMashup (QDEFF) blob:
 *   version | packageLen | package(zip) | permissionsLen | permissions |
 *   metadataLen | metadata | bindingsLen | bindings
 * Metadata itself is: version | xmlLen | xml | contentLen | content(zip).
 */
export async function writeDataMashup(archive: PqtArchive): Promise<Uint8Array> {
  // Inner OPC package with the M document.
  const pkg = new JSZip();
  pkg.file("[Content_Types].xml", QDEFF_PACKAGE_CONTENT_TYPES);
  pkg.file("Config/Package.xml", QDEFF_PACKAGE_XML);
  pkg.file("Formulas/Section1.m", archive.mashupDocument);
  const pkgBytes = await pkg.generateAsync({ type: "uint8array", compression: "DEFLATE" });

  const permissions = utf8.encode(QDEFF_PERMISSIONS_XML);

  const queryNames = Object.keys(archive.mashupMetadata.QueriesMetadata);
  const namesForMetadata =
    queryNames.length > 0 ? queryNames : parseQueryNames(archive.mashupDocument);
  const metadataXml = utf8.encode(buildMetadataXml(namesForMetadata));
  const emptyContent = await new JSZip().generateAsync({ type: "uint8array" });
  const metadata = concatBytes([
    u32le(0), // metadata version
    u32le(metadataXml.byteLength),
    metadataXml,
    u32le(emptyContent.byteLength),
    emptyContent,
  ]);

  return concatBytes([
    u32le(0), // QDEFF version
    u32le(pkgBytes.byteLength),
    pkgBytes,
    u32le(permissions.byteLength),
    permissions,
    u32le(metadata.byteLength),
    metadata,
    u32le(0), // no permission bindings
  ]);
}

/* ----- Minimal .xlsx scaffolding with the DataMashup custom XML part ------ */

const XLSX_RELS_ROOT =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />' +
  "</Relationships>";

const XLSX_CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />' +
  '<Default Extension="xml" ContentType="application/xml" />' +
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />' +
  '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />' +
  '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml" />' +
  "</Types>";

const XLSX_WORKBOOK =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" /></sheets></workbook>';

const XLSX_WORKBOOK_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml" />' +
  "</Relationships>";

const XLSX_SHEET1 =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData /></worksheet>';

const CUSTOMXML_ITEM_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml" />' +
  "</Relationships>";

function customXmlItemProps(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<ds:datastoreItem ds:itemID="{${cryptoUuid().toUpperCase()}}" ` +
    'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
    '<ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.microsoft.com/DataMashup" /></ds:schemaRefs>' +
    "</ds:datastoreItem>"
  );
}

function base64Encode(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }
  let binary = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  }
  return btoa(binary);
}

/**
 * EXPERIMENTAL: build a fresh .xlsx whose Power Query editor contains all
 * queries from the .pqt (connection-only; use "Load To…" in Excel to land
 * them on sheets). Returns the workbook bytes.
 */
export async function buildWorkbookWithQueries(archive: PqtArchive): Promise<Uint8Array> {
  const mashup = await writeDataMashup(archive);
  const item1 =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">' +
    base64Encode(mashup) +
    "</DataMashup>";

  const zip = new JSZip();
  zip.file("[Content_Types].xml", XLSX_CONTENT_TYPES);
  zip.file("_rels/.rels", XLSX_RELS_ROOT);
  zip.file("xl/workbook.xml", XLSX_WORKBOOK);
  zip.file("xl/_rels/workbook.xml.rels", XLSX_WORKBOOK_RELS);
  zip.file("xl/worksheets/sheet1.xml", XLSX_SHEET1);
  zip.file("customXml/item1.xml", item1);
  zip.file("customXml/itemProps1.xml", customXmlItemProps());
  zip.file("customXml/_rels/item1.xml.rels", CUSTOMXML_ITEM_RELS);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

/**
 * Build one Mapping per query in the .pqt (queries without FieldsMetadata
 * yield mappings with empty columns). Keyed by query name.
 */
export function mappingsFromPqtAll(
  archive: PqtArchive,
  opts: { environmentUrl: string }
): Record<string, Mapping> {
  const out: Record<string, Mapping> = {};
  const names = Object.keys(archive.mashupMetadata.QueriesMetadata);
  const all = names.length > 0 ? names : parseQueryNames(archive.mashupDocument);
  for (const name of all) {
    try {
      out[name] = mappingFromPqt(archive, name, { environmentUrl: opts.environmentUrl });
    } catch {
      // query listed in M but absent from metadata — skip
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* QDEFF / DataMashup parsing                                                  */
/* -------------------------------------------------------------------------- */

/**
 * Find and return the base64 contents of <DataMashup>...</DataMashup> in any
 * customXml part. Excel may put it in item1.xml, item2.xml, etc.
 */
async function findDataMashup(xlsxZip: JSZip): Promise<string | null> {
  const candidates = Object.values(xlsxZip.files).filter(
    (f) => /^customXml\/item\d+\.xml$/i.test(f.name) && !f.dir
  );
  // Most workbooks have it in item1, but iterate all to be safe.
  for (const f of candidates) {
    const xml = await f.async("string");
    // The DataMashup tag may have attributes (xmlns, version=…); match liberally.
    const m = /<DataMashup\b[^>]*>\s*([A-Za-z0-9+/=\s]+?)\s*<\/DataMashup>/i.exec(xml);
    if (m) return m[1].replace(/\s+/g, "");
  }
  return null;
}

/**
 * Parse the QDEFF wire format and return the inner package (a ZIP archive
 * containing Formulas/Section1.m, etc.).
 *
 * Layout:
 *   uint32 LE: version
 *   uint32 LE: package size       --+
 *   ...      : package bytes        | <- this is the inner ZIP we want
 *   uint32 LE: permissions size   --+
 *   ...      : permissions XML
 *   uint32 LE: metadata size
 *   ...      : metadata bytes (XML + binary)
 *   uint32 LE: permissionsBindings size
 *   ...      : permissionsBindings XML
 */
function parseQdeffPackage(blob: Uint8Array): Uint8Array {
  if (blob.byteLength < 8) throw new Error("DataMashup blob too short");
  const dv = new DataView(blob.buffer, blob.byteOffset, blob.byteLength);
  const version = dv.getUint32(0, true);
  const packageSize = dv.getUint32(4, true);
  if (version !== 0) {
    // Don't throw — newer versions still seem to put the package right after.
    // Just emit a warning; the typical version field is 0.
    console.warn(`DataMashup version ${version} (expected 0); attempting parse.`);
  }
  if (8 + packageSize > blob.byteLength) {
    throw new Error(
      `DataMashup truncated: declared package size ${packageSize} exceeds blob length`
    );
  }
  return blob.subarray(8, 8 + packageSize);
}

/** Parse `shared X = let ...` declarations out of an M document. */
export function parseQueryNames(mDocument: string): string[] {
  const names = new Set<string>();
  // Names may be quoted (`#"My Query"`) or bare identifiers.
  const re = /^[ \t]*shared[ \t]+(?:#"([^"]+)"|([A-Za-z_][A-Za-z0-9_]*))[ \t]*=/gm;
  let m: RegExpExecArray | null;
  while ((m = re.exec(mDocument)) !== null) {
    names.add(m[1] ?? m[2]);
  }
  return [...names];
}

/* -------------------------------------------------------------------------- */
/* Helpers                                                                     */
/* -------------------------------------------------------------------------- */

function makeMinimalMashupMetadata(queries: string[], locale: string): MashupMetadata {
  const QueriesMetadata: Record<string, QueryMetadata> = {};
  for (const name of queries) {
    QueriesMetadata[name] = {
      QueryId: cryptoUuid(),
      QueryName: name,
      QueryGroupId: null,
      LastKnownIsCalculatedEntity: false,
      LastKnownIsLinkedEntity: false,
      LastKnownIsParameter: false,
      IsHidden: null,
      LastKnownResultTypeName: null,
      LoadEnabled: false,
      DeleteExistingDataOnLoad: false,
      FieldsMetadata: {},
      HostProperties: {},
      JsonOutputDestinations: null,
      BindToDefaultOutputDestination: null,
      JsonIncrementalRefreshSettings: null,
      JsonStagingDefinition: null,
    };
  }
  return {
    QueryGroups: [],
    DocumentLocale: locale,
    QueriesMetadata,
    FastCombine: false,
    AllowNativeQueries: false,
  };
}

function dataverseKindToDataflowType(kind: DataverseFieldKind): DataflowFieldType {
  switch (kind) {
    case "string":
      return "String";
    case "memo":
      return "Memo";
    case "integer":
    case "status":
    case "state":
      return "Integer";
    case "decimal":
      return "Decimal";
    case "double":
      return "Double";
    case "money":
      return "Money";
    case "boolean":
      return "Boolean";
    case "datetime":
      return "DateAndTime";
    case "dateonly":
      return "DateOnly";
    case "uniqueidentifier":
      return "UniqueIdentifier";
    case "lookup":
      return "Lookup";
    case "choice":
    case "multichoice":
      return "Choice";
    default:
      return "String";
  }
}

function dataflowTypeToDataverseKind(t: string): DataverseFieldKind {
  switch (t) {
    case "Memo":
      return "memo";
    case "String":
      return "string";
    case "Integer":
      return "integer";
    case "Decimal":
      return "decimal";
    case "Double":
      return "double";
    case "Money":
      return "money";
    case "Boolean":
      return "boolean";
    case "DateAndTime":
      return "datetime";
    case "DateOnly":
      return "dateonly";
    case "UniqueIdentifier":
      return "uniqueidentifier";
    case "Lookup":
      return "lookup";
    case "Choice":
      return "choice";
    default:
      return "string"; // safe fallback
  }
}

function base64Decode(b64: string): Uint8Array {
  // Cross-runtime: use Buffer in Node, atob in the browser.
  if (typeof Buffer !== "undefined") {
    const buf = Buffer.from(b64, "base64");
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  }
  const binary = atob(b64);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

function cryptoUuid(): string {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID();
  }
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}
