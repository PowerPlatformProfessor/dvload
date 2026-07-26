// Public surface of @dvload/core. Both the CLI and the Office.js
// add-in consume this package; nothing here may depend on Node-only APIs
// unless it lives under ./node/.

export * from "./types.js";
export * from "./mapping.js";
export * from "./run-plan.js";
export * from "./coerce.js";
export * from "./dataverse.js";
export * from "./load.js";
export * from "./pqt.js";
export * from "./tablegen.js";

// xlsx-reader uses Node fs by default but ships a buffer-based variant
// so the add-in (which has a Blob, not a path) can use the same engine.
export { readTableFromBuffer, readTableFromFile, writeRowsToFile } from "./xlsx-reader.js";
export { parseCsv, readTableFromCsvString } from "./csv-reader.js";
// (CSV/TSV sources supported since v0.2)
