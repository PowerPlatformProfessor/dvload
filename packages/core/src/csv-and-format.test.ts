import { test } from "vitest";
import assert from "node:assert/strict";
import { parseCsv, readTableFromCsvString } from "./csv-reader.js";
import { coerceValue, parseWithFormat } from "./coerce.js";
import type { ColumnMapping } from "./mapping.js";

test("parseCsv handles quoted fields, embedded delimiters, quotes, and newlines", () => {
  const rows = parseCsv('a,b,c\r\n"1,x","say ""hi""","line1\nline2"\r\n');
  assert.deepEqual(rows, [
    ["a", "b", "c"],
    ["1,x", 'say "hi"', "line1\nline2"],
  ]);
});

test("parseCsv strips a UTF-8 BOM and tolerates a missing final newline", () => {
  const rows = parseCsv("﻿h1,h2\nv1,v2");
  assert.deepEqual(rows, [
    ["h1", "h2"],
    ["v1", "v2"],
  ]);
});

test("readTableFromCsvString maps headers, nulls empties, drops blank rows", () => {
  const { headers, rows } = readTableFromCsvString("name,age\nAda,36\n,\nBob,\n");
  assert.deepEqual(headers, ["name", "age"]);
  assert.deepEqual(rows, [
    { name: "Ada", age: "36" },
    { name: "Bob", age: null },
  ]);
});

test("parseWithFormat parses dd/MM/yyyy and rejects impossible dates", () => {
  assert.deepEqual(parseWithFormat("03/04/2025", "dd/MM/yyyy"), {
    y: 2025,
    mo: 4,
    d: 3,
    h: 0,
    mi: 0,
    s: 0,
  });
  assert.equal(parseWithFormat("30/02/2025", "dd/MM/yyyy"), null);
  assert.equal(parseWithFormat("2025-04-03", "dd/MM/yyyy"), null);
});

const col = (kind: ColumnMapping["kind"], format?: string): ColumnMapping => ({
  source: "s",
  target: "t",
  kind,
  treatEmptyAsNull: true,
  ...(format ? { format } : {}),
});

test("datetime honors an explicit format as UTC", () => {
  assert.equal(
    coerceValue("03/04/2025 13:30", col("datetime", "dd/MM/yyyy HH:mm")),
    "2025-04-03T13:30:00.000Z"
  );
});

test("dateonly takes ISO-shaped strings verbatim (no timezone day-shift)", () => {
  assert.equal(coerceValue("2025-01-01", col("dateonly")), "2025-01-01");
  assert.equal(coerceValue("2025-01-01 00:00", col("dateonly")), "2025-01-01");
});

test("dateonly with explicit format", () => {
  assert.equal(coerceValue("31.12.2025", col("dateonly", "dd.MM.yyyy")), "2025-12-31");
});

test("dateonly from a Date object uses UTC parts", () => {
  assert.equal(coerceValue(new Date(Date.UTC(2025, 0, 1)), col("dateonly")), "2025-01-01");
});
