// Generates load/edge-case test workbooks for the dvload test protocol.
//
// Usage:
//   node make-load-data.js --rows 1000                 -> contacts_1k.xlsx
//   node make-load-data.js --rows 100000               -> contacts_100k.xlsx
//   node make-load-data.js --rows 1000 --errors 25     -> 25 rows with bad data
//   node make-load-data.js --rows 1000 --mutate 10     -> same rows, 10% changed job titles (delta tests)
//   node make-load-data.js --rows 1000 --csv           -> .csv instead of .xlsx
//   node make-load-data.js --edge                      -> edge-cases.xlsx (type coercion tests)
//
// Deterministic: same --seed (default 42) always produces identical data,
// so a --mutate run differs from the base run ONLY in the mutated cells.
// All rows share Last Name "DVLT-Load" (edge: "DVLT-Edge") for easy bulk
// cleanup in Dataverse afterwards.
const ExcelJS = require("../../node_modules/exceljs");
const path = require("path");
const fs = require("fs");

// --- args ---------------------------------------------------------------
const args = process.argv.slice(2);
function argVal(name, def) {
  const i = args.indexOf("--" + name);
  return i >= 0 ? args[i + 1] : def;
}
const hasFlag = (name) => args.includes("--" + name);

const rows = parseInt(argVal("rows", "1000"), 10);
const seed = parseInt(argVal("seed", "42"), 10);
const errors = parseInt(argVal("errors", "0"), 10);
const mutatePct = parseFloat(argVal("mutate", "0"));
const asCsv = hasFlag("csv");
const edge = hasFlag("edge");

function defaultName() {
  if (edge) return "edge-cases.xlsx";
  const size = rows >= 1000 && rows % 1000 === 0 ? rows / 1000 + "k" : String(rows);
  const suffix =
    (mutatePct > 0 ? "_mutated" : "") + (errors > 0 ? "_err" + errors : "");
  return "contacts_" + size + suffix + (asCsv ? ".csv" : ".xlsx");
}
const out = path.resolve(__dirname, argVal("out", defaultName()));

// --- seeded RNG (mulberry32) --------------------------------------------
function mulberry32(a) {
  return function () {
    a |= 0; a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const FIRST = ["Alice","Bob","Carla","David","Elena","Farid","Greta","Hiro","Ines","Jonas","Katya","Liam","Mona","Nils","Olga","Priya","Quentin","Rosa","Sven","Tara"];
const TITLES = ["Analyst","Engineer","Manager","Consultant","Director","Coordinator","Specialist","Architect","Designer","Technician"];

const COLUMNS = ["First Name","Last Name","Email","Job Title","Birthday","Credit Limit","Do Not Email","Notes"];

function makeRow(i, rand, mutateSet) {
  const first = FIRST[Math.floor(rand() * FIRST.length)];
  let title = TITLES[Math.floor(rand() * TITLES.length)];
  const birthday = new Date(Date.UTC(1950 + Math.floor(rand() * 55), Math.floor(rand() * 12), 1 + Math.floor(rand() * 28)));
  const credit = Math.round(rand() * 100000) / 100;
  const dne = rand() < 0.5 ? "TRUE" : "FALSE";
  if (mutateSet && mutateSet.has(i)) title = title + " (updated)";
  return [
    first,
    "DVLT-Load",
    "user" + i + ".s" + seed + "@dvload-test.invalid",
    title,
    birthday.toISOString().slice(0, 10),
    credit,
    dne,
    "Generated row " + i + " for dvload load testing.",
  ];
}

function edgeRows() {
  const M = (n) => "M".repeat(n);
  return [
    // description of the case is encoded in Notes
    ["Zoë",       "DVLT-Edge", "zoe.edge@dvload-test.invalid",    "Ünïcode Diacritics", "1980-01-01", 100.5,  "TRUE",  "diacritics in strings"],
    ["田中",       "DVLT-Edge", "tanaka.edge@dvload-test.invalid", "CJK 部長",           "1975-06-15", 0,      "FALSE", "CJK characters"],
    ["Emoji🙂",   "DVLT-Edge", "emoji.edge@dvload-test.invalid",  "Tester",             "1990-12-31", 999.99, "TRUE",  "emoji in string"],
    ["Max",       "DVLT-Edge", "maxmemo.edge@dvload-test.invalid","Tester",             "1985-03-03", 1,      "FALSE", M(4000)],
    ["Empty",     "DVLT-Edge", "empty.edge@dvload-test.invalid",  "",                   "",           "",     "",      "empty cells -> treatEmptyAsNull"],
    ["Spaces",    "DVLT-Edge", "  padded.edge@dvload-test.invalid  ", "  padded  ",     "1970-01-01", 50,     "TRUE",  "leading/trailing whitespace"],
    ["NumText",   "DVLT-Edge", "numtext.edge@dvload-test.invalid","Tester",             "2000-02-29", "12345.67", "1", "numbers & booleans given as text"],
    ["Negative",  "DVLT-Edge", "negative.edge@dvload-test.invalid","Tester",            "1960-11-11", -500.25, "No",  "negative money, boolean as No"],
    ["BigInt",    "DVLT-Edge", "bigint.edge@dvload-test.invalid", "Tester",             "1955-05-05", 2147483647, "Yes", "int32 max in numeric column"],
    ["OldDate",   "DVLT-Edge", "olddate.edge@dvload-test.invalid","Tester",             "1899-12-31", 10,     "TRUE",  "date below Dataverse min (expect row error)"],
    ["FutureDate","DVLT-Edge", "future.edge@dvload-test.invalid", "Tester",             "2099-01-01", 10,     "FALSE", "far-future date"],
    ["BadDate",   "DVLT-Edge", "baddate.edge@dvload-test.invalid","Tester",             "not-a-date", 10,     "TRUE",  "invalid date literal (expect row error)"],
    ["BadNumber", "DVLT-Edge", "badnum.edge@dvload-test.invalid", "Tester",             "1980-01-01", "12,34abc", "TRUE", "invalid number literal (expect row error)"],
    ["DupeKey",   "DVLT-Edge", "dupe.edge@dvload-test.invalid",   "Tester",             "1980-01-01", 10,     "TRUE",  "duplicate email #1 (upsert-key collision)"],
    ["DupeKey2",  "DVLT-Edge", "dupe.edge@dvload-test.invalid",   "Tester",             "1980-01-01", 20,     "TRUE",  "duplicate email #2 (upsert-key collision)"],
    ["Quotes",    "DVLT-Edge", "quotes.edge@dvload-test.invalid", "O'Brien \"Quoted\"", "1980-01-01", 10,     "TRUE",  "quotes and apostrophes"],
    ["Newline",   "DVLT-Edge", "newline.edge@dvload-test.invalid","Line1\nLine2",       "1980-01-01", 10,     "TRUE",  "embedded newline in cell"],
  ];
}

// --- build --------------------------------------------------------------
async function main() {
  const rand = mulberry32(seed);
  let data;
  if (edge) {
    data = edgeRows();
  } else {
    const mutateSet = new Set();
    if (mutatePct > 0) {
      const mrand = mulberry32(seed + 1); // independent stream: base data stays identical
      const target = Math.round((rows * mutatePct) / 100);
      while (mutateSet.size < target) mutateSet.add(Math.floor(mrand() * rows));
    }
    data = [];
    for (let i = 0; i < rows; i++) data.push(makeRow(i, rand, mutateSet));
    if (errors > 0) {
      // corrupt evenly spaced rows: bad birthday -> coercion/row error
      const step = Math.max(1, Math.floor(rows / errors));
      let injected = 0;
      for (let i = 0; i < rows && injected < errors; i += step, injected++) {
        data[i][4] = "not-a-date";
        data[i][7] = "INJECTED ERROR ROW " + i;
      }
    }
  }

  if (asCsv) {
    const esc = (v) => {
      const s = String(v ?? "");
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    const lines = [COLUMNS.join(",")].concat(data.map((r) => r.map(esc).join(",")));
    fs.writeFileSync(out, lines.join("\r\n"), "utf8");
  } else {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Contacts");
    ws.addTable({
      name: edge ? "tblEdgeCases" : "tblContacts",
      ref: "A1",
      headerRow: true,
      columns: COLUMNS.map((name) => ({ name })),
      rows: data,
    });
    await wb.xlsx.writeFile(out);
  }
  console.log("Wrote", out, "(" + data.length + " rows)");
}

main().catch((e) => { console.error(e); process.exit(1); });
