// Creates tests/dummy-data/contacts.xlsx with a tblContacts table.
const ExcelJS = require("../../node_modules/exceljs");
const path = require("path");

const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet("Contacts");
ws.addTable({
  name: "tblContacts",
  ref: "A1",
  headerRow: true,
  columns: [
    { name: "First Name" },
    { name: "Last Name" },
    { name: "Email" },
  ],
  rows: [
    ["Alice", "Testsson", "alice.testsson@dvload-test.invalid"],
    ["Bob",   "Testsson", "bob.testsson@dvload-test.invalid"],
  ],
});

const out = path.join(__dirname, "contacts.xlsx");
wb.xlsx.writeFile(out).then(() => console.log("Wrote", out));
