// make-test.js
const ExcelJS = require("C:\\Users\\danij\\Documents\\Repos\\dataverse-load\\node_modules\\exceljs");
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet("Sheet1");
ws.addTable({
  name: "tblTest",
  ref: "A1",
  headerRow: true,
  columns: [{ name: "Email" }, { name: "FirstName" }, { name: "LastName" }],
  rows: [
    ["test1@example.com", "Test", "One"],
    ["test2@example.com", "Test", "Two"],
    ["test3@example.com", "Test", "Three"],
  ],
});
wb.xlsx.writeFile("C:\\Users\\danij\\test.xlsx").then(() => console.log("wrote test.xlsx"));