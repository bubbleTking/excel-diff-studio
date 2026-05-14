import assert from "node:assert/strict";
import * as XLSX from "xlsx";
import { compareWorkbooks } from "./src/excelDiff.js";

function workbookFromRows(sheetName, rows) {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
  return workbook;
}

const left = workbookFromRows("Sheet1", [
  ["Name", "Amount", "Note"],
  ["Alpha", 100, " same "],
  ["Beta", 200, ""],
]);

const right = workbookFromRows("Sheet1", [
  ["Name", "Amount", "Note"],
  ["Alpha", 150, "same"],
  ["Beta", 200, null],
]);

const diffs = compareWorkbooks(left, right, {
  trimText: true,
  strictBlank: false,
  compareFormulas: true,
});

assert.equal(diffs.length, 1);
assert.equal(diffs[0].address, "Sheet1!B2");
assert.equal(diffs[0].leftValue, 100);
assert.equal(diffs[0].rightValue, 150);

const strictDiffs = compareWorkbooks(left, right, {
  trimText: false,
  strictBlank: true,
  compareFormulas: true,
});

assert.deepEqual(
  strictDiffs.map((diff) => diff.address),
  ["Sheet1!B2", "Sheet1!C2", "Sheet1!C3"],
);

console.log("Smoke test passed.");
