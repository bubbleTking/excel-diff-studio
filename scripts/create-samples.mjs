import fs from "node:fs";
import path from "node:path";
import * as XLSX from "xlsx";

const outputDir = path.resolve("samples");
fs.mkdirSync(outputDir, { recursive: true });

function buildWorkbook(version) {
  const workbook = XLSX.utils.book_new();

  const orders =
    version === 1
      ? [
          ["Order ID", "Customer", "Region", "Amount", "Status", "Note"],
          ["A1001", "Bluebird Co.", "Hong Kong", 1250, "Paid", "priority"],
          ["A1002", "Northwind", "Singapore", 980, "Pending", "standard"],
          ["A1003", "Luna Labs", "Tokyo", 1760, "Paid", "ship Friday"],
          ["A1004", "Acme Ltd.", "Taipei", 640, "Review", ""],
        ]
      : [
          ["Order ID", "Customer", "Region", "Amount", "Status", "Note"],
          ["A1001", "Bluebird Co.", "Hong Kong", 1250, "Paid", "priority"],
          ["A1002", "Northwind", "Singapore", 1040, "Paid", "standard"],
          ["A1003", "Luna Labs", "Tokyo", 1760, "Paid", "ship Monday"],
          ["A1005", "New Harbor", "Seoul", 720, "Pending", "new row"],
        ];

  const summary =
    version === 1
      ? [
          ["Metric", "Value"],
          ["Total Orders", 4],
          ["Total Amount", { f: "SUM(Orders!D2:D5)", v: 4630 }],
          ["Paid Orders", 2],
        ]
      : [
          ["Metric", "Value"],
          ["Total Orders", 4],
          ["Total Amount", { f: "SUM(Orders!D2:D5)", v: 4770 }],
          ["Paid Orders", 3],
        ];

  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(orders),
    "Orders",
  );
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(summary),
    "Summary",
  );

  return workbook;
}

XLSX.writeFile(buildWorkbook(1), path.join(outputDir, "sample_file_1.xlsx"));
XLSX.writeFile(buildWorkbook(2), path.join(outputDir, "sample_file_2.xlsx"));

console.log("Created samples/sample_file_1.xlsx");
console.log("Created samples/sample_file_2.xlsx");
