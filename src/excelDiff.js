import * as XLSX from "xlsx";

export function cellLabel(row, col) {
  return XLSX.utils.encode_cell({ r: row, c: col });
}

export function displayValue(value) {
  if (value === null || value === undefined) return "<blank>";
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

export function normalizeValue(value, options) {
  let next = value;
  if (options.trimText && typeof next === "string") next = next.trim();
  if (!options.strictBlank && next === "") next = null;
  if (next === undefined) next = null;
  return next;
}

export function valueForCompare(cell, options) {
  if (!cell) return null;
  if (options.compareFormulas && cell.f) return `=${cell.f}`;
  return cell.v ?? null;
}

export function parseRange(sheet) {
  if (!sheet || !sheet["!ref"]) {
    return { s: { r: 0, c: 0 }, e: { r: -1, c: -1 } };
  }
  return XLSX.utils.decode_range(sheet["!ref"]);
}

export function compareWorkbooks(leftWb, rightWb, options) {
  const names = Array.from(
    new Set([...leftWb.SheetNames, ...rightWb.SheetNames]),
  ).sort((a, b) => a.localeCompare(b));

  const diffs = [];

  for (const sheetName of names) {
    const leftSheet = leftWb.Sheets[sheetName];
    const rightSheet = rightWb.Sheets[sheetName];
    const leftRange = parseRange(leftSheet);
    const rightRange = parseRange(rightSheet);

    const maxRow = Math.max(leftRange.e.r, rightRange.e.r);
    const maxCol = Math.max(leftRange.e.c, rightRange.e.c);

    for (let row = 0; row <= maxRow; row += 1) {
      for (let col = 0; col <= maxCol; col += 1) {
        const address = cellLabel(row, col);
        const leftValue = normalizeValue(
          valueForCompare(leftSheet?.[address], options),
          options,
        );
        const rightValue = normalizeValue(
          valueForCompare(rightSheet?.[address], options),
          options,
        );

        if (leftValue !== rightValue) {
          diffs.push({
            sheet: sheetName,
            cell: address,
            address: `${sheetName}!${address}`,
            leftValue,
            rightValue,
          });
        }
      }
    }
  }

  return diffs;
}
