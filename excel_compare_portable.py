#!/usr/bin/env python3
"""
Portable Excel Compare

Compare two .xlsx files locally and generate a visual HTML diff report.

No third-party packages are required. This script uses only Python's standard
library, so it is easy to copy to another computer.

Usage:
  python3 excel_compare_portable.py
  python3 excel_compare_portable.py file1.xlsx file2.xlsx
  python3 excel_compare_portable.py file1.xlsx file2.xlsx -o report.html
  python3 excel_compare_portable.py file1.xlsx file2.xlsx --data-only

Notes:
  - Supports .xlsx files.
  - Does not support legacy .xls files.
  - By default, formulas are compared as formulas. Use --data-only to compare
    cached formula results saved in the workbook.
"""

from __future__ import annotations

import argparse
import html
import os
import posixpath
import re
import sys
import urllib.parse
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


WINDOWS_PATH_RE = re.compile(
    r"^(?:[a-zA-Z]:\\|\\\\|//|\\\\\?\\|\\\\\.\\|[a-zA-Z]:/)"
)


@dataclass(frozen=True)
class WorkbookData:
    sheets: dict[str, dict[str, Any]]


@dataclass(frozen=True)
class Difference:
    sheet: str
    cell: str
    left_value: Any
    right_value: Any


def read_xml(zf: zipfile.ZipFile, path: str) -> ET.Element:
    with zf.open(path) as file:
        return ET.fromstring(file.read())


def text_content(element: ET.Element | None) -> str:
    if element is None:
        return ""
    return "".join(element.itertext())


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    path = "xl/sharedStrings.xml"
    if path not in zf.namelist():
        return []

    root = read_xml(zf, path)
    strings = []
    for item in root.findall("main:si", NS):
        strings.append(text_content(item))
    return strings


def load_workbook_relationships(zf: zipfile.ZipFile) -> dict[str, str]:
    root = read_xml(zf, "xl/_rels/workbook.xml.rels")
    relationships = {}
    for rel in root.findall("pkgrel:Relationship", NS):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if not rel_id or not target:
            continue
        relationships[rel_id] = posixpath.normpath(posixpath.join("xl", target))
    return relationships


def load_sheet_paths(zf: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = read_xml(zf, "xl/workbook.xml")
    relationships = load_workbook_relationships(zf)
    sheet_paths = []

    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        name = sheet.attrib.get("name", "Sheet")
        rel_id = sheet.attrib.get(f"{{{NS['rel']}}}id")
        path = relationships.get(rel_id)
        if path:
            sheet_paths.append((name, path))

    return sheet_paths


def parse_cell_value(cell: ET.Element, shared_strings: list[str], data_only: bool) -> Any:
    formula = cell.find("main:f", NS)
    if formula is not None and not data_only:
        formula_text = text_content(formula)
        return f"={formula_text}" if formula_text else "="

    cell_type = cell.attrib.get("t")

    if cell_type == "inlineStr":
        return text_content(cell.find("main:is", NS))

    value_element = cell.find("main:v", NS)
    value = text_content(value_element)

    if value == "":
        return None

    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError):
            return value

    if cell_type == "b":
        return value == "1"

    if cell_type in {"str", "e"}:
        return value

    try:
        number = float(value)
    except ValueError:
        return value

    if number.is_integer():
        return int(number)
    return number


def load_sheet(
    zf: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: list[str],
    data_only: bool,
) -> dict[str, Any]:
    root = read_xml(zf, sheet_path)
    cells: dict[str, Any] = {}

    for cell in root.findall(".//main:c", NS):
        address = cell.attrib.get("r")
        if not address:
            continue
        cells[address] = parse_cell_value(cell, shared_strings, data_only)

    return cells


def load_workbook(path: Path, data_only: bool) -> WorkbookData:
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"{path} is not an .xlsx file")

    with zipfile.ZipFile(path) as zf:
        shared_strings = load_shared_strings(zf)
        sheets = {}
        for sheet_name, sheet_path in load_sheet_paths(zf):
            sheets[sheet_name] = load_sheet(zf, sheet_path, shared_strings, data_only)

    return WorkbookData(sheets=sheets)


def normalize(value: Any, trim_text: bool, strict_blank: bool) -> Any:
    if trim_text and isinstance(value, str):
        value = value.strip()

    if not strict_blank and value == "":
        return None

    return value


def compare_workbooks(
    left: WorkbookData,
    right: WorkbookData,
    *,
    trim_text: bool,
    strict_blank: bool,
) -> list[Difference]:
    differences = []
    sheet_names = sorted(set(left.sheets) | set(right.sheets))

    for sheet_name in sheet_names:
        left_sheet = left.sheets.get(sheet_name, {})
        right_sheet = right.sheets.get(sheet_name, {})
        cell_names = sorted(set(left_sheet) | set(right_sheet), key=cell_sort_key)

        for cell in cell_names:
            left_value = normalize(left_sheet.get(cell), trim_text, strict_blank)
            right_value = normalize(right_sheet.get(cell), trim_text, strict_blank)

            if left_value != right_value:
                differences.append(Difference(sheet_name, cell, left_value, right_value))

    return differences


def cell_sort_key(cell: str) -> tuple[int, int, str]:
    letters = "".join(ch for ch in cell if ch.isalpha()).upper()
    digits = "".join(ch for ch in cell if ch.isdigit())

    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)

    row = int(digits) if digits else 0
    return row, col, cell


def display_value(value: Any) -> str:
    if value is None:
        return "<blank>"
    return str(value)


def write_report(
    output_path: Path,
    file1: Path,
    file2: Path,
    differences: list[Difference],
) -> None:
    sheet_counts: dict[str, int] = {}
    for diff in differences:
        sheet_counts[diff.sheet] = sheet_counts.get(diff.sheet, 0) + 1

    summary_rows = "\n".join(
        "<tr>"
        f"<td>{html.escape(sheet)}</td>"
        f"<td>{count}</td>"
        "</tr>"
        for sheet, count in sorted(sheet_counts.items())
    )

    diff_rows = "\n".join(
        "<tr>"
        f"<td><strong>{html.escape(diff.cell)}</strong><span>{html.escape(diff.sheet)}</span></td>"
        f"<td class=\"old-value\">{html.escape(display_value(diff.left_value))}</td>"
        f"<td class=\"new-value\">{html.escape(display_value(diff.right_value))}</td>"
        f"<td class=\"address\">{html.escape(diff.sheet)}!{html.escape(diff.cell)}</td>"
        "</tr>"
        for diff in differences
    )

    if not summary_rows:
        summary_rows = "<tr><td colspan=\"2\">No differences found.</td></tr>"

    if not diff_rows:
        diff_rows = "<tr><td class=\"empty\" colspan=\"4\">No differences found.</td></tr>"

    page = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Excel Compare Report</title>
  <style>
    :root {{
      color-scheme: light;
      --bg: #f4f7fb;
      --surface: #ffffff;
      --text: #172033;
      --muted: #64748b;
      --line: #d7e0ea;
      --accent: #0f766e;
      --old-bg: #fff1f2;
      --old-line: #f2a7b3;
      --new-bg: #ecfdf3;
      --new-line: #86d39d;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: var(--bg);
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
      line-height: 1.45;
    }}
    header {{
      padding: 26px 32px;
      background: var(--surface);
      border-bottom: 1px solid var(--line);
    }}
    h1 {{
      margin: 0;
      font-size: 28px;
      letter-spacing: 0;
    }}
    .files {{
      display: grid;
      gap: 4px;
      margin-top: 10px;
      color: var(--muted);
      font-size: 14px;
      word-break: break-all;
    }}
    main {{
      width: min(1180px, calc(100% - 32px));
      margin: 0 auto;
      padding: 24px 0 40px;
    }}
    .stats {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 14px;
      margin-bottom: 18px;
    }}
    .metric {{
      padding: 16px 18px;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
    }}
    .metric strong {{
      display: block;
      color: var(--accent);
      font-size: 30px;
      line-height: 1;
    }}
    .metric span {{
      display: block;
      margin-top: 8px;
      color: var(--muted);
      font-size: 13px;
    }}
    section {{
      overflow: hidden;
      margin-bottom: 18px;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
    }}
    h2 {{
      margin: 0;
      padding: 14px 16px;
      border-bottom: 1px solid var(--line);
      font-size: 16px;
    }}
    .table-wrap {{
      overflow: auto;
      max-height: 68vh;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }}
    th,
    td {{
      padding: 10px 12px;
      border-bottom: 1px solid var(--line);
      text-align: left;
      vertical-align: top;
    }}
    thead th {{
      position: sticky;
      top: 0;
      z-index: 1;
      background: #edf3f9;
      color: #334155;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0;
    }}
    td strong {{
      display: block;
      color: var(--accent);
    }}
    td span {{
      display: block;
      margin-top: 3px;
      color: var(--muted);
      font-size: 12px;
    }}
    .old-value,
    .new-value {{
      min-width: 220px;
      max-width: 420px;
      white-space: pre-wrap;
      overflow-wrap: anywhere;
    }}
    .old-value {{
      background: var(--old-bg);
      border-left: 4px solid var(--old-line);
    }}
    .new-value {{
      background: var(--new-bg);
      border-left: 4px solid var(--new-line);
    }}
    .address {{
      color: var(--muted);
      font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
      font-size: 13px;
    }}
    .empty {{
      height: 120px;
      text-align: center;
      color: var(--muted);
    }}
  </style>
</head>
<body>
  <header>
    <h1>Excel Compare Report</h1>
    <div class="files">
      <div><strong>File 1:</strong> {html.escape(str(file1))}</div>
      <div><strong>File 2:</strong> {html.escape(str(file2))}</div>
    </div>
  </header>
  <main>
    <div class="stats">
      <div class="metric"><strong>{len(differences)}</strong><span>Different cells</span></div>
      <div class="metric"><strong>{len(sheet_counts)}</strong><span>Affected sheets</span></div>
    </div>
    <section>
      <h2>Summary by Sheet</h2>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Sheet</th><th>Different Cells</th></tr></thead>
          <tbody>{summary_rows}</tbody>
        </table>
      </div>
    </section>
    <section>
      <h2>Cell Differences</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Cell</th>
              <th>File 1 Value</th>
              <th>File 2 Value</th>
              <th>Address</th>
            </tr>
          </thead>
          <tbody>{diff_rows}</tbody>
        </table>
      </div>
    </section>
  </main>
</body>
</html>
"""
    output_path.write_text(page, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare two .xlsx files and generate a visual HTML report."
    )
    parser.add_argument("file1", nargs="?", type=Path, help="First .xlsx file")
    parser.add_argument("file2", nargs="?", type=Path, help="Second .xlsx file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="HTML report path. Default: excel_compare_report.html",
    )
    parser.add_argument(
        "--data-only",
        action="store_true",
        help="Compare cached formula results instead of formula text.",
    )
    parser.add_argument(
        "--no-trim",
        action="store_true",
        help="Do not ignore leading/trailing spaces in text.",
    )
    parser.add_argument(
        "--strict-blank",
        action="store_true",
        help="Treat empty strings and blank cells as different.",
    )
    return parser.parse_args()


def is_windows_like_path(value: str) -> bool:
    return bool(WINDOWS_PATH_RE.match(value))


def strip_wrapping_quotes(value: str) -> str:
    while len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        value = value[1:-1].strip()
    return value


def clean_path_input(raw: str) -> Path:
    value = raw.strip()
    if not value:
        return Path("")

    value = value.strip("\ufeff")

    # PowerShell often represents a dragged or copied path as:
    #   & "C:\path\file.xlsx"
    # or:
    #   & "\\server\share\file.xlsx"
    if value.startswith("& "):
        value = value[2:].strip()

    assignment = None
    if not value.lower().startswith("file://"):
        assignment = re.match(r"^(?:path|file|路径)\s*=\s*(.+)$", value, re.IGNORECASE)
    if assignment:
        value = assignment.group(1).strip()

    value = strip_wrapping_quotes(value)

    # Handle Python-ish raw string notation if someone copies r"...".
    if len(value) >= 3 and value[0] in {"r", "R"} and value[1] in {"'", '"'}:
        value = strip_wrapping_quotes(value[1:].strip())

    if value.startswith("file://"):
        value = urllib.parse.unquote(urllib.parse.urlparse(value).path)

    if os.name != "nt" and not is_windows_like_path(value):
        # macOS/Linux terminals often escape dragged paths as "My\\ File.xlsx".
        value = re.sub(r"\\([\\ ()\[\]{}&;'\"$`!#*?<>|])", r"\1", value)

    value = os.path.expandvars(value)
    return Path(value).expanduser()


def ask_for_path(label: str) -> Path:
    while True:
        raw = input(f"{label}: ").strip()
        path = clean_path_input(raw)
        if path.exists():
            return path
        print(f"File not found after parsing: {path}")
        print(f"Original input was: {raw}")
        print("Tip: local path example: C:\\Users\\Name\\Desktop\\file.xlsx")
        print("Tip: shared drive example: \\\\ServerName\\ShareName\\file.xlsx")


def fill_interactive_args(args: argparse.Namespace) -> argparse.Namespace:
    if args.file1 and args.file2:
        return args

    print("Excel Compare Portable")
    print("Paste the file path, or drag the .xlsx file into this terminal.")
    print()

    if not args.file1:
        args.file1 = ask_for_path("File 1 path")
    if not args.file2:
        args.file2 = ask_for_path("File 2 path")

    if args.output is None:
        args.output = Path("excel_compare_report.html")
        print(f"Report will be saved as: {args.output}")

    return args


def main() -> int:
    args = fill_interactive_args(parse_args())
    if args.output is None:
        args.output = Path("excel_compare_report.html")

    if not args.file1.exists():
        print(f"File not found: {args.file1}", file=sys.stderr)
        return 2

    if not args.file2.exists():
        print(f"File not found: {args.file2}", file=sys.stderr)
        return 2

    try:
        left = load_workbook(args.file1, args.data_only)
        right = load_workbook(args.file2, args.data_only)
    except (zipfile.BadZipFile, KeyError, ET.ParseError, ValueError) as error:
        print(f"Failed to read workbook: {error}", file=sys.stderr)
        return 2

    differences = compare_workbooks(
        left,
        right,
        trim_text=not args.no_trim,
        strict_blank=args.strict_blank,
    )
    write_report(args.output, args.file1, args.file2, differences)

    print(f"Found {len(differences)} different cell(s).")
    print(f"Saved report: {args.output}")
    return 1 if differences else 0


if __name__ == "__main__":
    raise SystemExit(main())
