import * as XLSX from "xlsx";
import { compareWorkbooks as diffWorkbooks, displayValue } from "./excelDiff.js";
import "./styles.css";

const state = {
  left: null,
  right: null,
  diffs: [],
  sheetFilter: "all",
  search: "",
  trimText: true,
  strictBlank: false,
  compareFormulas: true,
  lang: "zh",
  busy: false,
  error: "",
};

const app = document.querySelector("#app");

const copy = {
  zh: {
    privacy: "所有比较都在你的浏览器本地完成，文件不会上传。",
    statusReady: "Ready",
    statusWaiting: "Waiting for files",
    file1: "文件 1",
    file2: "文件 2",
    dropHint: "拖拽 Excel 到这里，或点击选择文件",
    chooseFile: "选择文件",
    replaceFile: "更换文件",
    differentCells: "不同单元格",
    affectedSheets: "涉及工作表",
    currentlyShown: "当前显示",
    all: "全部",
    searchPlaceholder: "搜索 sheet / 单元格 / 内容",
    trimSpaces: "忽略首尾空格",
    compareFormulas: "比较公式",
    strictBlank: "严格空值",
    cell: "单元格",
    file1Value: "文件 1 内容",
    file2Value: "文件 2 内容",
    address: "地址",
    emptyBeforeUpload: "先导入两个 Excel 文件，差异会自动出现在这里。",
    emptyNoDiffs: "没有发现不同的单元格。",
    emptyFiltered: "当前筛选条件下没有结果。",
    manyRows:
      "结果很多，当前先显示前 1000 条；可以用搜索或 Sheet 标签缩小范围。",
    readFailed: "读取失败",
    switchLanguage: "English",
    uploadAria: "上传 Excel 文件",
    summaryAria: "比较摘要",
  },
  en: {
    privacy: "All comparison happens locally in your browser. Files are never uploaded.",
    statusReady: "Ready",
    statusWaiting: "Waiting for files",
    file1: "File 1",
    file2: "File 2",
    dropHint: "Drop an Excel file here, or click to choose one",
    chooseFile: "Choose file",
    replaceFile: "Replace file",
    differentCells: "Different cells",
    affectedSheets: "Affected sheets",
    currentlyShown: "Currently shown",
    all: "All",
    searchPlaceholder: "Search sheet / cell / content",
    trimSpaces: "Ignore leading/trailing spaces",
    compareFormulas: "Compare formulas",
    strictBlank: "Strict blanks",
    cell: "Cell",
    file1Value: "File 1 value",
    file2Value: "File 2 value",
    address: "Address",
    emptyBeforeUpload: "Import two Excel files and differences will appear here.",
    emptyNoDiffs: "No different cells found.",
    emptyFiltered: "No results match the current filters.",
    manyRows:
      "There are many results, so only the first 1,000 are shown. Use search or sheet tabs to narrow them down.",
    readFailed: "Read failed",
    switchLanguage: "中文",
    uploadAria: "Upload Excel files",
    summaryAria: "Comparison summary",
  },
};

function t(key) {
  return copy[state.lang][key];
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    cellFormula: true,
  });
  return {
    name: file.name,
    size: file.size,
    workbook,
  };
}

function compareWorkbooks() {
  if (!state.left || !state.right) {
    state.diffs = [];
    return;
  }

  state.diffs = diffWorkbooks(state.left.workbook, state.right.workbook, {
    trimText: state.trimText,
    strictBlank: state.strictBlank,
    compareFormulas: state.compareFormulas,
  });
  if (state.sheetFilter !== "all" && !sheetCounts()[state.sheetFilter]) {
    state.sheetFilter = "all";
  }
}

function sheetCounts() {
  return state.diffs.reduce((acc, diff) => {
    acc[diff.sheet] = (acc[diff.sheet] || 0) + 1;
    return acc;
  }, {});
}

function filteredDiffs() {
  const query = state.search.trim().toLowerCase();
  return state.diffs.filter((diff) => {
    const sheetMatch =
      state.sheetFilter === "all" || diff.sheet === state.sheetFilter;
    if (!sheetMatch) return false;
    if (!query) return true;

    return [
      diff.sheet,
      diff.cell,
      diff.address,
      displayValue(diff.leftValue),
      displayValue(diff.rightValue),
    ]
      .join(" ")
      .toLowerCase()
      .includes(query);
  });
}

function formatBytes(bytes) {
  if (!bytes && bytes !== 0) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function uploadZone(side, label) {
  const file = state[side];
  const ready = Boolean(file);
  return `
    <label class="drop-zone ${ready ? "is-ready" : ""}" data-side="${side}">
      <input class="file-input" data-side="${side}" type="file" accept=".xlsx,.xlsm,.xls,.csv" />
      <span class="zone-icon">${ready ? "✓" : "+"}</span>
      <span class="zone-copy">
        <span class="zone-title">${label}</span>
        <span class="zone-text">${
          ready
            ? `${escapeHtml(file.name)} · ${formatBytes(file.size)}`
            : t("dropHint")
        }</span>
      </span>
      <span class="zone-action">${ready ? t("replaceFile") : t("chooseFile")}</span>
    </label>
  `;
}

function tabsMarkup() {
  const counts = sheetCounts();
  const sheets = Object.keys(counts).sort((a, b) => a.localeCompare(b));
  const allSelected = state.sheetFilter === "all" ? "is-selected" : "";
  return `
    <button class="tab ${allSelected}" data-sheet="all">${t("all")} <span>${state.diffs.length}</span></button>
    ${sheets
      .map((sheet) => {
        const selected = state.sheetFilter === sheet ? "is-selected" : "";
        return `<button class="tab ${selected}" data-sheet="${escapeHtml(sheet)}">${escapeHtml(sheet)} <span>${counts[sheet]}</span></button>`;
      })
      .join("")}
  `;
}

function rowsMarkup(rows) {
  if (!state.left || !state.right) {
    return `
      <tr>
        <td class="empty-row" colspan="4">${t("emptyBeforeUpload")}</td>
      </tr>
    `;
  }

  if (state.diffs.length === 0) {
    return `
      <tr>
        <td class="empty-row" colspan="4">${t("emptyNoDiffs")}</td>
      </tr>
    `;
  }

  if (rows.length === 0) {
    return `
      <tr>
        <td class="empty-row" colspan="4">${t("emptyFiltered")}</td>
      </tr>
    `;
  }

  return rows
    .slice(0, 1000)
    .map(
      (diff) => `
        <tr>
          <td>
            <strong>${escapeHtml(diff.cell)}</strong>
            <span>${escapeHtml(diff.sheet)}</span>
          </td>
          <td class="old-value">${escapeHtml(displayValue(diff.leftValue))}</td>
          <td class="new-value">${escapeHtml(displayValue(diff.rightValue))}</td>
          <td class="address">${escapeHtml(diff.address)}</td>
        </tr>
      `,
    )
    .join("");
}

function render() {
  const rows = filteredDiffs();
  const hasBoth = Boolean(state.left && state.right);
  app.innerHTML = `
    <div class="app-shell">
      <header class="topbar">
        <div>
          <h1>Excel Diff Studio</h1>
          <p>${t("privacy")}</p>
        </div>
        <div class="topbar-actions">
          <button class="language-toggle" type="button">${t("switchLanguage")}</button>
          <div class="status-pill ${hasBoth ? "is-live" : ""}">
            ${hasBoth ? t("statusReady") : t("statusWaiting")}
          </div>
        </div>
      </header>

      <main>
        <section class="upload-grid" aria-label="${t("uploadAria")}">
          ${uploadZone("left", t("file1"))}
          ${uploadZone("right", t("file2"))}
        </section>

        ${state.error ? `<div class="error-banner">${escapeHtml(state.error)}</div>` : ""}
        ${state.busy ? `<div class="loading-bar"><span></span></div>` : ""}

        <section class="summary-row" aria-label="${t("summaryAria")}">
          <div class="metric">
            <strong>${state.diffs.length}</strong>
            <span>${t("differentCells")}</span>
          </div>
          <div class="metric">
            <strong>${Object.keys(sheetCounts()).length}</strong>
            <span>${t("affectedSheets")}</span>
          </div>
          <div class="metric">
            <strong>${rows.length}</strong>
            <span>${t("currentlyShown")}</span>
          </div>
        </section>

        <section class="results-panel">
          <div class="results-toolbar">
            <div class="tabs" role="tablist">
              ${tabsMarkup()}
            </div>
            <div class="controls">
              <input class="search" type="search" placeholder="${t("searchPlaceholder")}" value="${escapeHtml(state.search)}" />
              <label class="check">
                <input type="checkbox" data-option="trimText" ${state.trimText ? "checked" : ""} />
                <span>${t("trimSpaces")}</span>
              </label>
              <label class="check">
                <input type="checkbox" data-option="compareFormulas" ${state.compareFormulas ? "checked" : ""} />
                <span>${t("compareFormulas")}</span>
              </label>
              <label class="check">
                <input type="checkbox" data-option="strictBlank" ${state.strictBlank ? "checked" : ""} />
                <span>${t("strictBlank")}</span>
              </label>
            </div>
          </div>

          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>${t("cell")}</th>
                  <th>${t("file1Value")}</th>
                  <th>${t("file2Value")}</th>
                  <th>${t("address")}</th>
                </tr>
              </thead>
              <tbody>${rowsMarkup(rows)}</tbody>
            </table>
          </div>
          ${
            rows.length > 1000
              ? `<div class="table-note">${t("manyRows")}</div>`
              : ""
          }
        </section>
      </main>
    </div>
  `;

  bindEvents();
}

function bindEvents() {
  document.querySelectorAll(".file-input").forEach((input) => {
    input.addEventListener("change", async (event) => {
      const file = event.target.files?.[0];
      if (file) await handleFile(event.target.dataset.side, file);
    });
  });

  document.querySelector(".language-toggle")?.addEventListener("click", () => {
    state.lang = state.lang === "zh" ? "en" : "zh";
    render();
  });

  document.querySelectorAll(".drop-zone").forEach((zone) => {
    zone.addEventListener("dragover", (event) => {
      event.preventDefault();
      zone.classList.add("is-dragging");
    });
    zone.addEventListener("dragleave", () => zone.classList.remove("is-dragging"));
    zone.addEventListener("drop", async (event) => {
      event.preventDefault();
      zone.classList.remove("is-dragging");
      const file = event.dataTransfer.files?.[0];
      if (file) await handleFile(zone.dataset.side, file);
    });
  });

  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      state.sheetFilter = button.dataset.sheet;
      render();
    });
  });

  document.querySelector(".search")?.addEventListener("input", (event) => {
    state.search = event.target.value;
    render();
  });

  document.querySelectorAll("[data-option]").forEach((input) => {
    input.addEventListener("change", () => {
      state[input.dataset.option] = input.checked;
      compareWorkbooks();
      render();
    });
  });
}

async function handleFile(side, file) {
  state.busy = true;
  state.error = "";
  render();

  try {
    state[side] = await readWorkbook(file);
    compareWorkbooks();
  } catch (error) {
    state.error = `${t("readFailed")}: ${error.message}`;
  } finally {
    state.busy = false;
    render();
  }
}

render();
