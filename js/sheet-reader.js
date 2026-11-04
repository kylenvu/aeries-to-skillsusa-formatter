"use strict";
// Shared .xlsx reader. Loaded two ways:
// - in the page, before xlsx.full.min.js (used when workers can't start)
// - as the Web Worker script, where it also loads SheetJS and answers messages
Object.freeze(Object.prototype);

const SHEET_LIMITS = {
  maxSheets: 20,
  maxRows: 20000,
  maxColumns: 200,
  maxCells: 1000000,
};

const READ_OPTIONS = {
  type: "array",
  cellFormula: false,
  cellHTML: false,
  cellStyles: false,
  cellDates: false,
  bookVBA: false,
  bookFiles: false,
  sheetRows: SHEET_LIMITS.maxRows,
};

class SheetReadError extends Error {
  constructor(code) {
    super(code);
    this.name = "SheetReadError";
    this.code = code;
  }
}

// Reads an .xlsx ArrayBuffer and returns [{ name, rows: [[cell text, ...], ...] }].
function readWorkbookRows(data) {
  if (typeof XLSX === "undefined") throw new SheetReadError("missing-library");

  const workbook = XLSX.read(data, READ_OPTIONS);
  const names = workbook.SheetNames || [];
  if (names.length > SHEET_LIMITS.maxSheets) throw new SheetReadError("too-many-sheets");

  let totalCells = 0;
  for (const name of names) {
    const sheet = workbook.Sheets[name];
    if (!sheet || !sheet["!ref"]) continue;
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const columns = range.e.c - range.s.c + 1;
    const rows = range.e.r - range.s.r + 1;
    if (columns > SHEET_LIMITS.maxColumns) throw new SheetReadError("too-large");
    totalCells += columns * rows;
    if (totalCells > SHEET_LIMITS.maxCells) throw new SheetReadError("too-large");
  }

  return names.map((name) => ({
    name: String(name),
    rows: workbook.Sheets[name]
      ? XLSX.utils
          .sheet_to_json(workbook.Sheets[name], { header: 1, raw: false, defval: "", blankrows: false })
          .map((row) => row.map((cell) => String(cell ?? "")))
      : [],
  }));
}

if (typeof WorkerGlobalScope !== "undefined" && self instanceof WorkerGlobalScope) {
  importScripts("xlsx.full.min.js");

  self.postMessage({ type: "ready" });

  self.onmessage = (event) => {
    try {
      self.postMessage({ type: "result", sheets: readWorkbookRows(event.data) });
    } catch (err) {
      self.postMessage({ type: "error", code: err && err.code ? String(err.code) : "read" });
    }
  };
}
