const DEFAULT_SHEET_COLS = 15;
const LEGACY_PACKAGE_COLUMN_INDEX = 8;
const {
  FIRST_CONFIRMATION_COLUMN_INDEX,
  TAX_ITEM_CODE_COLUMN_INDEX,
} = require("../config/sheetConstants");

const defaultCellStyle = {
  fontWeight: "normal",
  fontStyle: "normal",
  textDecoration: "none",
  color: "#111827",
  backgroundColor: "#ffffff",
  fontSize: "14px",
  fontFamily: "Arial",
  textAlign: "center",
};

const createCell = (value = "") => ({
  value,
  formula: "",
  style: { ...defaultCellStyle },
});

const normalizeRowCells = (cells = []) => {
  const sourceCells =
    cells.length > DEFAULT_SHEET_COLS
      ? cells.filter((_, index) => index !== LEGACY_PACKAGE_COLUMN_INDEX)
      : cells;

  return Array.from({ length: DEFAULT_SHEET_COLS }, (_, colIndex) => {
    const cell = sourceCells[colIndex];

    if (cell && typeof cell === "object") {
      return {
        value: cell.value ?? "",
        formula: cell.formula || "",
        style: {
          ...defaultCellStyle,
          ...(cell.style || {}),
        },
      };
    }

    return createCell(cell || "");
  });
};

const compactRowCells = (cells = []) => {
  const compactCells = normalizeRowCells(cells).map((cell) => {
    const style = Object.fromEntries(
      Object.entries(cell.style || {}).filter(([key, value]) => value !== defaultCellStyle[key])
    );
    const hasCustomStyle = Object.keys(style).length > 0;
    const formula = cell.formula || "";
    const value = cell.value ?? "";

    if (!formula && !hasCustomStyle) {
      return typeof value === "string" ? value : { value, formula: "" };
    }
    return { value, formula, ...(hasCustomStyle ? { style } : {}) };
  });

  while (compactCells.length > 0 && compactCells[compactCells.length - 1] === "") {
    compactCells.pop();
  }

  return compactCells;
};

const compactStoredRowCells = (cells = []) => {
  const storedCells = normalizeRowCells(cells).map((cell) => ({
    value: cell.value ?? "",
    formula: cell.formula || "",
    style: Object.fromEntries(
      Object.entries(cell.style || {}).filter(([key, value]) => value !== defaultCellStyle[key])
    ),
  }));

  while (storedCells.length > 0) {
    const cell = storedCells[storedCells.length - 1];
    if (cell.value !== "" || cell.formula || Object.keys(cell.style).length > 0) break;
    storedCells.pop();
  }

  return storedCells;
};

const buildRowSearchText = (cells = []) => {
  return normalizeRowCells(cells)
    .map((cell) => `${cell.value ?? ""} ${cell.formula || ""}`)
    .join(" ")
    .trim()
    .toLowerCase();
};

const rowHasStoredData = (cells = []) => normalizeRowCells(cells).some((cell) => {
  if (String(cell.value ?? "").trim() || String(cell.formula || "").trim()) return true;

  const style = cell.style || {};
  return Object.entries(style).some(([key, value]) => (
    value !== undefined && value !== null && value !== "" && value !== defaultCellStyle[key]
  ));
});

const getCellText = (cells = [], colIndex) => {
  const cell = cells[colIndex] || {};
  if (cell && typeof cell === "object") return String(cell.formula || cell.value || "").trim();
  return String(cell || "").trim();
};

const rowNeedsCode = (cells = []) => {
  const normalizedCells = normalizeRowCells(cells);
  const firstConfirmation = getCellText(normalizedCells, FIRST_CONFIRMATION_COLUMN_INDEX);
  const itemCode = getCellText(normalizedCells, TAX_ITEM_CODE_COLUMN_INDEX);

  return Boolean(firstConfirmation && !itemCode);
};

const normalizeSearchText = (input = "") => {
  if (Array.isArray(input)) return buildRowSearchText(input);

  return String(input || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
};

const buildTextGrams = (text, minSize = 2, maxSize = 3) => {
  const normalized = normalizeSearchText(text).replace(/\s+/g, "");
  const grams = new Set();

  for (let size = minSize; size <= maxSize; size += 1) {
    if (normalized.length < size) continue;

    for (let index = 0; index <= normalized.length - size; index += 1) {
      grams.add(normalized.slice(index, index + size));
    }
  }

  return Array.from(grams);
};

const buildRowSearchTokens = (cells = []) => {
  const text = normalizeSearchText(cells);
  if (!text) return [];

  const tokens = new Set();

  text.split(/\s+/).forEach((word) => {
    if (!word) return;
    tokens.add(word);
    buildTextGrams(word).forEach((gram) => tokens.add(gram));
  });

  return Array.from(tokens).slice(0, 500);
};

const buildSearchQueryTokens = (query = "") => {
  const text = normalizeSearchText(query);
  if (!text) return [];

  return Array.from(new Set(
    text.split(/\s+/).flatMap((word) => {
      if (word.length <= 3) return [word];
      return buildTextGrams(word);
    })
  ));
};

module.exports = {
  DEFAULT_SHEET_COLS,
  defaultCellStyle,
  createCell,
  normalizeRowCells,
  compactRowCells,
  compactStoredRowCells,
  buildRowSearchText,
  buildRowSearchTokens,
  buildSearchQueryTokens,
  rowHasStoredData,
  rowNeedsCode,
};
