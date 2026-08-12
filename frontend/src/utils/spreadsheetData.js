import defaultErpOptions from "../defaultErpOptions.json";
import {
  AVERAGE_CHAR_WIDTH,
  COLS,
  COLUMN_WIDTH_PADDING,
  LEGACY_PACKAGE_COLUMN_INDEX,
  MIN_SHEET_ROWS,
  defaultCellStyle,
  defaultMeta,
  erpArabicHeaders,
  visibleErpHeaders,
} from "../spreadsheetConfig";

export const createEmptyCell = () => ({
  value: "",
  formula: "",
  style: { ...defaultCellStyle },
});

export const createEmptyRow = () => Array.from({ length: COLS }, createEmptyCell);

export const normalizeCell = (cell) => {
  if (typeof cell === "object" && cell !== null && "value" in cell) {
    const incomingStyle = cell.style || {};

    return {
      value: cell.value || "",
      formula: cell.formula || "",
      style: {
        ...defaultCellStyle,
        ...incomingStyle,
        textAlign:
          !incomingStyle.textAlign || incomingStyle.textAlign === "left"
            ? defaultCellStyle.textAlign
            : incomingStyle.textAlign,
      },
    };
  }

  return {
    value: cell || "",
    formula: "",
    style: { ...defaultCellStyle },
  };
};

export const normalizeData = (data = [], minRows = MIN_SHEET_ROWS) => {
  return Array.from({ length: Math.max(minRows, data.length) }, (_, rowIndex) => {
    const row = data[rowIndex] || [];
    const sourceRow =
      row.length > COLS
        ? row.filter((_, index) => index !== LEGACY_PACKAGE_COLUMN_INDEX)
        : row;

    return Array.from({ length: COLS }, (_, colIndex) =>
      normalizeCell(sourceRow[colIndex] || "")
    );
  });
};

export const normalizeSheet = (sheet, options = {}) => ({
  ...sheet,
  data: normalizeData(sheet.data, options.minRows ?? MIN_SHEET_ROWS),
  rowOwners: { ...(sheet.rowOwners || {}) },
  meta: { ...defaultMeta, ...(sheet.meta || {}) },
  erpOptions: { ...defaultErpOptions, ...(sheet.erpOptions || {}) },
});

export const excelColName = (index) => {
  let name = "";
  let n = index + 1;

  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }

  return name;
};

export const colName = (index) => {
  return visibleErpHeaders[index] || erpArabicHeaders[index] || `Column ${index + 1}`;
};

export const cellAddress = (row, col) => excelColName(col) + (row + 1);

const parseAddress = (address) => {
  const match = String(address).toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;

  const letters = match[1];
  const row = Number(match[2]) - 1;

  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }

  return { row, col: col - 1 };
};

const getCellNumber = (data, row, col) => {
  const cell = normalizeCell(data?.[row]?.[col]);
  const number = Number(cell.value);
  return Number.isFinite(number) ? number : 0;
};

export const evaluateFormula = (formula, data) => {
  const expression = String(formula || "").trim();
  if (!expression.startsWith("=")) return formula;

  const body = expression.slice(1).toUpperCase();

  const rangeValues = (rangeText) => {
    const [start, end] = rangeText.split(":");
    const a = parseAddress(start);
    const b = parseAddress(end);
    if (!a || !b) return [];

    const values = [];
    const rowStart = Math.min(a.row, b.row);
    const rowEnd = Math.max(a.row, b.row);
    const colStart = Math.min(a.col, b.col);
    const colEnd = Math.max(a.col, b.col);

    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let col = colStart; col <= colEnd; col += 1) {
        values.push(getCellNumber(data, row, col));
      }
    }

    return values;
  };

  const fnMatch = body.match(/^(SUM|AVERAGE|MIN|MAX|COUNT)\(([^)]+)\)$/);

  if (fnMatch) {
    const fn = fnMatch[1];
    const arg = fnMatch[2];
    const values = arg.includes(":")
      ? rangeValues(arg)
      : arg.split(",").map((value) => {
          const addr = parseAddress(value.trim());
          return addr ? getCellNumber(data, addr.row, addr.col) : Number(value) || 0;
        });

    if (fn === "SUM") return String(values.reduce((a, b) => a + b, 0));
    if (fn === "AVERAGE") {
      return String(values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0);
    }
    if (fn === "MIN") return String(Math.min(...values));
    if (fn === "MAX") return String(Math.max(...values));
    if (fn === "COUNT") return String(values.filter((value) => Number.isFinite(value)).length);
  }

  try {
    const safeExpression = body.replace(/[A-Z]+\d+/g, (addrText) => {
      const addr = parseAddress(addrText);
      return addr ? getCellNumber(data, addr.row, addr.col) : 0;
    });

    if (!/^[0-9+\-*/().\s]+$/.test(safeExpression)) return "#ERROR";
    return String(Function("return " + safeExpression)());
  } catch {
    return "#ERROR";
  }
};

export const recalculateData = (data) => {
  return data.map((row) =>
    row.map((cell) => {
      const normalized = normalizeCell(cell);
      if (normalized.formula) {
        return {
          ...normalized,
          value: evaluateFormula(normalized.formula, data),
        };
      }
      return normalized;
    })
  );
};

export const estimateTextWidth = (text) => {
  const longestLine = String(text || "")
    .split(/\r?\n/)
    .reduce((longest, line) => Math.max(longest, line.length), 0);

  return Math.ceil(longestLine * AVERAGE_CHAR_WIDTH + COLUMN_WIDTH_PADDING);
};
