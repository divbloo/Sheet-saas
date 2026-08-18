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

const MAX_FORMULA_RANGE_CELLS = 10000;

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

const isValidFormulaAddress = (address, data) => {
  return (
    address &&
    address.row >= 0 &&
    address.col >= 0 &&
    address.col < COLS &&
    address.row < Math.max(1, data?.length || 0)
  );
};

const getCellNumber = (data, row, col) => {
  const cell = normalizeCell(data?.[row]?.[col]);
  const number = Number(cell.value);
  return Number.isFinite(number) ? number : 0;
};

const createNumberSummary = () => ({
  count: 0,
  sum: 0,
  min: null,
  max: null,
});

const addNumberToSummary = (summary, value) => {
  const number = Number(value);
  const safeNumber = Number.isFinite(number) ? number : 0;

  summary.count += 1;
  summary.sum += safeNumber;
  summary.min = summary.min === null ? safeNumber : Math.min(summary.min, safeNumber);
  summary.max = summary.max === null ? safeNumber : Math.max(summary.max, safeNumber);
};

export const evaluateFormula = (formula, data) => {
  const expression = String(formula || "").trim();
  if (!expression.startsWith("=")) return formula;

  const body = expression.slice(1).toUpperCase();

  const summarizeRange = (rangeText) => {
    const [start, end] = rangeText.split(":");
    const a = parseAddress(start);
    const b = parseAddress(end);
    if (!a || !b) return null;

    const rowStart = Math.min(a.row, b.row);
    const rowEnd = Math.min(Math.max(a.row, b.row), Math.max(0, (data?.length || 1) - 1));
    const colStart = Math.min(a.col, b.col);
    const colEnd = Math.min(Math.max(a.col, b.col), COLS - 1);
    const cellCount = (rowEnd - rowStart + 1) * (colEnd - colStart + 1);

    if (
      rowStart < 0 ||
      colStart < 0 ||
      rowStart > rowEnd ||
      colStart > colEnd ||
      cellCount > MAX_FORMULA_RANGE_CELLS
    ) {
      return null;
    }

    const summary = createNumberSummary();

    for (let row = rowStart; row <= rowEnd; row += 1) {
      for (let col = colStart; col <= colEnd; col += 1) {
        addNumberToSummary(summary, getCellNumber(data, row, col));
      }
    }

    return summary;
  };

  const fnMatch = body.match(/^(SUM|AVERAGE|MIN|MAX|COUNT)\(([^)]+)\)$/);

  if (fnMatch) {
    const fn = fnMatch[1];
    const arg = fnMatch[2];
    const summary = arg.includes(":")
      ? summarizeRange(arg)
      : arg.split(",").reduce((result, value) => {
          const addr = parseAddress(value.trim());
          const nextValue = isValidFormulaAddress(addr, data)
            ? getCellNumber(data, addr.row, addr.col)
            : Number(value) || 0;

          addNumberToSummary(result, nextValue);
          return result;
        }, createNumberSummary());

    if (!summary) return "#ERROR";
    if (fn === "SUM") return String(summary.sum);
    if (fn === "AVERAGE") return String(summary.count ? summary.sum / summary.count : 0);
    if (fn === "MIN") return String(summary.min ?? 0);
    if (fn === "MAX") return String(summary.max ?? 0);
    if (fn === "COUNT") return String(summary.count);
  }

  try {
    const safeExpression = body.replace(/[A-Z]+\d+/g, (addrText) => {
      const addr = parseAddress(addrText);
      return isValidFormulaAddress(addr, data) ? getCellNumber(data, addr.row, addr.col) : 0;
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
