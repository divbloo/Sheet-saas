const DEFAULT_SHEET_COLS = 15;
const LEGACY_PACKAGE_COLUMN_INDEX = 8;

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

const buildRowSearchText = (cells = []) => {
  return normalizeRowCells(cells)
    .map((cell) => `${cell.value ?? ""} ${cell.formula || ""}`)
    .join(" ")
    .trim()
    .toLowerCase();
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
  buildRowSearchText,
  buildRowSearchTokens,
  buildSearchQueryTokens,
};
