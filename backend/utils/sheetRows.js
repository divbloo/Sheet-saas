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

module.exports = {
  DEFAULT_SHEET_COLS,
  defaultCellStyle,
  createCell,
  normalizeRowCells,
  buildRowSearchText,
};
