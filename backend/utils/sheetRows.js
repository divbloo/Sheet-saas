const DEFAULT_SHEET_COLS = 16;

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
  return Array.from({ length: DEFAULT_SHEET_COLS }, (_, colIndex) => {
    const cell = cells[colIndex];

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
    .toLowerCase();
};

module.exports = {
  DEFAULT_SHEET_COLS,
  defaultCellStyle,
  createCell,
  normalizeRowCells,
  buildRowSearchText,
};
