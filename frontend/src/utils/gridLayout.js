export const EXCEL_IMPORT_OPTIONS = Object.freeze({
  header: 1,
  defval: "",
  raw: false,
});

export const getSheetTableWidth = (columnWidths = [], rowHeaderWidth = 42) => (
  rowHeaderWidth + columnWidths.reduce((total, width) => total + Math.max(0, Number(width) || 0), 0)
);

export const getCellClipboardText = (cell = {}) => (
  String((cell?.formula || cell?.value) ?? "")
);
