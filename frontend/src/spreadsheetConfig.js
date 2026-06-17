export const API_URL =
  import.meta.env.VITE_API_URL ||
  (import.meta.env.DEV ? "http://localhost:5000" : window.location.origin);

export const defaultCellStyle = {
  fontWeight: "normal",
  fontStyle: "normal",
  textDecoration: "none",
  color: "#111827",
  backgroundColor: "#ffffff",
  fontSize: "14px",
  fontFamily: "Arial",
  textAlign: "center",
};

export const DEFAULT_ROW_HEIGHT = 36;
export const CELL_HORIZONTAL_PADDING = 16;
export const CELL_VERTICAL_PADDING = 12;
export const DEFAULT_LINE_HEIGHT = 20;
export const COLUMN_WIDTH_PADDING = 28;
export const AVERAGE_CHAR_WIDTH = 7.2;
export const DEFAULT_COLUMN_WIDTHS = {
  0: 300,
  1: 300,
  2: 220,
  3: 220,
  4: 145,
  5: 220,
  6: 145,
  7: 82,
  8: 64,
  9: 64,
  10: 220,
  11: 92,
  12: 150,
  13: 92,
  14: 220,
};
export const AUTO_COLUMN_LIMITS = {
  0: { min: 170, max: 300 },
  1: { min: 170, max: 300 },
  2: { min: 140, max: 240 },
  3: { min: 140, max: 240 },
  4: { min: 110, max: 180 },
  5: { min: 140, max: 240 },
  6: { min: 110, max: 180 },
  7: { min: 64, max: 110 },
  8: { min: 54, max: 76 },
  9: { min: 54, max: 76 },
  10: { min: 110, max: 165 },
  11: { min: 72, max: 110 },
  12: { min: 90, max: 170 },
  13: { min: 72, max: 110 },
  14: { min: 150, max: 260 },
};

export const defaultMeta = {
  colWidths: DEFAULT_COLUMN_WIDTHS,
  rowHeights: {},
  merges: [],
  versions: [],
};

export const rowPalette = [
  {
    backgroundColor: "#e0f2fe",
    color: "#075985",
  },
  {
    backgroundColor: "#ccfbf1",
    color: "#115e59",
  },
];

export const erpArabicHeaders = [
  "اسم الصنف",
  "الوصف",
  "المجموعة الرئيسية",
  "المجموعة الفرعية",
  "المجموعة تحت الفرعية",
  "المجموعة المساعدة",
  "المجموعة التفصيلية",
  "وحدة القياس",
  "الصلاحية",
  "التسلسل",
  "ملاحظات",
  "التأكيد الأول",
  "الكود",
  "التأكيد الثاني",
  "التأكيد الثالث",
];

export const visibleErpHeaders = [
  "اسم الصنف",
  "الوصف",
  "المجموعة الرئيسية",
  "المجموعة الفرعية",
  "المجموعة تحت الفرعية",
  "المجموعة المساعدة",
  "المجموعة التفصيلية",
  "وحدة القياس",
  "الصلاحية",
  "التسلسل",
  "ملاحظات",
  "التأكيد الأول",
  "الكود",
  "التأكيد الثاني",
  "Modified Description",
];

export const COLS = 15;
export const ROW_LOCK_LAST_COLUMN_INDEX = 10;
export const LEGACY_PACKAGE_COLUMN_INDEX = 8;
export const MIN_SHEET_ROWS = 5000;
export const IMPORT_BATCH_SIZE = 500;
export const EXPORT_ROW_BATCH_SIZE = 500;
export const INITIAL_VISIBLE_ROWS = 50;
export const ROW_LOAD_STEP = 50;
export const CONTEXT_MENU_MARGIN = 8;
export const CELL_SAVE_DEBOUNCE_MS = 400;
export const VIRTUAL_ROW_BUFFER = 12;
export const VIRTUAL_ROW_WINDOW = 90;
export const AUTO_FIT_SAMPLE_ROWS = 60;
export const AUTO_REFRESH_INTERVAL_MS = 5 * 60 * 1000;

export const clamp = (value, min, max) => Math.min(max, Math.max(min, value));
export const normalizeRowTotal = (...values) => Math.max(
  MIN_SHEET_ROWS,
  ...values.map((value) => Number(value) || 0)
);
