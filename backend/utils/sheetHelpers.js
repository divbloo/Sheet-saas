const zlib = require("zlib");

const defaultErpOptions = require("../config/defaultErpOptions.json");
const {
  CAIRO_TIMEZONE,
  DEFAULT_COLUMN_WIDTHS,
  FIRST_CONFIRMATION_EDIT_END_MINUTE,
  FIRST_CONFIRMATION_EDIT_START_MINUTE,
} = require("../config/sheetConstants");

const escapeCsvCell = (value) => {
  const text = String(value ?? "");
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
};

const escapeRegex = (value) => String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const parseRequiredDate = (value) => {
  const date = new Date(String(value || ""));
  return Number.isNaN(date.getTime()) ? null : date;
};

const compareText = (left, right) => String(left || "").localeCompare(String(right || ""), "en", {
  numeric: true,
  sensitivity: "base",
});

const sendJson = (req, res, payload) => {
  const body = JSON.stringify(payload);

  if (!String(req.headers["accept-encoding"] || "").includes("gzip")) {
    res.type("application/json").send(body);
    return;
  }

  zlib.gzip(body, (error, compressed) => {
    if (error) {
      res.type("application/json").send(body);
      return;
    }

    res.setHeader("Content-Encoding", "gzip");
    res.setHeader("Content-Type", "application/json; charset=utf-8");
    res.send(compressed);
  });
};

const cairoTimeMinutes = () => {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: CAIRO_TIMEZONE,
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(new Date());
  const hour = Number(parts.find((part) => part.type === "hour")?.value || 0) % 24;
  const minute = Number(parts.find((part) => part.type === "minute")?.value || 0);

  return (hour * 60) + minute;
};

const isFirstConfirmationEditOpen = () => {
  const minutes = cairoTimeMinutes();

  return minutes >= FIRST_CONFIRMATION_EDIT_START_MINUTE && minutes <= FIRST_CONFIRMATION_EDIT_END_MINUTE;
};

const createTimeLockedColumnError = () => {
  const error = new Error("First Confirmation can only be edited from 8:00 AM to 3:30 PM Cairo time");
  error.statusCode = 403;
  return error;
};

const createDefaultMeta = () => ({
  colWidths: { ...DEFAULT_COLUMN_WIDTHS },
  rowHeights: {},
  merges: [],
  versions: [],
});

const createDefaultErpOptions = () => JSON.parse(JSON.stringify(defaultErpOptions));

const colName = (index) => {
  let name = "";
  let n = index + 1;

  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }

  return name;
};

const cellAddress = (rowIndex, colIndex) => {
  return colName(colIndex) + String(rowIndex + 1);
};

module.exports = {
  cellAddress,
  compareText,
  createDefaultErpOptions,
  createDefaultMeta,
  createTimeLockedColumnError,
  escapeCsvCell,
  escapeRegex,
  isFirstConfirmationEditOpen,
  parseRequiredDate,
  sendJson,
};
