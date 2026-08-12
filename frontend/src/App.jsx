import { useEffect, useMemo, useRef, useState } from "react";
import { io } from "socket.io-client";
import defaultErpOptions from "./defaultErpOptions.json";
import {
  API_URL,
  AUTO_COLUMN_LIMITS,
  AUTO_FIT_SAMPLE_ROWS,
  AUTO_REFRESH_INTERVAL_MS,
  AVERAGE_CHAR_WIDTH,
  CELL_HORIZONTAL_PADDING,
  CELL_SAVE_DEBOUNCE_MS,
  CELL_VERTICAL_PADDING,
  COLS,
  COLUMN_WIDTH_PADDING,
  CONTEXT_MENU_MARGIN,
  DEFAULT_LINE_HEIGHT,
  DEFAULT_ROW_HEIGHT,
  DEFAULT_COLUMN_WIDTHS,
  EXPORT_ROW_BATCH_SIZE,
  CAIRO_TIMEZONE,
  FIRST_CONFIRMATION_COLUMN_INDEX,
  FIRST_CONFIRMATION_EDIT_END_MINUTE,
  FIRST_CONFIRMATION_EDIT_START_MINUTE,
  IMPORT_BATCH_SIZE,
  INITIAL_VISIBLE_ROWS,
  LEGACY_PACKAGE_COLUMN_INDEX,
  MIN_SHEET_ROWS,
  ROW_LOAD_STEP,
  ROW_LOCK_LAST_COLUMN_INDEX,
  VIRTUAL_ROW_BUFFER,
  VIRTUAL_ROW_WINDOW,
  clamp,
  defaultCellStyle,
  defaultMeta,
  erpArabicHeaders,
  normalizeRowTotal,
  rowPalette,
  visibleErpHeaders,
} from "./spreadsheetConfig";
import "./App.css";

const TAX_EXPORT_HEADERS = [
  "CodeType",
  "ItemCode",
  "CodeName",
  "CodeName2Ar",
  "Description",
  "DescriptionAr",
  "ActiveFrom",
  "ActiveTo",
  "GPCItemLinked",
  "EGSRelatedCode",
];

const TAX_EXPORT_COLUMN_WIDTHS = [14.73, 11.45, 12.54, 15.54, 42.73, 32.45, 30.45, 26.73, 42.18, 36.27];
const TAX_ITEM_NAME_COLUMN_INDEX = 0;
const TAX_ITEM_CODE_COLUMN_INDEX = 12;
const TAX_ACTIVE_FROM_DAYS_AGO = 5;
const TAX_EXPORT_FILE_NAME = "NewCodeBulkTemplate (2).xlsx";
const DESCRIPTION_BUILDER_COLUMN_INDEX = 0;
const EMPTY_DESCRIPTION_OPTIONS = { fields: [], rows: [] };

const padDatePart = (value) => String(value).padStart(2, "0");

const getCairoDateTimeParts = (date = new Date()) => {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: CAIRO_TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(date);

  return Object.fromEntries(parts.map((part) => [part.type, part.value]));
};

const formatCairoDateTimeInput = (date = new Date()) => {
  const parts = getCairoDateTimeParts(date);
  const hour = String(Number(parts.hour || 0) % 24);

  return [
    parts.year,
    padDatePart(parts.month),
    padDatePart(parts.day),
  ].join("-") + "T" + [padDatePart(hour), padDatePart(parts.minute)].join(":");
};

const getDefaultTaxExportPeriod = () => {
  const now = new Date();
  const parts = getCairoDateTimeParts(now);
  const date = [parts.year, padDatePart(parts.month), padDatePart(parts.day)].join("-");

  return {
    from: date + "T00:00",
    to: formatCairoDateTimeInput(now),
  };
};

function App() {
  const [mode, setMode] = useState("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [username, setUsername] = useState("");
  const [authFormKey, setAuthFormKey] = useState(0);
  const [token, setToken] = useState(sessionStorage.getItem("token") || "");
  const [currentUser, setCurrentUser] = useState(null);
  const [currentPage, setCurrentPage] = useState("sheet");
  const [profileForm, setProfileForm] = useState({ username: "", email: "", avatarUrl: "" });
  const [passwordForm, setPasswordForm] = useState({ currentPassword: "", newPassword: "" });

  const [sheets, setSheets] = useState([]);
  const [sheetSearch, setSheetSearch] = useState("");
  const [cellSearch, setCellSearch] = useState("");
  const [cellSearchMatches, setCellSearchMatches] = useState([]);
  const [cellSearchLoading, setCellSearchLoading] = useState(false);
  const [cellSearchIndex, setCellSearchIndex] = useState(-1);
  const [selectedSheet, setSelectedSheet] = useState(null);
  const [selectedCell, setSelectedCell] = useState(null);
  const [selectedRange, setSelectedRange] = useState(null);
  const [visibleRows, setVisibleRows] = useState(INITIAL_VISIBLE_ROWS);
  const [sheetRowTotal, setSheetRowTotal] = useState(0);
  const [exportRange, setExportRange] = useState({ from: "1", to: "" });
  const [taxExportPeriod, setTaxExportPeriod] = useState(getDefaultTaxExportPeriod);
  const [serverPendingCodeRows, setServerPendingCodeRows] = useState([]);
  const [visitedPendingCodeRows, setVisitedPendingCodeRows] = useState([]);
  const [pendingCodeCursor, setPendingCodeCursor] = useState(null);
  const [rowsLoading, setRowsLoading] = useState(false);
  const [gridScrollTop, setGridScrollTop] = useState(0);
  const [role, setRole] = useState(null);

  const [workspaces, setWorkspaces] = useState([]);
  const [selectedWorkspaceId, setSelectedWorkspaceId] = useState("");
  const [workspaceName, setWorkspaceName] = useState("Main Workspace");
  const [workspaceMemberEmail, setWorkspaceMemberEmail] = useState("");
  const [workspaceMemberRole, setWorkspaceMemberRole] = useState("member");

  const [sheetName, setSheetName] = useState("My First Sheet");
  const [erpType, setErpType] = useState("custom");
  const [analytics, setAnalytics] = useState(null);

  const [menuOpen, setMenuOpen] = useState(false);
  const [contextMenu, setContextMenu] = useState(null);
  const [shareEmail, setShareEmail] = useState("");
  const [shareRole, setShareRole] = useState("viewer");
  const [onlineUsers, setOnlineUsers] = useState([]);
  const [message, setMessage] = useState("");
  const [savingStatus, setSavingStatus] = useState("");
  const [connectionStatus, setConnectionStatus] = useState(token ? "connecting" : "offline");

  const [changesPanelOpen, setChangesPanelOpen] = useState(false);
  const [cellChanges, setCellChanges] = useState([]);
  const [itemDescriptionOptions, setItemDescriptionOptions] = useState(null);
  const [itemDescriptionOptionsLoading, setItemDescriptionOptionsLoading] = useState(false);
  const [descriptionBuilder, setDescriptionBuilder] = useState({
    open: false,
    rowIndex: null,
    type: "",
    category: "",
    filters: {},
    combination: "",
  });

  const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);
  const [erpOptionsOpen, setErpOptionsOpen] = useState(false);
  const [erpOptions, setErpOptions] = useState(defaultErpOptions);
  const [optionText, setOptionText] = useState("");
  const [optionTarget, setOptionTarget] = useState("mainGroups");
  const [optionParent, setOptionParent] = useState("");
  const [supportSourceParent, setSupportSourceParent] = useState("");

  const socketRef = useRef(null);
  const selectedSheetRef = useRef(null);
  const saveTimerRef = useRef(null);
  const cellSaveTimerRef = useRef(null);
  const cellSaveInFlightRef = useRef(false);
  const pendingCellSaveRef = useRef(null);
  const localCellDraftsRef = useRef(new Map());
  const rowsLoadingRef = useRef(false);
  const scrollFrameRef = useRef(null);
  const fileInputRef = useRef(null);
  const contextMenuRef = useRef(null);
  const gridRef = useRef(null);

  const canEdit = role === "owner" || role === "admin" || role === "editor";
  const canManage = role === "owner";
  const canManageSheetUsers = role === "owner" || role === "admin";
  const canGrantSheetAdmin = role === "owner";
  const canBypassRowLocks = role === "owner" || role === "admin";
  const firstConfirmationLockMessage = "First Confirmation can only be edited from 8:00 AM to 3:30 PM Cairo time";

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

  const isFirstConfirmationColumn = (colIndex) => colIndex === FIRST_CONFIRMATION_COLUMN_INDEX;
  const isRowLoaded = (rowIndex) => Array.isArray(selectedSheet?.data?.[rowIndex]);
  const getSheetCell = (rowIndex, colIndex) => normalizeCell(selectedSheet?.data?.[rowIndex]?.[colIndex]);
  const isFirstConfirmationEditOpen = () => {
    const minutes = cairoTimeMinutes();

    return minutes >= FIRST_CONFIRMATION_EDIT_START_MINUTE && minutes <= FIRST_CONFIRMATION_EDIT_END_MINUTE;
  };

  const canEditProtectedRow = (rowIndex) => {
    if (!canEdit) return false;
    if (canBypassRowLocks) return true;

    const rowOwner = selectedSheet?.rowOwners?.[rowIndex];
    return !rowOwner || rowOwner.userId === currentUser?.id;
  };
  const canEditCell = (rowIndex, colIndex) => (
    canEdit &&
    isRowLoaded(rowIndex) &&
    (!isFirstConfirmationColumn(colIndex) || isFirstConfirmationEditOpen()) &&
    (
      colIndex > ROW_LOCK_LAST_COLUMN_INDEX ||
      canEditProtectedRow(rowIndex)
    )
  );
  const getRowLockMessage = (rowIndex) => {
    if (!isRowLoaded(rowIndex)) return "Loading row...";

    const rowOwner = selectedSheet?.rowOwners?.[rowIndex];
    return rowOwner
      ? `Row ${rowIndex + 1} is locked by ${rowOwner.username || rowOwner.email || "another user"}`
      : "You cannot edit this row";
  };

  const getCellLockMessage = (rowIndex, colIndex) => {
    if (isFirstConfirmationColumn(colIndex) && !isFirstConfirmationEditOpen()) {
      return firstConfirmationLockMessage;
    }

    return getRowLockMessage(rowIndex);
  };
  const statusText = savingStatus || (
    connectionStatus === "online"
      ? "Online"
      : connectionStatus === "connecting"
        ? "Connecting..."
        : "Offline"
  );

  const getOptionParentSuggestions = (target = optionTarget) => {
    if (target === "subGroups") return Array.from(new Set(erpOptions.mainGroups || []));
    if (target === "subSubGroups") {
      return Array.from(new Set(Object.values(erpOptions.subGroups || {}).flat()));
    }
    if (target === "supportGroups") {
      return Array.from(new Set(Object.values(erpOptions.subGroups || {}).flat()));
    }
    if (target === "detailedGroups") {
      return Array.from(new Set(Object.values(erpOptions.supportGroups || {}).flat()));
    }

    return [];
  };

  const getSupportGroupOptions = (subGroup) => {
    if (!subGroup) return [];

    const directOptions = erpOptions.supportGroups?.[subGroup];
    if (Array.isArray(directOptions)) return directOptions;

    return defaultErpOptions.supportGroups?.[subGroup] || [];
  };

  const normalizeCell = (cell) => {
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

  const createEmptyRow = () => Array.from({ length: COLS }, () => normalizeCell(""));

  const normalizeData = (data = [], minRows = MIN_SHEET_ROWS) => {
    return Array.from({ length: Math.max(minRows, data.length) }, (_, r) => {
      const row = data[r] || [];
      const sourceRow =
        row.length > COLS
          ? row.filter((_, index) => index !== LEGACY_PACKAGE_COLUMN_INDEX)
          : row;

      return Array.from({ length: COLS }, (_, c) =>
        normalizeCell(sourceRow[c] || "")
      );
    });
  };

  const normalizeSheet = (sheet, options = {}) => ({
    ...sheet,
    data: normalizeData(sheet.data, options.minRows ?? MIN_SHEET_ROWS),
    rowOwners: { ...(sheet.rowOwners || {}) },
    meta: { ...defaultMeta, ...(sheet.meta || {}) },
    erpOptions: { ...defaultErpOptions, ...(sheet.erpOptions || {}) },
  });

  const excelColName = (index) => {
    let name = "";
    let n = index + 1;

    while (n > 0) {
      const rem = (n - 1) % 26;
      name = String.fromCharCode(65 + rem) + name;
      n = Math.floor((n - 1) / 26);
    }

    return name;
  };

  const colName = (index) => {
    return visibleErpHeaders[index] || erpArabicHeaders[index] || `Column ${index + 1}`;
  };

  const estimateTextWidth = (text) => {
    const longestLine = String(text || "")
      .split(/\r?\n/)
      .reduce((longest, line) => Math.max(longest, line.length), 0);

    return Math.ceil(longestLine * AVERAGE_CHAR_WIDTH + COLUMN_WIDTH_PADDING);
  };

  const getAutoColumnWidth = (colIndex, rows = []) => {
    const limits = AUTO_COLUMN_LIMITS[colIndex] || { min: 90, max: 220 };
    const headerWidth = estimateTextWidth(colName(colIndex));
    const dataWidth = rows.reduce((maxWidth, row, rowIndex) => {
      const cellText = normalizeCell(row?.[colIndex]).value;
      const cellWidth = estimateTextWidth(cellText);
      const dropdownOptions = getDropdownOptions(rowIndex, colIndex);
      const optionWidth = dropdownOptions
        ? dropdownOptions.reduce(
            (optionMax, option) => Math.max(optionMax, estimateTextWidth(option)),
            0
          )
        : 0;

      return Math.max(maxWidth, cellWidth, optionWidth);
    }, 0);

    return clamp(Math.max(headerWidth, dataWidth), limits.min, limits.max);
  };

  const getColumnWidth = (colIndex) => {
    return autoColumnWidths[colIndex] || DEFAULT_COLUMN_WIDTHS[colIndex] || 120;
  };

  const getTextLineHeight = (cellStyle = {}) => {
    const fontSize = Number.parseInt(cellStyle.fontSize, 10) || 14;
    return Math.max(DEFAULT_LINE_HEIGHT, Math.round(fontSize * 1.35));
  };

  const estimateAutoFitCellHeight = (cell, colIndex) => {
    const normalizedCell = normalizeCell(cell);
    const value = String(normalizedCell.value || normalizedCell.formula || "");
    if (!value) return DEFAULT_ROW_HEIGHT;

    const fontSize = Number.parseInt(normalizedCell.style.fontSize, 10) || 14;
    const usableWidth = Math.max(24, getColumnWidth(colIndex) - CELL_HORIZONTAL_PADDING);
    const averageCharWidth = Math.max(7, fontSize * 0.55);
    const charsPerLine = Math.max(1, Math.floor(usableWidth / averageCharWidth));
    const lineCount = value
      .split(/\r?\n/)
      .reduce((total, line) => total + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);

    return Math.max(
      DEFAULT_ROW_HEIGHT,
      lineCount * getTextLineHeight(normalizedCell.style) + CELL_VERTICAL_PADDING
    );
  };

  const getRowHeight = (row, rowIndex) => {
    const manualHeight = selectedSheet?.meta?.rowHeights?.[rowIndex] || DEFAULT_ROW_HEIGHT;
    const autoHeight = Math.max(
      estimateAutoFitCellHeight(row?.[0], 0),
      estimateAutoFitCellHeight(row?.[1], 1)
    );

    return Math.max(manualHeight, autoHeight);
  };

  const getRowPalette = (rowIndex) => rowPalette[rowIndex % rowPalette.length];

  const getCellBackground = (cellStyle, rowIndex) => {
    if (cellStyle.backgroundColor && cellStyle.backgroundColor !== defaultCellStyle.backgroundColor) {
      return cellStyle.backgroundColor;
    }

    return getRowPalette(rowIndex).backgroundColor;
  };

  const openCellContextMenu = (event, rowIndex, colIndex) => {
    event.preventDefault();
    event.stopPropagation();

    setSelectedCell({ rowIndex, colIndex });
    if (!isCellInSelectedRange(rowIndex, colIndex)) {
      setSelectedRange({
        start: { row: rowIndex, col: colIndex },
        end: { row: rowIndex, col: colIndex },
      });
    }
    setContextMenu({
      x: event.clientX,
      y: event.clientY,
    });
  };

  const closeContextMenu = () => setContextMenu(null);

  const runContextMenuAction = async (action) => {
    try {
      await action();
    } catch (error) {
      showMessage(error?.message || "Action failed");
    } finally {
      closeContextMenu();
    }
  };

  const cellAddress = (row, col) => excelColName(col) + (row + 1);

  const parseAddress = (address) => {
    const match = String(address).toUpperCase().match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;

    const letters = match[1];
    const row = Number(match[2]) - 1;

    let col = 0;
    for (let i = 0; i < letters.length; i++) {
      col = col * 26 + (letters.charCodeAt(i) - 64);
    }

    return { row, col: col - 1 };
  };

  const getCellNumber = (data, row, col) => {
    const cell = normalizeCell(data?.[row]?.[col]);
    const number = Number(cell.value);
    return Number.isFinite(number) ? number : 0;
  };

  const evaluateFormula = (formula, data) => {
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

      for (let r = rowStart; r <= rowEnd; r++) {
        for (let c = colStart; c <= colEnd; c++) {
          values.push(getCellNumber(data, r, c));
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
        : arg.split(",").map((x) => {
            const addr = parseAddress(x.trim());
            return addr ? getCellNumber(data, addr.row, addr.col) : Number(x) || 0;
          });

      if (fn === "SUM") return String(values.reduce((a, b) => a + b, 0));
      if (fn === "AVERAGE") {
        return String(values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0);
      }
      if (fn === "MIN") return String(Math.min(...values));
      if (fn === "MAX") return String(Math.max(...values));
      if (fn === "COUNT") return String(values.filter((v) => Number.isFinite(v)).length);
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

  const recalculateData = (data) => {
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

  const showMessage = (text) => {
    setMessage(text);
    setTimeout(() => setMessage(""), 3500);
  };

  const authFetch = (url, options = {}) => {
    const savedToken = sessionStorage.getItem("token");

    return fetch(url, {
      ...options,
      headers: {
        "Content-Type": "application/json",
        ...(options.headers || {}),
        Authorization: "Bearer " + savedToken,
      },
    });
  };

  const handleAuth = async () => {
    try {
      const res = await fetch(API_URL + "/" + mode, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email, password, username }),
      });

      const data = await res.json();

      if (!res.ok) {
        showMessage(data.message || "Authentication failed");
        return;
      }

      if (data.token) {
        sessionStorage.setItem("token", data.token);
        setToken(data.token);
        setConnectionStatus("connecting");
        setCurrentUser(data.user || null);
        setProfileForm({
          username: data.user?.username || "",
          email: data.user?.email || "",
          avatarUrl: data.user?.avatarUrl || "",
        });
      }

      if (mode === "signup") setMode("login");
    } catch {
      showMessage("Cannot connect to backend");
    }
  };

  const loadMe = async () => {
    const res = await authFetch(API_URL + "/me");
    const data = await res.json();
    if (res.ok) {
      setCurrentUser(data.user);
      setProfileForm({
        username: data.user?.username || "",
        email: data.user?.email || "",
        avatarUrl: data.user?.avatarUrl || "",
      });
    }
  };

  const saveProfile = async () => {
    const res = await authFetch(API_URL + "/me", {
      method: "PATCH",
      body: JSON.stringify(profileForm),
    });
    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to update profile");
      return;
    }

    if (data.token) {
      sessionStorage.setItem("token", data.token);
      setToken(data.token);
    }

    setCurrentUser(data.user);
    setProfileForm({
      username: data.user?.username || "",
      email: data.user?.email || "",
      avatarUrl: data.user?.avatarUrl || "",
    });
    await Promise.all([loadSheets(), loadWorkspaces()]);
    showMessage("Profile updated");
  };

  const updateProfileImage = (file) => {
    if (!file) return;

    if (!["image/png", "image/jpeg", "image/webp"].includes(file.type)) {
      showMessage("Use PNG, JPG, or WEBP image");
      return;
    }

    if (file.size > 2 * 1024 * 1024) {
      showMessage("Profile image must be under 2 MB");
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      setProfileForm((form) => ({
        ...form,
        avatarUrl: String(reader.result || ""),
      }));
    };
    reader.readAsDataURL(file);
  };

  const changePassword = async () => {
    const res = await authFetch(API_URL + "/me/password", {
      method: "PATCH",
      body: JSON.stringify(passwordForm),
    });
    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to change password");
      return;
    }

    setPasswordForm({ currentPassword: "", newPassword: "" });
    showMessage("Password changed");
  };

  const loadWorkspaces = async () => {
    const res = await authFetch(API_URL + "/workspaces");
    const data = await res.json();
    if (res.ok) setWorkspaces(data);
  };

  const createWorkspace = async () => {
    const res = await authFetch(API_URL + "/workspaces", {
      method: "POST",
      body: JSON.stringify({ name: workspaceName || "Workspace" }),
    });

    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to create workspace");
      return;
    }

    setSelectedWorkspaceId(data._id);
    await loadWorkspaces();
    showMessage("Workspace created");
  };

  const addWorkspaceMember = async () => {
    if (!selectedWorkspaceId) {
      showMessage("Select workspace first");
      return;
    }

    const res = await authFetch(API_URL + "/workspaces/" + selectedWorkspaceId + "/members", {
      method: "POST",
      body: JSON.stringify({
        email: workspaceMemberEmail,
        role: workspaceMemberRole,
      }),
    });

    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to add member");
      return;
    }

    setWorkspaceMemberEmail("");
    await loadWorkspaces();
    showMessage("Workspace member added");
  };

  const loadSheets = async () => {
    const url = selectedWorkspaceId
      ? API_URL + "/sheets?workspaceId=" + selectedWorkspaceId
      : API_URL + "/sheets";

    const res = await authFetch(url);
    const data = await res.json();
    if (res.ok) setSheets(data);
  };

  const loadAnalytics = async () => {
    const res = await authFetch(API_URL + "/admin/analytics");
    const data = await res.json();
    if (res.ok) setAnalytics(data);
  };

  const loadErpOptions = async (sheetId) => {
    const res = await authFetch(API_URL + "/sheet/" + sheetId + "/erp-options");
    const data = await res.json();

    if (res.ok) {
      setErpOptions({
        ...defaultErpOptions,
        ...(data || {}),
      });
    }
  };

  const loadPendingCodeRows = async (sheetId) => {
    if (!sheetId) return;

    const res = await authFetch(API_URL + "/sheet/" + sheetId + "/pending-code-rows");
    const data = await res.json();

    if (res.ok) {
      setServerPendingCodeRows(
        Array.from(new Set(Array.isArray(data.rowIndexes) ? data.rowIndexes : []))
          .map(Number)
          .filter(Number.isInteger)
          .sort((left, right) => left - right)
      );
    }
  };

  const createSheet = async () => {
    const res = await authFetch(API_URL + "/sheet", {
      method: "POST",
      body: JSON.stringify({
        name: sheetName || "New Sheet",
        workspaceId: selectedWorkspaceId || null,
        erpType,
      }),
    });

    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to create sheet");
      return;
    }

    setSheetName("");
    setErpType("custom");
    await loadSheets();
    await loadAnalytics();
    await openSheet(data._id);
    setMenuOpen(false);
  };

  const buildInitialSheetRows = async (_sheetId, rows, start) => {
    if (!start) return rows;

    return mergeRowsAtStart([], rows, start);
  };

  const openSheet = async (id) => {
    flushPendingCellSave();

    if (socketRef.current && selectedSheet?._id) {
      socketRef.current.emit("leave-sheet", selectedSheet._id);
    }

    const res = await authFetch(API_URL + "/sheet/" + id + "?rowLimit=" + INITIAL_VISIBLE_ROWS + "&focus=lastData");
    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to open sheet");
      return;
    }

    const normalized = normalizeSheet(data.sheet, { minRows: 0 });
    const rowStart = Number(data.rows?.start || 0);
    const lastDataRowIndex = Number(data.rows?.lastDataRowIndex || 0);
    const focusedRows = await buildInitialSheetRows(id, normalized.data, rowStart);
    const focusedSheet = { ...normalized, data: focusedRows };
    const totalRows = normalizeRowTotal(data.rows?.total, focusedRows.length);
    const initialVisibleRows = Math.min(totalRows, Math.max(INITIAL_VISIBLE_ROWS, rowStart + normalized.data.length));
    const initialScrollTop = Math.max(0, lastDataRowIndex * DEFAULT_ROW_HEIGHT);

    setSelectedSheet(focusedSheet);
    setSheetRowTotal(totalRows);
    setErpOptions(normalized.erpOptions || defaultErpOptions);
    setRole(data.role);
    setSelectedCell(null);
    setSelectedRange(null);
    setGridScrollTop(initialScrollTop);
    setVisibleRows(initialVisibleRows);
    setCurrentPage("sheet");
    setMenuOpen(false);

    await Promise.all([
      loadErpOptions(id),
      loadPendingCodeRows(id),
    ]);
    setVisitedPendingCodeRows([]);
    setPendingCodeCursor(null);

    window.requestAnimationFrame(() => {
      if (gridRef.current) gridRef.current.scrollTop = initialScrollTop;
    });

    if (socketRef.current?.connected) {
      socketRef.current.emit("join-sheet", id);
    }
  };

  const saveErpOptions = async (newOptions) => {
    if (!selectedSheet || !canManage) return;

    const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id + "/erp-options", {
      method: "PATCH",
      body: JSON.stringify({ erpOptions: newOptions }),
    });

    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to save ERP options");
      return;
    }

    setErpOptions(data.erpOptions);
    setSelectedSheet((prev) => (prev ? { ...prev, erpOptions: data.erpOptions } : prev));
    showMessage("ERP options saved");
  };

  const addErpOption = async () => {
    const value = optionText.trim();
    if (!value) return;

    const newOptions = {
      ...erpOptions,
      subGroups: { ...(erpOptions.subGroups || {}) },
      subSubGroups: { ...(erpOptions.subSubGroups || {}) },
      supportGroups: { ...(erpOptions.supportGroups || {}) },
      detailedGroups: { ...(erpOptions.detailedGroups || {}) },
    };

    const independentKeys = [
      "mainGroups",
      "units",
      "shelfLife",
      "sequence",
      "confirmation1",
      "confirmation2",
    ];

    if (independentKeys.includes(optionTarget)) {
      newOptions[optionTarget] = Array.from(
        new Set([...(newOptions[optionTarget] || []), value])
      );
    } else {
      if (!optionParent.trim()) {
        showMessage("Select parent value first");
        return;
      }

      const parent = optionParent.trim();

      newOptions[optionTarget] = {
        ...(newOptions[optionTarget] || {}),
        [parent]: Array.from(new Set([...(newOptions[optionTarget]?.[parent] || []), value])),
      };
    }

    setOptionText("");
    await saveErpOptions(newOptions);
  };

  const copySupportOptionsToSubGroup = async () => {
    const targetSubGroup = optionParent.trim();
    const sourceParent = supportSourceParent.trim();

    if (!targetSubGroup || !sourceParent) {
      showMessage("Select sub group and source list first");
      return;
    }

    const sourceOptions = erpOptions.supportGroups?.[sourceParent] || [];

    if (!sourceOptions.length) {
      showMessage("Source list has no options");
      return;
    }

    const newOptions = {
      ...erpOptions,
      supportGroups: {
        ...(erpOptions.supportGroups || {}),
        [targetSubGroup]: Array.from(new Set([
          ...(erpOptions.supportGroups?.[targetSubGroup] || []),
          ...sourceOptions,
        ])),
      },
    };

    setSupportSourceParent("");
    await saveErpOptions(newOptions);
  };

  const deleteSheet = async () => {
    if (!selectedSheet || !canManage) return;

    const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id, {
      method: "DELETE",
    });

    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to delete sheet");
      return;
    }

    setDeleteConfirmOpen(false);
    setSelectedSheet(null);
    setRole(null);
    await loadSheets();
    await loadAnalytics();
    showMessage("Sheet deleted successfully");
  };

  const showCellChanges = async () => {
    if (!selectedSheet || !selectedCell) return;

    const res = await authFetch(
      API_URL +
        "/sheet/" +
        selectedSheet._id +
        "/changes?rowIndex=" +
        selectedCell.rowIndex +
        "&colIndex=" +
        selectedCell.colIndex
    );

    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to load changes");
      return;
    }

    setCellChanges(data);
    setChangesPanelOpen(true);
    setContextMenu(null);
  };

  const showAllChanges = async () => {
    if (!selectedSheet) return;

    const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id + "/changes");
    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to load changes");
      return;
    }

    setCellChanges(data);
    setChangesPanelOpen(true);
    setContextMenu(null);
  };

  const getDropdownOptions = (rowIndex, colIndex) => {
    if (rowIndex === 0) return null;

    const row = selectedSheet?.data?.[rowIndex] || [];

    const mainGroup = normalizeCell(row[2]).value;
    const subGroup = normalizeCell(row[3]).value;
    const supportGroup = normalizeCell(row[5]).value;

    if (colIndex === 2) return erpOptions.mainGroups || [];
    if (colIndex === 3) return erpOptions.subGroups?.[mainGroup] || defaultErpOptions.subGroups?.[mainGroup] || [];
    if (colIndex === 4) return erpOptions.subSubGroups?.[subGroup] || defaultErpOptions.subSubGroups?.[subGroup] || [];
    if (colIndex === 5) return getSupportGroupOptions(subGroup);
    if (colIndex === 6) {
      return erpOptions.detailedGroups?.[supportGroup] || defaultErpOptions.detailedGroups?.[supportGroup] || [];
    }
    if (colIndex === 7) return erpOptions.units || [];
    if (colIndex === 8) return erpOptions.shelfLife || [];
    if (colIndex === 9) return erpOptions.sequence || [];
    if (colIndex === 11) return erpOptions.confirmation1 || [];
    if (colIndex === 13) return erpOptions.confirmation2 || [];

    return null;
  };

  const visibleSheetRows = selectedSheet?.data?.slice(0, AUTO_FIT_SAMPLE_ROWS) || [];
  const autoColumnWidths = Array.from({ length: COLS }, (_, colIndex) =>
    getAutoColumnWidth(colIndex, visibleSheetRows)
  );

  const validateErpOption = (rowIndex, colIndex, options) => {
    const value = normalizeCell(selectedSheet?.data?.[rowIndex]?.[colIndex]).value;

    if (!value || options.includes(value)) return;

    updateCell(rowIndex, colIndex, "");
    showMessage("Select a value from the list");
  };

  const queueMetadataSave = () => {
    setSavingStatus("Unsaved changes...");
  };

  const applyRowOwnerUpdates = (currentOwners = {}, rowOwners = {}) => {
    const nextOwners = { ...currentOwners };

    Object.entries(rowOwners || {}).forEach(([rowIndex, owner]) => {
      if (owner) {
        nextOwners[rowIndex] = owner;
      } else {
        delete nextOwners[rowIndex];
      }
    });

    return nextOwners;
  };

  const mergeRowOwners = (rowOwners = {}) => {
    if (!rowOwners || Object.keys(rowOwners).length === 0) return;

    setSelectedSheet((prev) => (
      prev
        ? {
            ...prev,
            rowOwners: applyRowOwnerUpdates(prev.rowOwners, rowOwners),
          }
        : prev
    ));
  };

  const applyPendingCodeUpdates = (updates = []) => {
    if (!Array.isArray(updates) || updates.length === 0) return;

    setServerPendingCodeRows((rows) => {
      const rowSet = new Set(rows);

      updates.forEach((update) => {
        const numericRowIndex = Number(update?.rowIndex);
        if (!Number.isInteger(numericRowIndex)) return;

        if (update.pending) {
          rowSet.add(numericRowIndex);
        } else {
          rowSet.delete(numericRowIndex);
        }
      });

      return Array.from(rowSet).sort((left, right) => left - right);
    });
  };

  const handleSocketSaveResult = (result) => {
    mergeRowOwners(result?.rowOwners);
    applyPendingCodeUpdates(result?.pendingCodeUpdates);

    if (!result?.ok) {
      setSavingStatus("Unsaved changes...");
      showMessage(result?.message || "Realtime save failed");
      return;
    }

    setSavingStatus("Saved");
    setTimeout(() => setSavingStatus(""), 1200);
  };

  const getCellPatchKey = (sheetId, patch) => {
    return `${sheetId}:${patch.rowIndex}:${patch.colIndex}`;
  };

  const getPayloadPatches = (payload) => {
    return Array.isArray(payload?.patches) && payload.patches.length > 0
      ? payload.patches
      : [payload];
  };

  const rememberLocalCellDrafts = (payload) => {
    getPayloadPatches(payload).forEach((patch) => {
      localCellDraftsRef.current.set(getCellPatchKey(payload.sheetId, patch), {
        value: patch.value ?? "",
        formula: patch.formula || "",
      });
    });
  };

  const clearSavedLocalCellDrafts = (payload) => {
    getPayloadPatches(payload).forEach((patch) => {
      const key = getCellPatchKey(payload.sheetId, patch);
      const currentDraft = localCellDraftsRef.current.get(key);

      if (
        currentDraft &&
        currentDraft.value === (patch.value ?? "") &&
        currentDraft.formula === (patch.formula || "")
      ) {
        localCellDraftsRef.current.delete(key);
      }
    });
  };

  const saveCellPatchesHttp = async (payload) => {
    if (!payload?.sheetId) return false;

    const res = await authFetch(API_URL + "/sheet/" + payload.sheetId + "/cells", {
      method: "PATCH",
      body: JSON.stringify(payload),
    });
    const data = await res.json();

    if (!res.ok) {
      setSavingStatus("Unsaved changes...");
      showMessage(data.message || "Failed to save cell");
      return false;
    }

    mergeRowOwners(data.rowOwners);
    applyPendingCodeUpdates(data.pendingCodeUpdates);
    setSavingStatus("Saved");
    setTimeout(() => setSavingStatus(""), 1200);
    return true;
  };

  const saveCellStyleHttp = async (rowIndex, colIndex, style) => {
    if (!selectedSheet?._id) return false;

    const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id + "/cell-style", {
      method: "PATCH",
      body: JSON.stringify({ rowIndex, colIndex, style }),
    });
    const data = await res.json();

    if (!res.ok) {
      setSavingStatus("Unsaved changes...");
      showMessage(data.message || "Failed to save style");
      return false;
    }

    mergeRowOwners(data.rowOwners);
    setSavingStatus("Saved");
    setTimeout(() => setSavingStatus(""), 1200);
    return true;
  };

  const flushPendingCellSave = () => {
    if (cellSaveTimerRef.current) {
      clearTimeout(cellSaveTimerRef.current);
      cellSaveTimerRef.current = null;
    }

    if (cellSaveInFlightRef.current) return;

    const payload = pendingCellSaveRef.current;
    if (!payload) return;
    pendingCellSaveRef.current = null;
    cellSaveInFlightRef.current = true;

    void saveCellPatchesHttp(payload)
      .then((saved) => {
        if (saved) clearSavedLocalCellDrafts(payload);
      })
      .catch(() => {
        setSavingStatus("Unsaved changes...");
        showMessage("Failed to save cell");
      })
      .finally(() => {
        cellSaveInFlightRef.current = false;

        if (pendingCellSaveRef.current) {
          cellSaveTimerRef.current = setTimeout(flushPendingCellSave, 0);
        }
      });
  };

  const queueCellSocketSave = (payload) => {
    const lockedPatch = getPayloadPatches(payload).find(
      (patch) => !canEditCell(patch.rowIndex, patch.colIndex)
    );
    if (lockedPatch) return;

    rememberLocalCellDrafts(payload);
    const pending = pendingCellSaveRef.current;

    if (pending && pending.sheetId === payload.sheetId) {
      const patchMap = new Map();
      [...(pending.patches || []), ...(payload.patches || [])].forEach((patch) => {
        patchMap.set(`${patch.rowIndex}:${patch.colIndex}`, patch);
      });

      const mergedPatches = Array.from(patchMap.values());
      pendingCellSaveRef.current = {
        ...payload,
        patches: mergedPatches,
        rowIndex: mergedPatches[0]?.rowIndex ?? payload.rowIndex,
        colIndex: mergedPatches[0]?.colIndex ?? payload.colIndex,
        value: mergedPatches[0]?.value ?? payload.value,
        formula: mergedPatches[0]?.formula ?? payload.formula,
      };
    } else {
      pendingCellSaveRef.current = payload;
    }

    setSavingStatus("Saving...");

    if (cellSaveTimerRef.current) {
      clearTimeout(cellSaveTimerRef.current);
    }

    cellSaveTimerRef.current = setTimeout(flushPendingCellSave, CELL_SAVE_DEBOUNCE_MS);
  };

  const shouldRecalculateForPatches = (patches) => {
    return (
      selectedSheet?.analytics?.totalFormulaCells > 0 ||
      patches.some((patch) => patch.formula || String(patch.value || "").startsWith("="))
    );
  };

  const touchesPendingCodeColumns = (patches = []) => {
    return patches.some((patch) => (
      Number(patch.colIndex) === FIRST_CONFIRMATION_COLUMN_INDEX ||
      Number(patch.colIndex) === TAX_ITEM_CODE_COLUMN_INDEX
    ));
  };

  const mergeRowsAtStart = (currentRows, incomingRows, start) => {
    const nextRows = [...currentRows];

    incomingRows.forEach((row, index) => {
      nextRows[start + index] = row;
    });

    return nextRows;
  };

  const patchMatchesCell = (cell, patch) => {
    const normalizedCell = normalizeCell(cell);

    return (
      String(normalizedCell.value ?? "") === String(patch.value ?? "") &&
      (normalizedCell.formula || "") === (patch.formula || "")
    );
  };

  const applyCellPatchesLocal = (patches, options = {}) => {
    if (!canEdit) {
      showMessage("Viewer access: you cannot edit this sheet");
      return;
    }

    const lockedPatch = patches.find((patch) => !canEditCell(patch.rowIndex, patch.colIndex));
    if (lockedPatch) {
      showMessage(getCellLockMessage(lockedPatch.rowIndex, lockedPatch.colIndex));
      return;
    }

    setSelectedSheet((prev) => {
      if (!prev) return prev;

      const data = [...prev.data];
      const touchedRows = new Set(patches.map((patch) => patch.rowIndex));

      touchedRows.forEach((rowIndex) => {
        data[rowIndex] = [...(data[rowIndex] || [])];
      });

      patches.forEach((patch) => {
        const currentCell = normalizeCell(data[patch.rowIndex]?.[patch.colIndex]);
        const patchFormula = patch.formula || "";

        data[patch.rowIndex][patch.colIndex] = {
          ...currentCell,
          value: patchFormula ? evaluateFormula(patchFormula, data) : patch.value,
          formula: patchFormula,
        };
      });

      return {
        ...prev,
        data: options.recalculate ? recalculateData(data) : data,
      };
    });

    if (options.queueSave !== false) {
      setSavingStatus("Unsaved changes...");
    }
  };

  const applyCellStyleLocal = (rowIndex, colIndex, style, options = {}) => {
    if (!canEdit) {
      showMessage("Viewer access: you cannot edit this sheet");
      return;
    }

    if (!canEditCell(rowIndex, colIndex)) {
      showMessage(getCellLockMessage(rowIndex, colIndex));
      return;
    }

    setSelectedSheet((prev) => {
      if (!prev) return prev;

      const data = [...prev.data];
      data[rowIndex] = [...(data[rowIndex] || [])];
      const target = normalizeCell(data[rowIndex]?.[colIndex]);

      data[rowIndex][colIndex] = {
        ...target,
        style: options.replace ? { ...style } : { ...target.style, ...style },
      };

      return { ...prev, data };
    });

    if (options.queueSave !== false) {
      setSavingStatus("Unsaved changes...");
    }
  };

  const updateCell = (rowIndex, colIndex, inputValue) => {
    if (!canEditCell(rowIndex, colIndex)) {
      showMessage(getCellLockMessage(rowIndex, colIndex));
      return;
    }

    if (colIndex === 1 && rowIndex > 0) {
      showMessage("Description is auto-filled from Item Name");
      return;
    }

    const makeCellPatch = (row, col, value) => ({
      rowIndex: row,
      colIndex: col,
      value: String(value).startsWith("=") ? evaluateFormula(value, selectedSheet?.data || []) : value,
      formula: String(value).startsWith("=") ? value : "",
    });

    const patches = [makeCellPatch(rowIndex, colIndex, inputValue)];

    if (colIndex === 0 && rowIndex > 0) {
      patches.push(makeCellPatch(rowIndex, 1, inputValue));
    }

    if (colIndex === 2 && rowIndex > 0) {
      [3, 4, 5, 6].forEach((col) => patches.push(makeCellPatch(rowIndex, col, "")));
    }

    if (colIndex === 3 && rowIndex > 0) {
      [4, 5, 6].forEach((col) => patches.push(makeCellPatch(rowIndex, col, "")));
    }

    if (colIndex === 4 && rowIndex > 0) {
      [5, 6].forEach((col) => patches.push(makeCellPatch(rowIndex, col, "")));
    }

    if (colIndex === 5 && rowIndex > 0) {
      patches.push(makeCellPatch(rowIndex, 6, ""));
    }

    applyCellPatchesLocal(patches, {
      queueSave: !socketRef.current,
      recalculate: shouldRecalculateForPatches(patches),
    });

    if (selectedSheet) {
      queueCellSocketSave({
        sheetId: selectedSheet._id,
        rowIndex,
        colIndex,
        value: patches[0].value,
        formula: patches[0].formula,
        patches,
      });
    }
  };

  const updateCellStyle = (styleKey, styleValue) => {
    if (!selectedCell || !selectedSheet || !canEdit) return;

    const { rowIndex, colIndex } = selectedCell;
    if (!canEditCell(rowIndex, colIndex)) {
      showMessage(getCellLockMessage(rowIndex, colIndex));
      return;
    }

    applyCellStyleLocal(rowIndex, colIndex, { [styleKey]: styleValue }, { queueSave: !socketRef.current });

    if (socketRef.current?.connected) {
      setSavingStatus("Saving...");
      socketRef.current.emit("cell-style-change", {
        sheetId: selectedSheet._id,
        rowIndex,
        colIndex,
        style: { [styleKey]: styleValue },
      }, (result) => {
        if (!result?.ok) {
          void saveCellStyleHttp(rowIndex, colIndex, { [styleKey]: styleValue });
          return;
        }
        handleSocketSaveResult(result);
      });
    } else {
      void saveCellStyleHttp(rowIndex, colIndex, { [styleKey]: styleValue });
    }
  };

  const resetSelectedCellStyle = () => {
    if (!selectedCell || !selectedSheet || !canEdit) return;

    const { rowIndex, colIndex } = selectedCell;
    if (!canEditCell(rowIndex, colIndex)) {
      showMessage(getCellLockMessage(rowIndex, colIndex));
      return;
    }

    applyCellStyleLocal(rowIndex, colIndex, defaultCellStyle, {
      queueSave: !socketRef.current,
      replace: true,
    });

    if (socketRef.current?.connected) {
      setSavingStatus("Saving...");
      socketRef.current.emit("cell-style-change", {
        sheetId: selectedSheet._id,
        rowIndex,
        colIndex,
        style: { ...defaultCellStyle },
      }, (result) => {
        if (!result?.ok) {
          void saveCellStyleHttp(rowIndex, colIndex, { ...defaultCellStyle });
          return;
        }
        handleSocketSaveResult(result);
      });
    } else {
      void saveCellStyleHttp(rowIndex, colIndex, { ...defaultCellStyle });
    }
  };

  const mergeCells = () => {
    if (!selectedRange || !selectedSheet || !canEdit) return;

    const rowStart = Math.min(selectedRange.start.row, selectedRange.end.row);
    const rowEnd = Math.max(selectedRange.start.row, selectedRange.end.row);
    const colStart = Math.min(selectedRange.start.col, selectedRange.end.col);
    const colEnd = Math.max(selectedRange.start.col, selectedRange.end.col);

    setSelectedSheet((prev) => ({
      ...prev,
      meta: {
        ...prev.meta,
        merges: [...(prev.meta?.merges || []), { rowStart, rowEnd, colStart, colEnd }],
      },
    }));

    queueMetadataSave();
  };

  const unmergeCells = () => {
    if (!selectedCell || !selectedSheet || !canEdit) return;

    const { rowIndex, colIndex } = selectedCell;

    setSelectedSheet((prev) => ({
      ...prev,
      meta: {
        ...prev.meta,
        merges: (prev.meta?.merges || []).filter(
          (m) =>
            !(
              rowIndex >= m.rowStart &&
              rowIndex <= m.rowEnd &&
              colIndex >= m.colStart &&
              colIndex <= m.colEnd
            )
        ),
      },
    }));

    queueMetadataSave();
  };

  const getMergeInfo = (rowIndex, colIndex) => {
    const merges = selectedSheet?.meta?.merges || [];
    const merge = merges.find(
      (m) =>
        rowIndex >= m.rowStart &&
        rowIndex <= m.rowEnd &&
        colIndex >= m.colStart &&
        colIndex <= m.colEnd
    );

    if (!merge) return null;
    const isMaster = rowIndex === merge.rowStart && colIndex === merge.colStart;
    return { ...merge, isMaster };
  };

  const isCellInSelectedRange = (rowIndex, colIndex) => {
    if (!selectedRange) return false;

    const rowStart = Math.min(selectedRange.start.row, selectedRange.end.row);
    const rowEnd = Math.max(selectedRange.start.row, selectedRange.end.row);
    const colStart = Math.min(selectedRange.start.col, selectedRange.end.col);
    const colEnd = Math.max(selectedRange.start.col, selectedRange.end.col);

    return (
      rowIndex >= rowStart &&
      rowIndex <= rowEnd &&
      colIndex >= colStart &&
      colIndex <= colEnd
    );
  };

  const selectRow = (rowIndex) => {
    const lastCol = Math.max((selectedSheet?.data?.[rowIndex]?.length || 1) - 1, 0);

    setSelectedCell({ rowIndex, colIndex: 0 });
    setSelectedRange({
      start: { row: rowIndex, col: 0 },
      end: { row: rowIndex, col: lastCol },
    });
  };

  const resizeRow = (rowIndex) => {
    const height = prompt("Row height", selectedSheet?.meta?.rowHeights?.[rowIndex] || DEFAULT_ROW_HEIGHT);
    if (!height) return;

    setSelectedSheet((prev) => ({
      ...prev,
      meta: {
        ...prev.meta,
        rowHeights: { ...(prev.meta?.rowHeights || {}), [rowIndex]: Number(height) },
      },
    }));

    queueMetadataSave();
  };

  const loadMoreRows = async () => {
    if (!selectedSheet || rowsLoadingRef.current) return;

    const loadedRows = selectedSheet.data.length;

    if (visibleRows < loadedRows) {
      setVisibleRows((count) => Math.min(count + ROW_LOAD_STEP, loadedRows));
      return;
    }

    if (loadedRows >= sheetRowTotal) return;

    rowsLoadingRef.current = true;
    setRowsLoading(true);

    try {
      const res = await authFetch(
        API_URL + "/sheet/" + selectedSheet._id + "/rows?start=" + loadedRows + "&limit=" + ROW_LOAD_STEP
      );
      const data = await res.json();

      if (!res.ok) {
        showMessage(data.message || "Failed to load rows");
        return;
      }

      const responseStart = Number.isInteger(data.start) ? data.start : loadedRows;
      const expectedRowCount = Math.min(ROW_LOAD_STEP, sheetRowTotal - responseStart);
      const normalizedRows = normalizeData(data.rows || [], expectedRowCount).slice(0, expectedRowCount);

      setSelectedSheet((prev) => {
        if (!prev || prev._id !== selectedSheet._id) return prev;

        return {
          ...prev,
          data: mergeRowsAtStart(prev.data, normalizedRows, responseStart),
          rowOwners: { ...(prev.rowOwners || {}), ...(data.rowOwners || {}) },
        };
      });
      setSheetRowTotal(normalizeRowTotal(data.total, sheetRowTotal));
      setVisibleRows((count) => Math.min(
        Math.max(count + ROW_LOAD_STEP, responseStart + normalizedRows.length),
        normalizeRowTotal(data.total, sheetRowTotal)
      ));
    } catch {
      showMessage("Failed to load rows");
    } finally {
      rowsLoadingRef.current = false;
      setRowsLoading(false);
    }
  };

  const handleGridScroll = (event) => {
    const grid = event.currentTarget;
    const nextScrollTop = grid.scrollTop;
    const nearBottom = grid.scrollHeight - nextScrollTop - grid.clientHeight <= DEFAULT_ROW_HEIGHT * 4;

    if (scrollFrameRef.current) {
      window.cancelAnimationFrame(scrollFrameRef.current);
    }

    scrollFrameRef.current = window.requestAnimationFrame(() => {
      setGridScrollTop(nextScrollTop);
      scrollFrameRef.current = null;
    });

    if (nearBottom) {
      void loadMoreRows();
    }
  };

  const ensureRowsLoaded = async (targetCount, options = {}) => {
    if (!selectedSheet) return;

    const loadedRowMap = selectedSheet.data || [];
    const targetStart = Math.max(0, Number(options.start ?? 0));
    const targetEnd = Math.min(Math.max(targetCount, targetStart), sheetRowTotal);
    if (targetEnd <= targetStart) return;

    const missingRanges = [];
    let rangeStart = null;

    for (let rowIndex = targetStart; rowIndex < targetEnd; rowIndex += 1) {
      if (loadedRowMap[rowIndex]) {
        if (rangeStart !== null) {
          missingRanges.push({ start: rangeStart, end: rowIndex });
          rangeStart = null;
        }
      } else if (rangeStart === null) {
        rangeStart = rowIndex;
      }
    }

    if (rangeStart !== null) {
      missingRanges.push({ start: rangeStart, end: targetEnd });
    }

    if (!missingRanges.length) return;

    setRowsLoading(true);

    const loadedChunks = [];
    const loadedRowOwners = {};

    try {
      for (const range of missingRanges) {
        let loadedRows = range.start;

        while (loadedRows < range.end) {
          const limit = Math.min(ROW_LOAD_STEP, range.end - loadedRows);
          const res = await authFetch(
            API_URL + "/sheet/" + selectedSheet._id + "/rows?start=" + loadedRows + "&limit=" + limit
          );
          const data = await res.json();

          if (!res.ok) {
            showMessage(data.message || "Failed to load rows");
            return;
          }

          const responseStart = Number.isInteger(data.start) ? data.start : loadedRows;
          const expectedRowCount = Math.min(limit, sheetRowTotal - responseStart);
          const normalizedRows = normalizeData(data.rows || [], expectedRowCount).slice(0, expectedRowCount);

          loadedChunks.push({ start: responseStart, rows: normalizedRows });
          Object.assign(loadedRowOwners, data.rowOwners || {});
          loadedRows = responseStart + normalizedRows.length;
          setSheetRowTotal(normalizeRowTotal(data.total, sheetRowTotal));
        }
      }

      if (loadedChunks.length > 0) {
        setSelectedSheet((prev) => {
          if (!prev || prev._id !== selectedSheet._id) return prev;

          const nextData = loadedChunks.reduce(
            (rows, chunk) => mergeRowsAtStart(rows, chunk.rows, chunk.start),
            prev.data
          );

          return {
            ...prev,
            data: nextData,
            rowOwners: { ...(prev.rowOwners || {}), ...loadedRowOwners },
          };
        });
      }
    } finally {
      setRowsLoading(false);
    }
  };

  const parseExportRowNumber = (value, fallback) => {
    const text = String(value ?? "").trim() || String(fallback);
    return /^\d+$/.test(text) ? Number(text) : NaN;
  };

  const getCairoDateDaysAgo = (daysAgo) => {
    const parts = new Intl.DateTimeFormat("en-GB", {
      timeZone: CAIRO_TIMEZONE,
      year: "numeric",
      month: "numeric",
      day: "numeric",
    }).formatToParts(new Date());
    const year = Number(parts.find((part) => part.type === "year")?.value || 1970);
    const month = Number(parts.find((part) => part.type === "month")?.value || 1);
    const day = Number(parts.find((part) => part.type === "day")?.value || 1);
    const date = new Date(year, month - 1, day, 12, 0, 0, 0);

    date.setDate(date.getDate() - daysAgo);
    return date;
  };

  const getCellExportText = (row, colIndex) => {
    const normalized = normalizeCell(row?.[colIndex]);
    return String(normalized.value || normalized.formula || "").trim();
  };

  const uniqueDescriptionValues = (rows, key) => (
    Array.from(new Set(rows.map((row) => row?.[key]).filter(Boolean)))
  );

  const loadItemDescriptionOptions = async () => {
    if (itemDescriptionOptions) return itemDescriptionOptions;

    setItemDescriptionOptionsLoading(true);
    try {
      const module = await import("./itemDescriptionOptions.json");
      setItemDescriptionOptions(module.default);
      return module.default;
    } catch {
      showMessage("Failed to load item description options");
      return null;
    } finally {
      setItemDescriptionOptionsLoading(false);
    }
  };

  const findDescriptionMatch = (value, options) => {
    const text = String(value || "").trim();
    if (!text) return null;

    for (const type of Object.keys(options || {})) {
      const option = options[type];
      const row = option?.rows?.find((item) => item.combination === text);
      if (row) return { type, row };
    }

    return null;
  };

  const openDescriptionBuilder = async (event, rowIndex) => {
    event.preventDefault();
    event.stopPropagation();

    if (!canEditCell(rowIndex, DESCRIPTION_BUILDER_COLUMN_INDEX)) {
      showMessage(getCellLockMessage(rowIndex, DESCRIPTION_BUILDER_COLUMN_INDEX));
      return;
    }

    const options = await loadItemDescriptionOptions();
    if (!options) return;

    const descriptionTypes = Object.keys(options);
    const cellValue = normalizeCell(selectedSheet?.data?.[rowIndex]?.[DESCRIPTION_BUILDER_COLUMN_INDEX]).value;
    const match = findDescriptionMatch(cellValue, options);
    const type = match?.type || descriptionTypes[0] || "";
    const typeOptions = options[type] || EMPTY_DESCRIPTION_OPTIONS;
    const filters = {};

    if (match?.row) {
      typeOptions.fields.forEach((field) => {
        filters[field.key] = match.row[field.key] || "";
      });
    }

    setSelectedCell({ rowIndex, colIndex: DESCRIPTION_BUILDER_COLUMN_INDEX });
    setSelectedRange({
      start: { row: rowIndex, col: DESCRIPTION_BUILDER_COLUMN_INDEX },
      end: { row: rowIndex, col: DESCRIPTION_BUILDER_COLUMN_INDEX },
    });
    setDescriptionBuilder({
      open: true,
      rowIndex,
      type,
      category: match?.row?.category || "",
      filters,
      combination: match?.row?.combination || "",
    });
  };

  const updateDescriptionBuilderType = (type) => {
    setDescriptionBuilder((builder) => ({
      ...builder,
      type,
      category: "",
      filters: {},
      combination: "",
    }));
  };

  const updateDescriptionBuilderCategory = (category) => {
    setDescriptionBuilder((builder) => ({
      ...builder,
      category,
      filters: {},
      combination: "",
    }));
  };

  const updateDescriptionBuilderFilter = (fieldKey, value) => {
    setDescriptionBuilder((builder) => ({
      ...builder,
      filters: { ...builder.filters, [fieldKey]: value },
      combination: "",
    }));
  };

  const applyDescriptionBuilder = () => {
    if (!selectedDescriptionRow || descriptionBuilder.rowIndex === null) {
      showMessage("Select a matching item description first");
      return;
    }

    updateCell(
      descriptionBuilder.rowIndex,
      DESCRIPTION_BUILDER_COLUMN_INDEX,
      selectedDescriptionRow.combination
    );
    setDescriptionBuilder((builder) => ({ ...builder, open: false }));
  };

  const getTaxExportDateBounds = () => {
    const fromDate = new Date(taxExportPeriod.from);
    const toDate = new Date(taxExportPeriod.to || formatCairoDateTimeInput(new Date()));

    if (
      !taxExportPeriod.from ||
      Number.isNaN(fromDate.getTime()) ||
      Number.isNaN(toDate.getTime()) ||
      fromDate > toDate
    ) {
      showMessage("Choose a valid tax export date and time range");
      return null;
    }

    return {
      from: fromDate.toISOString(),
      to: toDate.toISOString(),
    };
  };

  const getExportRangeBounds = () => {
    if (!selectedSheet) return null;

    const totalRows = Math.max(sheetRowTotal || 0, selectedSheet.data?.length || 0);
    if (!totalRows) {
      showMessage("No rows available to export");
      return null;
    }

    const fromRow = parseExportRowNumber(exportRange.from, 1);
    const toRow = parseExportRowNumber(exportRange.to, totalRows);

    if (
      !Number.isInteger(fromRow) ||
      !Number.isInteger(toRow) ||
      fromRow < 1 ||
      toRow < fromRow ||
      toRow > totalRows
    ) {
      showMessage(`Choose export rows between 1 and ${totalRows}`);
      return null;
    }

    return {
      fromRow,
      toRow,
      startIndex: fromRow - 1,
      endIndex: toRow,
      label: `rows-${fromRow}-to-${toRow}`,
    };
  };

  const loadRowsForExportRange = async (range) => {
    if (!selectedSheet || !range) return [];

    const expectedRows = range.endIndex - range.startIndex;
    const rows = [];
    let loadedRows = 0;

    while (loadedRows < expectedRows) {
      const requestStart = range.startIndex + loadedRows;
      const limit = Math.min(EXPORT_ROW_BATCH_SIZE, expectedRows - loadedRows);
      const res = await authFetch(
        API_URL + "/sheet/" + selectedSheet._id + "/rows?start=" + requestStart + "&limit=" + limit
      );
      const data = await res.json();

      if (!res.ok) {
        showMessage(data.message || "Failed to load rows for export");
        return [];
      }

      const normalizedRows = normalizeData(data.rows || [], limit).slice(0, limit);

      rows.push(...normalizedRows);
      loadedRows += normalizedRows.length;
    }

    return rows.slice(0, expectedRows);
  };

  const clearSelectedCell = () => {
    if (!selectedCell) return;
    updateCell(selectedCell.rowIndex, selectedCell.colIndex, "");
  };

  const runCellSearch = async () => {
    const query = cellSearch.trim();

    if (!selectedSheet || !query) {
      setCellSearchMatches([]);
      setCellSearchIndex(-1);
      return [];
    }

    setCellSearchLoading(true);

    const res = await authFetch(
      API_URL + "/sheet/" + selectedSheet._id + "/search?q=" + encodeURIComponent(query)
    );
    const data = await res.json();
    setCellSearchLoading(false);

    if (!res.ok) {
      showMessage(data.message || "Search failed");
      return [];
    }

    const matches = data.matches || [];
    setCellSearchMatches(matches);
    setCellSearchIndex(-1);

    if (data.truncated) {
      showMessage("Showing first 500 matches");
    }

    return matches;
  };

  const centerRowInGrid = (rowIndex) => {
    const grid = gridRef.current;
    const viewportHeight = grid?.clientHeight || window.innerHeight;
    const scrollTop = Math.max(
      0,
      (rowIndex * DEFAULT_ROW_HEIGHT) - (viewportHeight / 2) + (DEFAULT_ROW_HEIGHT / 2)
    );

    setGridScrollTop(scrollTop);
    window.requestAnimationFrame(() => {
      gridRef.current?.scrollTo({ top: scrollTop, behavior: "smooth" });
    });
  };

  const goToPendingCodeRow = async () => {
    if (!pendingCodeRows.length) {
      showMessage("No confirmed rows waiting for code");
      return;
    }

    const currentPendingIndex = pendingCodeRows.indexOf(pendingCodeCursor);
    const rowIndex = unvisitedPendingCodeRows.length
      ? unvisitedPendingCodeRows[0]
      : pendingCodeRows[currentPendingIndex >= 0 ? (currentPendingIndex + 1) % pendingCodeRows.length : 0];

    await ensureRowsLoaded(rowIndex + 1);
    setVisibleRows((count) => Math.max(count, rowIndex + 1));
    setSelectedCell({ rowIndex, colIndex: TAX_ITEM_CODE_COLUMN_INDEX });
    setSelectedRange({
      start: { row: rowIndex, col: TAX_ITEM_CODE_COLUMN_INDEX },
      end: { row: rowIndex, col: TAX_ITEM_CODE_COLUMN_INDEX },
    });
    setVisitedPendingCodeRows((rows) => (
      rows.includes(rowIndex) ? rows : [...rows, rowIndex]
    ));
    setPendingCodeCursor(rowIndex);
    centerRowInGrid(rowIndex);
  };

  const goToCellSearchResult = async (direction = 1) => {
    if (!cellSearch.trim()) {
      showMessage("Type something to search");
      return;
    }

    const matches = cellSearchMatches.length > 0 ? cellSearchMatches : await runCellSearch();

    if (matches.length === 0) {
      setCellSearchIndex(-1);
      showMessage("No matching cells");
      return;
    }

    const nextIndex =
      cellSearchIndex < 0
        ? 0
        : (cellSearchIndex + direction + matches.length) % matches.length;
    const match = matches[nextIndex];

    setCellSearchIndex(nextIndex);
    await ensureRowsLoaded(match.rowIndex + 1);
    setSelectedCell(match);
    setSelectedRange({
      start: { row: match.rowIndex, col: match.colIndex },
      end: { row: match.rowIndex, col: match.colIndex },
    });
    setVisibleRows((count) => Math.max(count, match.rowIndex + 1));
  };

  const copySelection = async () => {
    if (!selectedCell || !selectedSheet) return;
    const cell = getSheetCell(selectedCell.rowIndex, selectedCell.colIndex);
    const text = cell.formula || cell.value || "";

    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const helper = document.createElement("textarea");
      helper.value = text;
      helper.style.position = "fixed";
      helper.style.opacity = "0";
      document.body.appendChild(helper);
      helper.select();
      document.execCommand("copy");
      document.body.removeChild(helper);
    }
  };

  const pasteToCell = async () => {
    if (!selectedCell || !canEdit) return;
    const text = await navigator.clipboard.readText();
    updateCell(selectedCell.rowIndex, selectedCell.colIndex, text);
  };

  const dragFillDown = () => {
    if (!selectedCell || !selectedSheet || !canEdit) return;

    const { rowIndex, colIndex } = selectedCell;
    const source = getSheetCell(rowIndex, colIndex);
    const patches = [];

    for (let r = rowIndex + 1; r < Math.min(rowIndex + 6, selectedSheet.data.length); r++) {
      if (!isRowLoaded(r)) continue;
      patches.push({
        rowIndex: r,
        colIndex,
        value: source.value,
        formula: source.formula || "",
      });
    }

    if (patches.length === 0) return;

    applyCellPatchesLocal(patches, {
      queueSave: false,
      recalculate: shouldRecalculateForPatches(patches),
    });
    queueCellSocketSave({
      sheetId: selectedSheet._id,
      rowIndex: patches[0].rowIndex,
      colIndex,
      value: patches[0].value,
      formula: patches[0].formula,
      patches,
    });
  };

  const saveVersion = () => {
    if (!selectedSheet) return;

    const version = {
      createdAt: new Date().toISOString(),
      data: selectedSheet.data,
      meta: selectedSheet.meta,
    };

    setSelectedSheet((prev) => ({
      ...prev,
      meta: {
        ...prev.meta,
        versions: [version, ...(prev.meta?.versions || [])].slice(0, 10),
      },
    }));

    queueMetadataSave();
    showMessage("Version saved");
  };

  const restoreVersion = (index) => {
    const version = selectedSheet?.meta?.versions?.[index];
    if (!version) return;

    setSelectedSheet((prev) => ({
      ...prev,
      data: normalizeData(version.data),
      meta: { ...prev.meta, ...(version.meta || {}) },
    }));

    queueMetadataSave();
  };

  const loadExcelTools = async () => import("xlsx");

  const loadPdfTools = async () => {
    const [{ default: jsPDF }, { default: autoTable }] = await Promise.all([
      import("jspdf"),
      import("jspdf-autotable"),
    ]);

    return { jsPDF, autoTable };
  };

  const exportCSV = async () => {
    if (!selectedSheet) return;
    const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id + "/export.csv");

    if (!res.ok) {
      showMessage("Failed to export sheet");
      return;
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = (selectedSheet.name || "sheet") + ".csv";
    link.click();
    URL.revokeObjectURL(url);
  };

  const exportExcel = async () => {
    if (!selectedSheet) return;

    const range = getExportRangeBounds();
    if (!range) return;

    const XLSX = await loadExcelTools();
    const exportRows = await loadRowsForExportRange(range);
    if (!exportRows.length) return;

    const rows = exportRows.map((row) =>
      row.slice(0, COLS).map((cell) => {
        const normalized = normalizeCell(cell);
        return normalized.formula || normalized.value || "";
      })
    );
    const worksheet = XLSX.utils.aoa_to_sheet(rows);

    exportRows.forEach((row, rowIndex) => {
      row.slice(0, COLS).forEach((cell, colIndex) => {
        const normalized = normalizeCell(cell);
        if (!normalized.formula) return;

        const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        worksheet[address] = {
          ...(worksheet[address] || {}),
          f: normalized.formula.replace(/^=/, ""),
          v: normalized.value || undefined,
        };
      });
    });

    worksheet["!cols"] = Array.from({ length: COLS }, (_, colIndex) => ({
      wch: Math.max(8, Math.round(getColumnWidth(colIndex) / 7)),
    }));

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    XLSX.writeFile(workbook, `${selectedSheet.name || "sheet"}-${range.label}.xlsx`);
  };

  const exportPDF = async () => {
    if (!selectedSheet) return;

    const range = getExportRangeBounds();
    if (!range) return;

    const { jsPDF, autoTable } = await loadPdfTools();
    const doc = new jsPDF({ orientation: "landscape" });
    const exportRows = await loadRowsForExportRange(range);
    if (!exportRows.length) return;

    const rows = exportRows.map((row) =>
      row.slice(0, COLS).map((cell) => normalizeCell(cell).value)
    );

    autoTable(doc, {
      head: [Array.from({ length: COLS }, (_, i) => colName(i))],
      body: rows,
    });

    doc.save(`${selectedSheet.name || "sheet"}-${range.label}.pdf`);
  };

  const exportTaxExcel = async () => {
    if (!selectedSheet) return;

    const dateBounds = getTaxExportDateBounds();
    if (!dateBounds) return;

    const params = new URLSearchParams(dateBounds);
    const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id + "/tax-export-rows?" + params.toString());
    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to load tax export rows");
      return;
    }

    const XLSX = await loadExcelTools();
    const activeFrom = getCairoDateDaysAgo(TAX_ACTIVE_FROM_DAYS_AGO);
    const taxRows = (data.rows || []).map(({ cells }) => {
      const itemCode = getCellExportText(cells, TAX_ITEM_CODE_COLUMN_INDEX);
      const itemName = getCellExportText(cells, TAX_ITEM_NAME_COLUMN_INDEX);

      return ["EGS", itemCode, itemName, itemName, itemName, itemName, activeFrom, null, null, null];
    });

    if (!taxRows.length) {
      showMessage("No new confirmed coded items found in selected period");
      return;
    }

    const worksheet = XLSX.utils.aoa_to_sheet([TAX_EXPORT_HEADERS, ...taxRows], { cellDates: true });
    worksheet["!cols"] = TAX_EXPORT_COLUMN_WIDTHS.map((width) => ({ width }));
    worksheet["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: taxRows.length, c: TAX_EXPORT_HEADERS.length - 1 },
    });

    taxRows.forEach((_, rowIndex) => {
      const sheetRowIndex = rowIndex + 1;
      const itemCodeAddress = XLSX.utils.encode_cell({ r: sheetRowIndex, c: 1 });
      const activeFromAddress = XLSX.utils.encode_cell({ r: sheetRowIndex, c: 6 });
      const activeToAddress = XLSX.utils.encode_cell({ r: sheetRowIndex, c: 7 });

      worksheet[itemCodeAddress] = {
        ...(worksheet[itemCodeAddress] || {}),
        t: "s",
        z: "@",
      };
      worksheet[activeFromAddress] = {
        t: "d",
        v: activeFrom,
        z: "m/d/yy",
      };
      worksheet[activeToAddress] = {
        t: "z",
        z: "m/d/yy",
      };
    });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "DataEntry");
    XLSX.writeFile(workbook, TAX_EXPORT_FILE_NAME, { cellDates: true });
  };

  const uploadExcel = async (event) => {
    const file = event.target.files?.[0];
    if (!file || !canBypassRowLocks || !selectedSheet) return;

    const buffer = await file.arrayBuffer();
    const XLSX = await loadExcelTools();
    const wb = XLSX.read(buffer);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const maxCols = rows.reduce((max, row) => Math.max(max, row.length), 0);

    if (maxCols > COLS) {
      const ok = window.confirm(
        `This file has more than ${COLS} columns. Extra columns will be ignored. Continue?`
      );

      if (!ok) {
        event.target.value = "";
        return;
      }
    }

    const importedData = normalizeData(
      rows.map((row) => row.map((value) => ({ value: value || "", style: defaultCellStyle })))
    );

    setSavingStatus("Importing...");

    for (let start = 0; start < importedData.length; start += IMPORT_BATCH_SIZE) {
      const batch = importedData.slice(start, start + IMPORT_BATCH_SIZE);
      const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id + "/import-rows", {
        method: "POST",
        body: JSON.stringify({
          start,
          rows: batch,
          reset: start === 0,
        }),
      });
      const data = await res.json();

      if (!res.ok) {
        showMessage(data.message || "Import failed");
        setSavingStatus("");
        event.target.value = "";
        return;
      }
    }

    const initialRows = importedData.slice(0, INITIAL_VISIBLE_ROWS);

    setSelectedSheet((prev) => ({ ...prev, data: initialRows }));
    setSheetRowTotal(normalizeRowTotal(importedData.length));
    setVisibleRows(Math.min(INITIAL_VISIBLE_ROWS, importedData.length));
    setGridScrollTop(0);
    setSavingStatus("Saved");
    showMessage("Import completed");
    event.target.value = "";
  };

  const saveSheet = async (silent = false) => {
    flushPendingCellSave();

    const sheet = selectedSheetRef.current;
    if (!sheet || !canEdit) return;

    setSavingStatus("Saving...");

    const res = await authFetch(API_URL + "/sheet/" + sheet._id, {
      method: "PUT",
      body: JSON.stringify({ meta: sheet.meta }),
    });

    const data = await res.json();

    if (!res.ok) {
      setSavingStatus("");
      showMessage(data.message || "Failed to save sheet");
      return;
    }

    const normalized = normalizeSheet(data.sheet, { minRows: 0 });

    setSelectedSheet((prev) => (
      prev
        ? { ...prev, meta: normalized.meta, analytics: normalized.analytics }
        : normalized
    ));
    setSavingStatus("Saved");
    if (!silent) showMessage("Sheet saved");
    await loadAnalytics();

    setTimeout(() => setSavingStatus(""), 1500);
  };

  const mergeSheetMetadata = (sheet) => {
    const normalized = normalizeSheet(sheet, { minRows: 0 });

    setSelectedSheet((prev) => (
      prev && prev._id === normalized._id
        ? {
            ...prev,
            ...normalized,
            data: normalized.data?.length ? normalized.data : prev.data,
            rowOwners: Object.keys(normalized.rowOwners || {}).length
              ? normalized.rowOwners
              : prev.rowOwners,
          }
        : normalized
    ));
  };

  const renameSheet = async () => {
    if (!selectedSheet || !canManage) return;
    const newName = prompt("New sheet name", selectedSheet.name);
    if (!newName) return;

    const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id + "/name", {
      method: "PATCH",
      body: JSON.stringify({ name: newName }),
    });

    const data = await res.json();
    if (res.ok) {
      mergeSheetMetadata(data.sheet);
      await loadSheets();
    }
  };

  const shareSheet = async () => {
    if (!selectedSheet || !canManageSheetUsers) return;

    const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id + "/share", {
      method: "POST",
      body: JSON.stringify({ identifier: shareEmail, role: shareRole }),
    });

    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to share sheet");
      return;
    }

    mergeSheetMetadata(data.sheet);
    setShareEmail("");
  };

  const updateCollaboratorRole = async (user, nextRole) => {
    if (!selectedSheet || !canManageSheetUsers) return;

    const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id + "/share", {
      method: "POST",
      body: JSON.stringify({ identifier: user.username || user.email, role: nextRole }),
    });
    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to update role");
      return;
    }

    mergeSheetMetadata(data.sheet);
    showMessage("User role updated");
  };

  const removeCollaborator = async (userId) => {
    if (!selectedSheet || !canManageSheetUsers) return;

    const res = await authFetch(
      API_URL + "/sheet/" + selectedSheet._id + "/collaborator/" + userId,
      { method: "DELETE" }
    );
    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to stop sharing");
      return;
    }

    mergeSheetMetadata(data.sheet);
    showMessage("Sharing stopped");
  };

  const stopAllSharing = async () => {
    if (!selectedSheet || !canManageSheetUsers) return;

    const collaborators = (selectedSheet.collaborators || []).filter(
      (user) => user.role !== "owner" && (canGrantSheetAdmin || user.role !== "admin")
    );

    if (collaborators.length === 0) {
      showMessage("No shared users to remove");
      return;
    }

    const ok = window.confirm("Stop sharing this sheet with all collaborators?");

    if (!ok) return;

    for (const user of collaborators) {
      const res = await authFetch(
        API_URL + "/sheet/" + selectedSheet._id + "/collaborator/" + user.userId,
        { method: "DELETE" }
      );

      if (!res.ok) {
        const data = await res.json();
        showMessage(data.message || "Failed to stop all sharing");
        return;
      }
    }

    await openSheet(selectedSheet._id);
    showMessage("All sharing stopped");
  };

  const logout = () => {
    flushPendingCellSave();
    sessionStorage.removeItem("token");
    setToken("");
    setCurrentUser(null);
    setConnectionStatus("offline");
    setEmail("");
    setPassword("");
    setUsername("");
    setCurrentPage("sheet");
    setProfileForm({ username: "", email: "", avatarUrl: "" });
    setPasswordForm({ currentPassword: "", newPassword: "" });
    setAuthFormKey((key) => key + 1);
    setSheets([]);
    setSelectedSheet(null);
    setSheetRowTotal(0);
    setServerPendingCodeRows([]);
    setVisitedPendingCodeRows([]);
    setPendingCodeCursor(null);
    setGridScrollTop(0);
    setMenuOpen(false);
  };

  const dashboardStats = useMemo(() => {
    return {
      sheets: sheets.length,
      collaborators: selectedSheet?.collaborators?.length || 0,
      versions: selectedSheet?.meta?.versions?.length || 0,
    };
  }, [sheets, selectedSheet]);

  const filteredSheets = useMemo(() => {
    const query = sheetSearch.trim().toLowerCase();
    if (!query) return sheets;

    return sheets.filter((sheet) => sheet.name.toLowerCase().includes(query));
  }, [sheetSearch, sheets]);

  const descriptionOptions = descriptionBuilder.open ? itemDescriptionOptions || {} : {};
  const descriptionTypes = Object.keys(descriptionOptions);
  const descriptionTypeOptions = descriptionOptions[descriptionBuilder.type] || EMPTY_DESCRIPTION_OPTIONS;
  const descriptionRows = descriptionBuilder.open ? descriptionTypeOptions.rows || [] : [];
  const descriptionCategoryOptions = descriptionBuilder.open
    ? uniqueDescriptionValues(descriptionRows, "category")
    : [];
  const descriptionFilteredRows = descriptionBuilder.open
    ? descriptionRows.filter((row) => {
        if (descriptionBuilder.category && row.category !== descriptionBuilder.category) return false;

        return Object.entries(descriptionBuilder.filters).every(([key, value]) => !value || row[key] === value);
      })
    : [];
  const selectedDescriptionRow =
    descriptionFilteredRows.find((row) => row.combination === descriptionBuilder.combination) ||
    (descriptionFilteredRows.length === 1 ? descriptionFilteredRows[0] : null);

  const getDescriptionFieldOptions = (fieldKey) => {
    const rows = descriptionRows.filter((row) => {
      if (descriptionBuilder.category && row.category !== descriptionBuilder.category) return false;

      return Object.entries(descriptionBuilder.filters).every(
        ([key, value]) => key === fieldKey || !value || row[key] === value
      );
    });

    return uniqueDescriptionValues(rows, fieldKey);
  };

  const pendingCodeRows = Array.from(new Set(serverPendingCodeRows))
    .map(Number)
    .filter(Number.isInteger)
    .sort((left, right) => left - right);
  const visitedPendingCodeSet = new Set(visitedPendingCodeRows);
  const pendingCodeSet = new Set(pendingCodeRows);
  const unvisitedPendingCodeRows = pendingCodeRows.filter(
    (rowIndex) => !visitedPendingCodeSet.has(rowIndex)
  );
  const pendingCodeBadgeCount = pendingCodeRows.length;

  const selectedCellStyle = selectedCell && selectedSheet
    ? normalizeCell(selectedSheet.data?.[selectedCell.rowIndex]?.[selectedCell.colIndex]).style
    : defaultCellStyle;
  const selectedCellCanEdit = selectedCell
    ? canEditCell(selectedCell.rowIndex, selectedCell.colIndex)
    : canEdit;

  const isBold = selectedCellStyle.fontWeight === "bold";
  const isItalic = selectedCellStyle.fontStyle === "italic";
  const isUnderline = selectedCellStyle.textDecoration === "underline";
  const canChangeSheetUserRole = (user) => {
    return user.role !== "owner" && (canGrantSheetAdmin || user.role !== "admin");
  };
  const visibleCellSearchIndex =
    cellSearchIndex >= 0 && cellSearchIndex < cellSearchMatches.length ? cellSearchIndex : -1;
  const virtualRowStart = Math.max(
    0,
    Math.floor(gridScrollTop / DEFAULT_ROW_HEIGHT) - VIRTUAL_ROW_BUFFER
  );
  const virtualRowEnd = Math.min(visibleRows, virtualRowStart + VIRTUAL_ROW_WINDOW);
  const renderedRows = Array.from(
    { length: Math.max(0, virtualRowEnd - virtualRowStart) },
    (_, offset) => selectedSheet?.data?.[virtualRowStart + offset] || createEmptyRow()
  );
  const topSpacerHeight = virtualRowStart * DEFAULT_ROW_HEIGHT;
  const bottomSpacerHeight = Math.max(0, visibleRows - virtualRowEnd) * DEFAULT_ROW_HEIGHT;

  useEffect(() => {
    if (!selectedSheet?._id || virtualRowEnd <= virtualRowStart) return undefined;

    const timer = window.setTimeout(() => {
      void ensureRowsLoaded(virtualRowEnd, { start: virtualRowStart });
    }, 0);

    return () => window.clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedSheet?._id, virtualRowStart, virtualRowEnd]);

  useEffect(() => {
    if (!token) return;

    const refreshData = async () => {
      if (document.visibilityState === "hidden") return;

      flushPendingCellSave();
      const activeSheetId = selectedSheetRef.current?._id;

      try {
        await Promise.all([
          loadSheets(),
          activeSheetId ? loadPendingCodeRows(activeSheetId) : Promise.resolve(),
        ]);
        setSavingStatus((current) => current === "Unsaved changes..." ? current : "Refreshed");
        window.setTimeout(() => {
          setSavingStatus((current) => current === "Refreshed" ? "" : current);
        }, 1200);
      } catch {
        setConnectionStatus((current) => current === "online" ? current : "offline");
      }
    };

    const refreshTimer = window.setInterval(() => {
      void refreshData();
    }, AUTO_REFRESH_INTERVAL_MS);

    return () => window.clearInterval(refreshTimer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [token, selectedWorkspaceId]);

  useEffect(() => {
    selectedSheetRef.current = selectedSheet;
  }, [selectedSheet]);

  useEffect(() => {
    if (!contextMenu || !contextMenuRef.current) return;

    const menuBox = contextMenuRef.current.getBoundingClientRect();
    const maxLeft = window.innerWidth - menuBox.width - CONTEXT_MENU_MARGIN;
    const maxTop = window.innerHeight - menuBox.height - CONTEXT_MENU_MARGIN;
    const nextX = clamp(contextMenu.x, CONTEXT_MENU_MARGIN, Math.max(CONTEXT_MENU_MARGIN, maxLeft));
    const nextY = clamp(contextMenu.y, CONTEXT_MENU_MARGIN, Math.max(CONTEXT_MENU_MARGIN, maxTop));

    if (nextX !== contextMenu.x || nextY !== contextMenu.y) {
      setContextMenu((current) => (current ? { ...current, x: nextX, y: nextY } : current));
    }
  }, [contextMenu]);

  useEffect(() => {
    if (!token) return;

    socketRef.current = io(API_URL, { auth: { token } });

    socketRef.current.on("connect", () => {
      setConnectionStatus("online");
      const activeSheetId = selectedSheetRef.current?._id;
      if (activeSheetId) {
        socketRef.current.emit("join-sheet", activeSheetId);
      }
    });

    socketRef.current.on("connect_error", () => {
      setConnectionStatus("offline");
    });

    socketRef.current.on("disconnect", () => {
      setConnectionStatus("offline");
    });

    socketRef.current.on("socket-error", (text) => {
      showMessage(text || "Realtime connection error");
    });

    socketRef.current.on("presence-updated", setOnlineUsers);

    socketRef.current.on("cell-change", ({ rowIndex, colIndex, value, formula, patches, rowOwners, pendingCodeUpdates }) => {
      const socketPatches = Array.isArray(patches) && patches.length > 0
        ? patches
        : [{ rowIndex, colIndex, value, formula }];

      if (Array.isArray(pendingCodeUpdates) && pendingCodeUpdates.length > 0) {
        applyPendingCodeUpdates(pendingCodeUpdates);
      } else if (touchesPendingCodeColumns(socketPatches)) {
        const activeSheetId = selectedSheetRef.current?._id;
        if (activeSheetId) void loadPendingCodeRows(activeSheetId);
      }

      setSelectedSheet((prev) => {
        if (!prev) return prev;
        const incomingPatches = socketPatches.filter((patch) => (
          !localCellDraftsRef.current.has(getCellPatchKey(prev._id, patch)) &&
          !patchMatchesCell(prev.data?.[patch.rowIndex]?.[patch.colIndex], patch)
        ));

        if (incomingPatches.length === 0) return prev;

        const data = [...prev.data];
        const touchedRows = new Set(incomingPatches.map((patch) => patch.rowIndex));

        touchedRows.forEach((patchRowIndex) => {
          data[patchRowIndex] = [...(data[patchRowIndex] || [])];
        });

        incomingPatches.forEach((patch) => {
          const target = normalizeCell(data[patch.rowIndex]?.[patch.colIndex]);
          const patchFormula = patch.formula || "";

          data[patch.rowIndex][patch.colIndex] = {
            ...target,
            value: patchFormula ? evaluateFormula(patchFormula, data) : patch.value,
            formula: patchFormula,
          };
        });

        const shouldRecalculate =
          prev.analytics?.totalFormulaCells > 0 ||
          incomingPatches.some((patch) => patch.formula || String(patch.value || "").startsWith("="));

        return {
          ...prev,
          data: shouldRecalculate ? recalculateData(data) : data,
          rowOwners: applyRowOwnerUpdates(prev.rowOwners, rowOwners),
        };
      });
    });

    socketRef.current.on("cell-style-change", ({ rowIndex, colIndex, style, rowOwners }) => {
      setSelectedSheet((prev) => {
        if (!prev) return prev;
        const data = [...prev.data];
        data[rowIndex] = [...(data[rowIndex] || [])];
        const target = normalizeCell(data[rowIndex]?.[colIndex]);
        data[rowIndex][colIndex] = {
          ...target,
          style: { ...target.style, ...style },
        };
        return {
          ...prev,
          data,
          rowOwners: applyRowOwnerUpdates(prev.rowOwners, rowOwners),
        };
      });
    });

    socketRef.current.on("sheet-saved", (sheet) => {
      setSelectedSheet((prev) => {
        const normalized = normalizeSheet(sheet, { minRows: 0 });

        if ((!normalized.data || normalized.data.length === 0) && prev?._id === normalized._id) {
          return {
            ...prev,
            meta: normalized.meta,
            analytics: normalized.analytics,
            erpOptions: normalized.erpOptions,
          };
        }

        return normalizeSheet(sheet);
      });
      setSavingStatus("Synced");
      setTimeout(() => setSavingStatus(""), 1200);
    });

    socketRef.current.on("sheet-deleted", ({ sheetId }) => {
      if (selectedSheetRef.current?._id === sheetId) {
        setSelectedSheet(null);
        setRole(null);
        showMessage("This sheet was deleted by owner");
        loadSheets();
      }
    });

    socketRef.current.on("erp-options-updated", (updatedOptions) => {
      setErpOptions({
        ...defaultErpOptions,
        ...(updatedOptions || {}),
      });
    });

    socketRef.current.on("collaborators-updated", (collaborators) => {
      setSelectedSheet((prev) => (prev ? { ...prev, collaborators } : prev));
    });

    return () => {
      flushPendingCellSave();
      socketRef.current.disconnect();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [token]);

  useEffect(() => {
    if (!token) return;

    const loadInitialData = async () => {
      await Promise.all([loadMe(), loadWorkspaces(), loadSheets(), loadAnalytics()]);
    };

    void loadInitialData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [token]);

  useEffect(() => {
    if (!token) return;

    const loadWorkspaceSheets = async () => {
      await loadSheets();
    };

    void loadWorkspaceSheets();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedWorkspaceId]);

  useEffect(() => {
    if (!selectedSheet || !canEdit || savingStatus !== "Unsaved changes...") return;
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => saveSheet(true), 1800);
    return () => clearTimeout(saveTimerRef.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedSheet?.meta, savingStatus, canEdit]);

  useEffect(() => {
    return () => {
      if (scrollFrameRef.current) {
        window.cancelAnimationFrame(scrollFrameRef.current);
      }
    };
  }, []);

  useEffect(() => {
    const flushOnPageHide = () => flushPendingCellSave();
    const flushOnVisibilityChange = () => {
      if (document.visibilityState === "hidden") flushPendingCellSave();
    };

    window.addEventListener("pagehide", flushOnPageHide);
    document.addEventListener("visibilitychange", flushOnVisibilityChange);

    return () => {
      window.removeEventListener("pagehide", flushOnPageHide);
      document.removeEventListener("visibilitychange", flushOnVisibilityChange);
    };
  });

  useEffect(() => {
    const isShortcut = (event, code) => {
      return (event.ctrlKey || event.metaKey) && event.code === code;
    };

    const isTypingTarget = (target) => {
      return ["INPUT", "TEXTAREA", "SELECT"].includes(target?.tagName);
    };

    const handler = (e) => {
      if (isShortcut(e, "KeyS")) {
        e.preventDefault();
        saveSheet(false);
      }
      if (isShortcut(e, "KeyC") && !isTypingTarget(e.target)) {
        e.preventDefault();
        copySelection();
      }
      if (isShortcut(e, "KeyV") && !isTypingTarget(e.target)) {
        e.preventDefault();
        pasteToCell();
      }
      if (e.key === "Delete" && selectedCell && !isTypingTarget(e.target)) {
        updateCell(selectedCell.rowIndex, selectedCell.colIndex, "");
      }
      if (isShortcut(e, "KeyB")) {
        e.preventDefault();
        updateCellStyle("fontWeight", "bold");
      }
    };

    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  });

  if (!token) {
    return (
      <div className="auth-page">
        <section className="auth-showcase">
          <div className="brand-mark">SS</div>
          <h1>Sheet SaaS</h1>
          <p>Collaborative ERP item sheets with controlled dropdowns, live edits, and clean exports.</p>
          <div className="auth-preview" aria-hidden="true">
            <div className="preview-toolbar">
              <span />
              <span />
              <span />
            </div>
            <div className="preview-grid">
              <strong>Item Name</strong>
              <strong>Main Group</strong>
              <strong>Unit</strong>
              <span>Pump Set</span>
              <span>Pumps</span>
              <span>PCS</span>
              <span>Pipe Work</span>
              <span>Fittings</span>
              <span>m</span>
              <span>Tank Base</span>
              <span>Metal works</span>
              <span>m2</span>
            </div>
          </div>
          <div className="auth-stats">
            <span>Realtime</span>
            <span>ERP Ready</span>
            <span>Secure Login</span>
          </div>
        </section>

        <div className="auth-card" key={authFormKey}>
          <div className="auth-card-heading">
            <span>{mode === "login" ? "Welcome back" : "Create workspace"}</span>
            <h2>{mode === "login" ? "Login to continue" : "Start with your account"}</h2>
            <p>{mode === "login" ? "Use your email and password to open your sheets." : "Add your username, email, and password."}</p>
          </div>

          {mode === "signup" && (
            <>
              <label>Username</label>
              <input
                autoComplete="off"
                name={`username-${authFormKey}`}
                value={username}
                onChange={(e) => setUsername(e.target.value)}
              />
            </>
          )}

          <label>Email</label>
          <input
            autoComplete="off"
            name={`email-${authFormKey}`}
            value={email}
            onChange={(e) => setEmail(e.target.value)}
          />

          <label>Password</label>
          <input
            autoComplete="new-password"
            name={`password-${authFormKey}`}
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
          />

          <button className="primary-btn" onClick={handleAuth}>
            {mode === "login" ? "Login" : "Signup"}
          </button>

          <button className="ghost-btn" onClick={() => setMode(mode === "login" ? "signup" : "login")}>
            {mode === "login" ? "Create account" : "Back to login"}
          </button>

          {message && <div className="message">{message}</div>}
        </div>
      </div>
    );
  }

  return (
    <div className="excel-app" onClick={() => setContextMenu(null)}>
      <header className="excel-header">
        <button className="burger" onClick={(e) => { e.stopPropagation(); setMenuOpen(true); }}>
          ☰
        </button>

        <div className="sheet-title">
          <strong>{selectedSheet?.name || "Sheet SaaS"}</strong>
          <span className={`status-pill ${connectionStatus}`}>{statusText}</span>
        </div>

        {selectedSheet && (
          <div className="excel-toolbar">
            <div className="toolbar-group toolbar-search">
              <input
                type="search"
                placeholder="Search cells"
                value={cellSearch}
                onChange={(e) => {
                  setCellSearch(e.target.value);
                  setCellSearchIndex(-1);
                  setCellSearchMatches([]);
                }}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    goToCellSearchResult(e.shiftKey ? -1 : 1);
                  }
                }}
              />
              <span>
                {cellSearchLoading
                  ? "..."
                  : cellSearch.trim()
                  ? `${cellSearchMatches.length ? visibleCellSearchIndex + 1 || 1 : 0}/${cellSearchMatches.length}`
                  : "0/0"}
              </span>
              <button title="Previous result" onClick={() => goToCellSearchResult(-1)}>
                Prev
              </button>
              <button title="Next result" onClick={() => goToCellSearchResult(1)}>
                Next
              </button>
            </div>

            <div className="toolbar-group pending-code-group">
              <button
                className={pendingCodeRows.length ? "pending-code-button active" : "pending-code-button"}
                disabled={!pendingCodeRows.length || rowsLoading}
                title="Go to confirmed rows waiting for item code"
                onClick={goToPendingCodeRow}
              >
                Needs Code
                {pendingCodeBadgeCount > 0 && (
                  <span className="pending-code-badge">{pendingCodeBadgeCount}</span>
                )}
              </button>
            </div>

            <div className="toolbar-group">
              <button
                className={isBold ? "tool-button active" : "tool-button"}
                disabled={!canEdit}
                title="Bold"
                onClick={() => updateCellStyle("fontWeight", isBold ? "normal" : "bold")}
              >
                B
              </button>
              <button
                className={isItalic ? "tool-button active" : "tool-button"}
                disabled={!canEdit}
                title="Italic"
                onClick={() => updateCellStyle("fontStyle", isItalic ? "normal" : "italic")}
              >
                I
              </button>
              <button
                className={isUnderline ? "tool-button active" : "tool-button"}
                disabled={!canEdit}
                title="Underline"
                onClick={() => updateCellStyle("textDecoration", isUnderline ? "none" : "underline")}
              >
                U
              </button>
            </div>

            <div className="toolbar-group toolbar-selects">
              <select
                value={selectedCellStyle.fontFamily}
                disabled={!canEdit}
                title="Font family"
                onChange={(e) => updateCellStyle("fontFamily", e.target.value)}
              >
                <option>Arial</option>
                <option>Calibri</option>
                <option>Tahoma</option>
                <option>Times New Roman</option>
                <option>Verdana</option>
              </select>

              <select
                value={selectedCellStyle.fontSize}
                disabled={!canEdit}
                title="Font size"
                onChange={(e) => updateCellStyle("fontSize", e.target.value)}
              >
                <option value="12px">12</option>
                <option value="14px">14</option>
                <option value="16px">16</option>
                <option value="18px">18</option>
                <option value="22px">22</option>
              </select>
            </div>

            <div className="toolbar-group color-tools">
              <label className="color-tool" title="Text color">
                <span>A</span>
                <input
                  type="color"
                  value={selectedCellStyle.color}
                  disabled={!canEdit}
                  onChange={(e) => updateCellStyle("color", e.target.value)}
                />
                <i style={{ backgroundColor: selectedCellStyle.color }} />
              </label>
              <label className="color-tool fill-tool" title="Fill color">
                <span>Fill</span>
                <input
                  type="color"
                  value={selectedCellStyle.backgroundColor}
                  disabled={!canEdit}
                  onChange={(e) => updateCellStyle("backgroundColor", e.target.value)}
                />
                <i style={{ backgroundColor: selectedCellStyle.backgroundColor }} />
              </label>
            </div>

            <div className="toolbar-group">
              <button
                className={selectedCellStyle.textAlign === "left" ? "tool-button active" : "tool-button"}
                disabled={!canEdit}
                title="Align left"
                onClick={() => updateCellStyle("textAlign", "left")}
              >
                L
              </button>
              <button
                className={selectedCellStyle.textAlign === "center" ? "tool-button active" : "tool-button"}
                disabled={!canEdit}
                title="Align center"
                onClick={() => updateCellStyle("textAlign", "center")}
              >
                C
              </button>
              <button
                className={selectedCellStyle.textAlign === "right" ? "tool-button active" : "tool-button"}
                disabled={!canEdit}
                title="Align right"
                onClick={() => updateCellStyle("textAlign", "right")}
              >
                R
              </button>
            </div>

            <div className="toolbar-group toolbar-actions">
              <button disabled={!canEdit} onClick={mergeCells}>Merge</button>
              <button disabled={!canEdit} onClick={unmergeCells}>Unmerge</button>
              <button disabled={!canEdit} onClick={dragFillDown}>Fill Down</button>
              <button className="save-tool" onClick={() => saveSheet(false)}>Save</button>
            </div>
          </div>
        )}
      </header>

      {menuOpen && (
        <aside className="left-drawer">
          <button className="close-menu" onClick={() => setMenuOpen(false)}>×</button>

          <h3>Menu</h3>
          <div className="drawer-profile">
            <div className="profile-avatar small">
              {currentUser?.avatarUrl ? (
                <img src={currentUser.avatarUrl} alt="" />
              ) : (
                <span>{(currentUser?.username || currentUser?.email || "U").charAt(0).toUpperCase()}</span>
              )}
            </div>
            <p>{currentUser?.username || currentUser?.email}</p>
          </div>

          <div className="drawer-section">
            <button onClick={() => { setCurrentPage("sheet"); setMenuOpen(false); }}>
              Sheets
            </button>
            <button onClick={() => { setCurrentPage("profile"); setMenuOpen(false); }}>
              Profile
            </button>
            {currentUser?.role === "admin" && (
              <button onClick={() => { setCurrentPage("admin"); setMenuOpen(false); }}>
                Admin
              </button>
            )}
          </div>

          <div className="drawer-section">
            <h4>Admin Analytics</h4>
            <div className="stat">Workspaces: {analytics?.totalWorkspaces || 0}</div>
            <div className="stat">Sheets: {analytics?.totalSheets || dashboardStats.sheets}</div>
            <div className="stat">Users: {analytics?.totalUsers || 0}</div>
            <div className="stat">Changes: {analytics?.totalChanges || 0}</div>
            <div className="stat">ERP Sheets: {analytics?.erpSheets || 0}</div>
          </div>

          <div className="drawer-section">
            <h4>Team Workspaces</h4>
            <input value={workspaceName} onChange={(e) => setWorkspaceName(e.target.value)} />
            <button onClick={createWorkspace}>Create Workspace</button>

            <select value={selectedWorkspaceId} onChange={(e) => setSelectedWorkspaceId(e.target.value)}>
              <option value="">Personal Sheets</option>
              {workspaces.map((w) => (
                <option key={w._id} value={w._id}>{w.name}</option>
              ))}
            </select>

            <input
              placeholder="Member email"
              value={workspaceMemberEmail}
              onChange={(e) => setWorkspaceMemberEmail(e.target.value)}
            />

            <select value={workspaceMemberRole} onChange={(e) => setWorkspaceMemberRole(e.target.value)}>
              <option value="admin">Admin</option>
              <option value="member">Member</option>
              <option value="viewer">Viewer</option>
            </select>

            <button onClick={addWorkspaceMember}>Add Workspace Member</button>
          </div>

          <div className="drawer-section">
            <h4>ERP Mini Sheet</h4>
            <input value={sheetName} onChange={(e) => setSheetName(e.target.value)} />

            <select value={erpType} onChange={(e) => setErpType(e.target.value)}>
              <option value="custom">Blank Sheet</option>
              <option value="item-master">ERP Item Master</option>
            </select>

            <button onClick={createSheet}>Create Sheet</button>
          </div>

          <div className="drawer-section">
            <h4>Sheets</h4>
            <input
              placeholder="Search sheets"
              value={sheetSearch}
              onChange={(e) => setSheetSearch(e.target.value)}
            />
            {filteredSheets.map((s) => (
              <button className="drawer-item" key={s._id} onClick={() => openSheet(s._id)}>
                {s.name}
              </button>
            ))}
          </div>

          {selectedSheet && canManageSheetUsers && (
            <div className="drawer-section">
              <h4>{canManage ? "Owner Controls" : "Admin Controls"}</h4>
              {canManage && (
                <>
                  <button className="danger-btn" onClick={() => setDeleteConfirmOpen(true)}>
                    Delete Sheet
                  </button>
                  <button onClick={() => setErpOptionsOpen(true)}>
                    ERP Options Manager
                  </button>
                </>
              )}

              <h4>Sheet Users</h4>
              {selectedSheet.collaborators?.some((user) => canChangeSheetUserRole(user)) && (
                <button className="danger-btn" onClick={stopAllSharing}>
                  Stop All Sharing
                </button>
              )}
              {selectedSheet.collaborators?.map((user) => (
                <div className="user-row" key={user.userId}>
                  <div>
                    <span>{user.username || user.email}</span>
                    {user.username && user.email && <small>{user.email}</small>}
                    {canChangeSheetUserRole(user) ? (
                      <select
                        value={user.role}
                        onChange={(e) => updateCollaboratorRole(user, e.target.value)}
                      >
                        <option value="viewer">Viewer</option>
                        <option value="editor">Editor</option>
                        {canGrantSheetAdmin && <option value="admin">Admin</option>}
                      </select>
                    ) : (
                      <small>{user.role}</small>
                    )}
                  </div>
                  {canChangeSheetUserRole(user) && (
                    <button
                      className="user-remove-btn"
                      onClick={() => removeCollaborator(user.userId)}
                    >
                      Stop Share
                    </button>
                  )}
                </div>
              ))}
            </div>
          )}

          {selectedSheet && canManageSheetUsers && (
            <div className="drawer-section">
              <h4>Share Sheet</h4>
              <input
                placeholder="username or email"
                value={shareEmail}
                onChange={(e) => setShareEmail(e.target.value)}
              />
              <select value={shareRole} onChange={(e) => setShareRole(e.target.value)}>
                <option value="viewer">Viewer</option>
                <option value="editor">Editor</option>
                {canGrantSheetAdmin && <option value="admin">Admin</option>}
              </select>
              <button onClick={shareSheet}>Share</button>
            </div>
          )}

          {selectedSheet && (
            <div className="drawer-section">
              <h4>File & History</h4>
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={uploadExcel} hidden />
              <button
                disabled={!canBypassRowLocks}
                onClick={() => fileInputRef.current.click()}
              >
                Upload Excel
              </button>
              <button onClick={exportCSV}>Export CSV</button>
              <div className="export-range-controls">
                <label>
                  <span>From row</span>
                  <input
                    type="number"
                    min="1"
                    max={sheetRowTotal || selectedSheet.data.length}
                    value={exportRange.from}
                    onChange={(event) =>
                      setExportRange((range) => ({ ...range, from: event.target.value }))
                    }
                  />
                </label>
                <label>
                  <span>To row</span>
                  <input
                    type="number"
                    min="1"
                    max={sheetRowTotal || selectedSheet.data.length}
                    placeholder={String(sheetRowTotal || selectedSheet.data.length)}
                    value={exportRange.to}
                    onChange={(event) =>
                      setExportRange((range) => ({ ...range, to: event.target.value }))
                    }
                  />
                </label>
              </div>
              <button onClick={exportExcel}>Export Excel</button>
              <button onClick={exportPDF}>Export PDF</button>
              <div className="tax-export-controls">
                <label>
                  <span>Tax from</span>
                  <input
                    type="datetime-local"
                    value={taxExportPeriod.from}
                    onChange={(event) =>
                      setTaxExportPeriod((period) => ({ ...period, from: event.target.value }))
                    }
                  />
                </label>
                <label>
                  <span>Tax to</span>
                  <input
                    type="datetime-local"
                    value={taxExportPeriod.to}
                    onChange={(event) =>
                      setTaxExportPeriod((period) => ({ ...period, to: event.target.value }))
                    }
                  />
                </label>
              </div>
              <button onClick={exportTaxExcel}>Export Taxes Excel</button>
              <button onClick={saveVersion}>Save Version</button>
              <button onClick={showAllChanges}>Show All Changes</button>
              <button onClick={renameSheet}>Rename Sheet</button>
            </div>
          )}

          {selectedSheet?.meta?.versions?.length > 0 && (
            <div className="drawer-section">
              <h4>Version History</h4>
              {selectedSheet.meta.versions.map((v, i) => (
                <button className="drawer-item" key={i} onClick={() => restoreVersion(i)}>
                  {new Date(v.createdAt).toLocaleString()}
                </button>
              ))}
            </div>
          )}

          <button className="logout-btn" onClick={logout}>Logout</button>
        </aside>
      )}

      {message && <div className="floating-message">{message}</div>}

      {changesPanelOpen && (
        <aside className="changes-panel">
          <button className="close-menu changes-close" onClick={() => setChangesPanelOpen(false)}>×</button>
          <h3>Cell Changes</h3>

          {cellChanges.length === 0 && <p>No changes found.</p>}

          {cellChanges.map((change) => (
            <div className="change-card" key={change._id}>
              <strong>{change.cellAddress}</strong>
              <span>{change.changeType}</span>
              <p>By: {change.userEmail}</p>
              <p>At: {new Date(change.createdAt).toLocaleString()}</p>
              <small>Old: {JSON.stringify(change.oldValue)}</small>
              <small>New: {JSON.stringify(change.newValue)}</small>
            </div>
          ))}
        </aside>
      )}

      {currentPage === "profile" ? (
        <main className="settings-page">
          <section className="settings-panel">
            <h2>Profile</h2>
            <div className="profile-image-row">
              <div className="profile-avatar">
                {profileForm.avatarUrl ? (
                  <img src={profileForm.avatarUrl} alt="" />
                ) : (
                  <span>{(profileForm.username || profileForm.email || "U").charAt(0).toUpperCase()}</span>
                )}
              </div>
              <div className="profile-image-actions">
                <label className="file-pick-btn">
                  Upload Photo
                  <input
                    type="file"
                    accept="image/png,image/jpeg,image/webp"
                    onChange={(e) => updateProfileImage(e.target.files?.[0])}
                  />
                </label>
                {profileForm.avatarUrl && (
                  <button
                    className="ghost-inline-btn"
                    onClick={() => setProfileForm((form) => ({ ...form, avatarUrl: "" }))}
                  >
                    Remove
                  </button>
                )}
              </div>
            </div>

            <label>Username</label>
            <input
              value={profileForm.username}
              onChange={(e) => setProfileForm((form) => ({ ...form, username: e.target.value }))}
            />

            <label>Email</label>
            <input
              value={profileForm.email}
              onChange={(e) => setProfileForm((form) => ({ ...form, email: e.target.value }))}
            />

            <button onClick={saveProfile}>Save Profile</button>
          </section>

          <section className="settings-panel">
            <h2>Password</h2>
            <label>Current Password</label>
            <input
              type="password"
              value={passwordForm.currentPassword}
              onChange={(e) => setPasswordForm((form) => ({ ...form, currentPassword: e.target.value }))}
            />

            <label>New Password</label>
            <input
              type="password"
              value={passwordForm.newPassword}
              onChange={(e) => setPasswordForm((form) => ({ ...form, newPassword: e.target.value }))}
            />

            <button onClick={changePassword}>Change Password</button>
          </section>
        </main>
      ) : currentPage === "admin" ? (
        <main className="settings-page">
          <section className="settings-panel admin-panel">
            <h2>Admin</h2>
            <div className="admin-grid">
              <div className="stat">Workspaces: {analytics?.totalWorkspaces || 0}</div>
              <div className="stat">Sheets: {analytics?.totalSheets || dashboardStats.sheets}</div>
              <div className="stat">Users: {analytics?.totalUsers || 0}</div>
              <div className="stat">Changes: {analytics?.totalChanges || 0}</div>
              <div className="stat">ERP Sheets: {analytics?.erpSheets || 0}</div>
            </div>
          </section>
        </main>
      ) : !selectedSheet ? (
        <main className="empty-excel">
          <section className="welcome-panel">
            <span className="welcome-kicker">Sheet SaaS</span>
            <h2>ERP item master workspace built for clean, collaborative data.</h2>
            <p>
              This app helps teams create and manage structured spreadsheet data with controlled
              dropdowns, live collaboration, import/export tools, sharing permissions, and change
              tracking in one focused workspace.
            </p>

            <div className="welcome-features">
              <div>
                <strong>Controlled Data</strong>
                <span>ERP dropdowns keep groups, units, confirmations, and item details consistent.</span>
              </div>
              <div>
                <strong>Team Workflow</strong>
                <span>Share sheets with collaborators and work together with realtime updates.</span>
              </div>
              <div>
                <strong>Export Ready</strong>
                <span>Import Excel files and export clean CSV, Excel, or PDF outputs when needed.</span>
              </div>
            </div>

            <div className="creator-card">
              <div>
                <small>Created by</small>
                <strong>Mohamed Helmy</strong>
              </div>
              <a
                href="https://www.linkedin.com/in/mohamed-helmy-94503b314?utm_source=share&utm_campaign=share_via&utm_content=profile&utm_medium=android_app"
                target="_blank"
                rel="noreferrer"
              >
                LinkedIn Profile
              </a>
            </div>
          </section>
        </main>
      ) : (
        <main className="sheet-fullscreen">
          <div className="formula-bar">
            <span>{selectedCell ? cellAddress(selectedCell.rowIndex, selectedCell.colIndex) : "A1"}</span>
            <input
              readOnly={!selectedCellCanEdit}
              title={selectedCell && !selectedCellCanEdit ? getCellLockMessage(selectedCell.rowIndex, selectedCell.colIndex) : ""}
              value={
                selectedCell
                  ? getSheetCell(selectedCell.rowIndex, selectedCell.colIndex).formula ||
                    getSheetCell(selectedCell.rowIndex, selectedCell.colIndex).value
                  : ""
              }
              onChange={(e) => {
                if (selectedCell) updateCell(selectedCell.rowIndex, selectedCell.colIndex, e.target.value);
              }}
            />
            <div className="online-users">
              {onlineUsers.map((u) => (
                <small key={u.socketId}>{u.username || u.email}</small>
              ))}
            </div>
          </div>

          <div className="grid-wrap-full" ref={gridRef} onScroll={handleGridScroll}>
            <table className="sheet-table-full">
              <colgroup>
                <col className="corner-col" />
                {Array.from({ length: COLS }, (_, colIndex) => (
                  <col key={colIndex} style={{ width: getColumnWidth(colIndex) }} />
                ))}
              </colgroup>

              <thead>
                <tr>
                  <th className="corner-cell"></th>
                  {Array.from({ length: COLS }, (_, colIndex) => (
                    <th
                      key={colIndex}
                      style={{
                        width: getColumnWidth(colIndex),
                        backgroundColor: rowPalette[0].backgroundColor,
                        color: rowPalette[0].color,
                      }}
                    >
                      {colName(colIndex)}
                    </th>
                  ))}
                </tr>
              </thead>

              <tbody>
                {topSpacerHeight > 0 && (
                  <tr className="virtual-spacer-row" style={{ height: topSpacerHeight }}>
                    <th></th>
                    <td colSpan={COLS}></td>
                  </tr>
                )}

                {renderedRows.map((row, rowOffset) => {
                  const rowIndex = virtualRowStart + rowOffset;
                  const rowHeight = getRowHeight(row, rowIndex);
                  return (
                    <tr
                      key={rowIndex}
                      className={[
                        pendingCodeSet.has(rowIndex) ? "pending-code-row" : "",
                        pendingCodeCursor === rowIndex ? "pending-code-row-focused" : "",
                      ].filter(Boolean).join(" ")}
                      style={{ height: rowHeight }}
                    >
                      <th
                        className={
                          selectedRange?.start.row === rowIndex &&
                          selectedRange?.end.row === rowIndex
                            ? "selected-row-header"
                            : ""
                        }
                        onMouseDown={() => selectRow(rowIndex)}
                        onDoubleClick={() => resizeRow(rowIndex)}
                        style={{
                          backgroundColor: getRowPalette(rowIndex).backgroundColor,
                          color: getRowPalette(rowIndex).color,
                        }}
                      >
                        {rowIndex + 1}
                      </th>

                    {row.map((cell, colIndex) => {
                      const mergeInfo = getMergeInfo(rowIndex, colIndex);
                      if (mergeInfo && !mergeInfo.isMaster) return null;

                      const normalizedCell = normalizeCell(cell);
                      const isSelected =
                        selectedCell?.rowIndex === rowIndex && selectedCell?.colIndex === colIndex;
                      const isInSelectedRange = isCellInSelectedRange(rowIndex, colIndex);
                      const isCodeColumn = colIndex === 12;
                      const cellCanEdit = canEditCell(rowIndex, colIndex);
                      const isDescriptionCell = colIndex === DESCRIPTION_BUILDER_COLUMN_INDEX;
                      const showDescriptionBuilderTrigger =
                        isDescriptionCell && !String(normalizedCell.value || normalizedCell.formula || "").trim();

                      const dropdownOptions = getDropdownOptions(rowIndex, colIndex);

                      return (
                        <td
                          key={rowIndex + "-" + colIndex}
                          rowSpan={mergeInfo ? mergeInfo.rowEnd - mergeInfo.rowStart + 1 : 1}
                          colSpan={mergeInfo ? mergeInfo.colEnd - mergeInfo.colStart + 1 : 1}
                          className={[
                            isInSelectedRange ? "selected-range-cell" : "",
                            isSelected ? "selected-cell" : "",
                            !cellCanEdit && colIndex <= ROW_LOCK_LAST_COLUMN_INDEX
                              ? "locked-sheet-cell"
                              : "",
                          ].filter(Boolean).join(" ")}
                          title={!cellCanEdit ? getCellLockMessage(rowIndex, colIndex) : ""}
                          onMouseDown={() => {
                            setSelectedCell({ rowIndex, colIndex });
                            setSelectedRange({
                              start: { row: rowIndex, col: colIndex },
                              end: { row: rowIndex, col: colIndex },
                            });
                          }}
                          onMouseEnter={(e) => {
                            if (e.buttons === 1 && selectedRange) {
                              setSelectedRange((prev) => ({
                                ...prev,
                                end: { row: rowIndex, col: colIndex },
                              }));
                            }
                          }}
                          onContextMenu={(e) => {
                            openCellContextMenu(e, rowIndex, colIndex);
                          }}
                        >
                          {dropdownOptions ? (
                            <>
                              <input
                                list={`options-${rowIndex}-${colIndex}`}
                                value={normalizedCell.value}
                                readOnly={!cellCanEdit}
                                style={{
                                  ...normalizedCell.style,
                                  width: "100%",
                                  minHeight: rowHeight,
                                  height: rowHeight,
                                  backgroundColor: getCellBackground(normalizedCell.style, rowIndex),
                                }}
                                onFocus={() => setSelectedCell({ rowIndex, colIndex })}
                                onContextMenu={(e) => {
                                  openCellContextMenu(e, rowIndex, colIndex);
                                }}
                                onChange={(e) => updateCell(rowIndex, colIndex, e.target.value)}
                                onBlur={() => validateErpOption(rowIndex, colIndex, dropdownOptions)}
                              />
                              <datalist id={`options-${rowIndex}-${colIndex}`}>
                                {dropdownOptions.map((option) => (
                                  <option key={option} value={option} />
                                ))}
                              </datalist>
                            </>
                          ) : (
                            <div className={isDescriptionCell ? "description-builder-cell" : ""}>
                              <textarea
                                value={normalizedCell.value}
                                readOnly={!cellCanEdit || (colIndex === 1 && rowIndex > 0)}
                                style={{
                                  ...normalizedCell.style,
                                  width: "100%",
                                  minHeight: rowHeight,
                                  height: rowHeight,
                                  lineHeight: `${
                                    isCodeColumn
                                      ? Math.max(getTextLineHeight(normalizedCell.style), rowHeight - CELL_VERTICAL_PADDING)
                                      : getTextLineHeight(normalizedCell.style)
                                  }px`,
                                  backgroundColor:
                                    colIndex === 1 && rowIndex > 0
                                      ? "#f8fafc"
                                      : getCellBackground(normalizedCell.style, rowIndex),
                                }}
                                onFocus={() => setSelectedCell({ rowIndex, colIndex })}
                                onContextMenu={(e) => {
                                  openCellContextMenu(e, rowIndex, colIndex);
                                }}
                                onChange={(e) => updateCell(rowIndex, colIndex, e.target.value)}
                              />
                              {showDescriptionBuilderTrigger && (
                                <button
                                  type="button"
                                  className="description-builder-trigger"
                                  disabled={!cellCanEdit}
                                  title="Build item description"
                                  onMouseDown={(event) => event.stopPropagation()}
                                  onClick={(event) => openDescriptionBuilder(event, rowIndex)}
                                >
                                  {itemDescriptionOptionsLoading ? (
                                    <small />
                                  ) : (
                                    <>
                                      <span />
                                      <span />
                                      <span />
                                    </>
                                  )}
                                </button>
                              )}
                            </div>
                          )}
                        </td>
                      );
                    })}
                    </tr>
                  );
                })}

                {bottomSpacerHeight > 0 && (
                  <tr className="virtual-spacer-row" style={{ height: bottomSpacerHeight }}>
                    <th></th>
                    <td colSpan={COLS}></td>
                  </tr>
                )}
              </tbody>
            </table>

            {visibleRows < sheetRowTotal && (
              <div className="load-rows-bar">
                <button disabled={rowsLoading} onClick={loadMoreRows}>
                  {rowsLoading
                    ? "Loading rows..."
                    : `Load more rows (${Math.min(visibleRows, sheetRowTotal)}/${sheetRowTotal})`}
                </button>
              </div>
            )}

          </div>

          {contextMenu && (
            <div
              ref={contextMenuRef}
              className="context-menu"
              style={{ top: contextMenu.y, left: contextMenu.x }}
              onClick={(e) => e.stopPropagation()}
            >
              <button onClick={() => runContextMenuAction(copySelection)}>Copy</button>
              <button disabled={!canEdit} onClick={() => runContextMenuAction(pasteToCell)}>
                Paste
              </button>
              <button disabled={!canEdit} onClick={() => runContextMenuAction(dragFillDown)}>
                Fill Down
              </button>
              <span className="context-menu-divider" />
              <button disabled={!canEdit} onClick={() => runContextMenuAction(mergeCells)}>
                Merge Cells
              </button>
              <button disabled={!canEdit} onClick={() => runContextMenuAction(unmergeCells)}>
                Unmerge Cells
              </button>
              <button disabled={!canEdit} onClick={() => runContextMenuAction(resetSelectedCellStyle)}>
                Clear Formatting
              </button>
              <button disabled={!canEdit} onClick={() => runContextMenuAction(clearSelectedCell)}>
                Clear Cell
              </button>
              <button
                disabled={!canEdit || !selectedCell}
                onClick={() => runContextMenuAction(() => resizeRow(selectedCell.rowIndex))}
              >
                Row Height
              </button>
              <span className="context-menu-divider" />
              <button onClick={() => runContextMenuAction(showCellChanges)}>Show Changes</button>
            </div>
          )}
        </main>
      )}

      {descriptionBuilder.open && (
        <div className="modal-backdrop">
          <div className="modal-box large-modal description-builder-modal">
            <h3>Item Description Builder</h3>

            <div className="description-builder-grid">
              <label>
                <span>Item Type</span>
                <select
                  value={descriptionBuilder.type}
                  onChange={(event) => updateDescriptionBuilderType(event.target.value)}
                >
                  {descriptionTypes.map((type) => (
                    <option key={type} value={type}>
                      {descriptionOptions[type].label || type}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                <span>Category</span>
                <select
                  value={descriptionBuilder.category}
                  onChange={(event) => updateDescriptionBuilderCategory(event.target.value)}
                >
                  <option value="">Select category</option>
                  {descriptionCategoryOptions.map((category) => (
                    <option key={category} value={category}>{category}</option>
                  ))}
                </select>
              </label>

              {descriptionTypeOptions.fields.map((field) => (
                <label key={field.key}>
                  <span>{field.label}</span>
                  <select
                    value={descriptionBuilder.filters[field.key] || ""}
                    disabled={!descriptionBuilder.category}
                    onChange={(event) => updateDescriptionBuilderFilter(field.key, event.target.value)}
                  >
                    <option value="">Select {field.label}</option>
                    {getDescriptionFieldOptions(field.key).map((option) => (
                      <option key={option} value={option}>{option}</option>
                    ))}
                  </select>
                </label>
              ))}
            </div>

            <label>Matching Combination</label>
            <select
              value={descriptionBuilder.combination}
              disabled={!descriptionBuilder.category || !descriptionFilteredRows.length}
              onChange={(event) =>
                setDescriptionBuilder((builder) => ({ ...builder, combination: event.target.value }))
              }
            >
              <option value="">Select combination</option>
              {descriptionFilteredRows.map((row) => (
                <option key={row.combination} value={row.combination}>
                  {row.combination}
                </option>
              ))}
            </select>

            <div className="description-builder-preview">
              {selectedDescriptionRow?.combination || "Select item specs to preview the final description"}
            </div>

            <div className="modal-actions">
              <button onClick={applyDescriptionBuilder} disabled={!selectedDescriptionRow}>
                Use Description
              </button>
              <button onClick={() => setDescriptionBuilder((builder) => ({ ...builder, open: false }))}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {deleteConfirmOpen && (
        <div className="modal-backdrop">
          <div className="modal-box">
            <h3>Delete Sheet</h3>
            <p>Are you sure you want to delete this sheet? This action cannot be undone.</p>
            <div className="modal-actions">
              <button className="danger-btn" onClick={deleteSheet}>Yes, Delete</button>
              <button onClick={() => setDeleteConfirmOpen(false)}>Cancel</button>
            </div>
          </div>
        </div>
      )}

      {erpOptionsOpen && (
        <div className="modal-backdrop">
          <div className="modal-box large-modal">
            <h3>ERP Options Manager</h3>

            <label>Option Type</label>
            <select
              value={optionTarget}
              onChange={(e) => {
                setOptionTarget(e.target.value);
                setOptionParent("");
              }}
            >
              <option value="mainGroups">Main Groups</option>
              <option value="subGroups">Sub Groups</option>
              <option value="subSubGroups">Sub Sub Groups</option>
              <option value="supportGroups">Support Groups</option>
              <option value="detailedGroups">Detailed Groups</option>
              <option value="units">Unit of Measure</option>
              <option value="shelfLife">Shelf Life</option>
              <option value="sequence">Sequence</option>
              <option value="confirmation1">First Confirmation</option>
              <option value="confirmation2">Second Confirmation</option>
            </select>

            {["subGroups", "subSubGroups", "supportGroups", "detailedGroups"].includes(optionTarget) && (
              <>
                <label>Parent Value</label>
                <input
                  list="erp-parent-options"
                  placeholder={
                    optionTarget === "supportGroups"
                      ? "Sub Group"
                      : "Parent option name exactly"
                  }
                  value={optionParent}
                  onChange={(e) => setOptionParent(e.target.value)}
                />
                <datalist id="erp-parent-options">
                  {getOptionParentSuggestions().map((parent) => (
                    <option key={parent} value={parent} />
                  ))}
                </datalist>
              </>
            )}

            {optionTarget === "supportGroups" && (
              <>
                <label>Copy Existing List</label>
                <select
                  value={supportSourceParent}
                  onChange={(e) => setSupportSourceParent(e.target.value)}
                >
                  <option value="">Select source list</option>
                  {Object.keys(erpOptions.supportGroups || {}).map((parent) => (
                    <option key={parent} value={parent}>{parent}</option>
                  ))}
                </select>
                <button onClick={copySupportOptionsToSubGroup}>Copy To Sub Group</button>
              </>
            )}

            <label>New Option</label>
            <input
              value={optionText}
              onChange={(e) => setOptionText(e.target.value)}
              placeholder="Write option name"
            />

            <button onClick={addErpOption}>Add Option</button>

            <div className="options-preview">
              <pre>{JSON.stringify(erpOptions, null, 2)}</pre>
            </div>

            <div className="modal-actions">
              <button onClick={() => setErpOptionsOpen(false)}>Close</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
