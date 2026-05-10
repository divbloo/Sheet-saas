import { useEffect, useMemo, useRef, useState } from "react";
import { io } from "socket.io-client";
import defaultErpOptions from "./defaultErpOptions.json";
import "./App.css";

const API_URL = import.meta.env.VITE_API_URL || "http://localhost:5000";

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

const DEFAULT_ROW_HEIGHT = 36;
const CELL_HORIZONTAL_PADDING = 16;
const CELL_VERTICAL_PADDING = 12;
const DEFAULT_LINE_HEIGHT = 20;
const COLUMN_WIDTH_PADDING = 28;
const AVERAGE_CHAR_WIDTH = 7.2;
const DEFAULT_COLUMN_WIDTHS = {
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
  10: 64,
  11: 220,
  12: 92,
  13: 150,
  14: 92,
  15: 92,
};
const AUTO_COLUMN_LIMITS = {
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
  10: { min: 54, max: 76 },
  11: { min: 110, max: 165 },
  12: { min: 72, max: 110 },
  13: { min: 90, max: 170 },
  14: { min: 72, max: 110 },
  15: { min: 72, max: 110 },
};

const defaultMeta = {
  colWidths: DEFAULT_COLUMN_WIDTHS,
  rowHeights: {},
  merges: [],
  versions: [],
};

const rowPalette = [
  {
    backgroundColor: "#e0f2fe",
    color: "#075985",
  },
  {
    backgroundColor: "#ccfbf1",
    color: "#115e59",
  },
];

const erpArabicHeaders = [
  "اسم الصنف",
  "الوصف",
  "المجموعة الرئيسية",
  "المجموعة الفرعية",
  "المجموعة تحت الفرعية",
  "المجموعة المساعدة",
  "المجموعة التفصيلية",
  "وحدة القياس",
  "العبوة",
  "الصلاحية",
  "التسلسل",
  "ملاحظات",
  "التأكيد الأول",
  "الكود",
  "التأكيد الثاني",
  "التأكيد الثالث",
];

const COLS = 16;
const MIN_SHEET_ROWS = 500;
const ADD_ROWS_STEP = 500;
const IMPORT_BATCH_SIZE = 500;
const INITIAL_VISIBLE_ROWS = 50;
const ROW_LOAD_STEP = 50;
const CONTEXT_MENU_MARGIN = 8;
const CELL_SAVE_DEBOUNCE_MS = 400;
const VIRTUAL_ROW_BUFFER = 12;
const VIRTUAL_ROW_WINDOW = 90;
const AUTO_FIT_SAMPLE_ROWS = 60;

const clamp = (value, min, max) => Math.min(max, Math.max(min, value));

function App() {
  const [mode, setMode] = useState("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [authFormKey, setAuthFormKey] = useState(0);
  const [token, setToken] = useState(sessionStorage.getItem("token") || "");
  const [currentUser, setCurrentUser] = useState(null);

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

  const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);
  const [erpOptionsOpen, setErpOptionsOpen] = useState(false);
  const [erpOptions, setErpOptions] = useState(defaultErpOptions);
  const [optionText, setOptionText] = useState("");
  const [optionTarget, setOptionTarget] = useState("mainGroups");
  const [optionParent, setOptionParent] = useState("");

  const socketRef = useRef(null);
  const selectedSheetRef = useRef(null);
  const saveTimerRef = useRef(null);
  const cellSaveTimerRef = useRef(null);
  const pendingCellSaveRef = useRef(null);
  const fileInputRef = useRef(null);
  const contextMenuRef = useRef(null);

  const canEdit = role === "owner" || role === "editor";
  const canManage = role === "owner";
  const statusText = savingStatus || (
    connectionStatus === "online"
      ? "Online"
      : connectionStatus === "connecting"
        ? "Connecting..."
        : "Offline"
  );

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

  const normalizeData = (data = [], minRows = MIN_SHEET_ROWS) => {
    return Array.from({ length: Math.max(minRows, data.length) }, (_, r) => {
      const row = data[r] || [];
      return Array.from({ length: COLS }, (_, c) =>
        normalizeCell(row[c] || "")
      );
    });
  };

  const normalizeSheet = (sheet, options = {}) => ({
    ...sheet,
    data: normalizeData(sheet.data, options.minRows ?? MIN_SHEET_ROWS),
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
    return erpArabicHeaders[index] || `Column ${index + 1}`;
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
    if (colIndex === 0 || colIndex === 1) {
      return getAutoColumnWidth(colIndex, selectedSheet?.data?.slice(0, visibleRows) || []);
    }

    return autoColumnWidths[colIndex] || getAutoColumnWidth(colIndex, selectedSheet?.data?.slice(0, AUTO_FIT_SAMPLE_ROWS) || []);
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
        body: JSON.stringify({ email, password }),
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
      }

      if (mode === "signup") setMode("login");
    } catch {
      showMessage("Cannot connect to backend");
    }
  };

  const loadMe = async () => {
    const res = await authFetch(API_URL + "/me");
    const data = await res.json();
    if (res.ok) setCurrentUser(data.user);
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

  const openSheet = async (id) => {
    flushPendingCellSave();

    if (socketRef.current && selectedSheet?._id) {
      socketRef.current.emit("leave-sheet", selectedSheet._id);
    }

    const res = await authFetch(API_URL + "/sheet/" + id + "?rowLimit=" + INITIAL_VISIBLE_ROWS);
    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to open sheet");
      return;
    }

    const normalized = normalizeSheet(data.sheet, { minRows: 0 });

    setSelectedSheet(normalized);
    setSheetRowTotal(data.rows?.total || normalized.data.length);
    setErpOptions(normalized.erpOptions || defaultErpOptions);
    setRole(data.role);
    setSelectedCell(null);
    setSelectedRange(null);
    setSheetRowTotal(0);
    setGridScrollTop(0);
    setVisibleRows(INITIAL_VISIBLE_ROWS);
    setMenuOpen(false);

    await loadErpOptions(id);

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
      "packages",
      "shelfLife",
      "sequence",
      "confirmation1",
      "confirmation2",
      "confirmation3",
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
    const subSubGroup = normalizeCell(row[4]).value;
    const supportGroup = normalizeCell(row[5]).value;

    if (colIndex === 2) return erpOptions.mainGroups || [];
    if (colIndex === 3) return erpOptions.subGroups?.[mainGroup] || [];
    if (colIndex === 4) return erpOptions.subSubGroups?.[subGroup] || [];
    if (colIndex === 5) return erpOptions.supportGroups?.[subSubGroup] || [];
    if (colIndex === 6) return erpOptions.detailedGroups?.[supportGroup] || [];
    if (colIndex === 7) return erpOptions.units || [];
    if (colIndex === 8) return erpOptions.packages || [];
    if (colIndex === 9) return erpOptions.shelfLife || [];
    if (colIndex === 10) return erpOptions.sequence || [];
    if (colIndex === 12) return erpOptions.confirmation1 || [];
    if (colIndex === 14) return erpOptions.confirmation2 || [];
    if (colIndex === 15) return erpOptions.confirmation3 || [];

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

  const handleSocketSaveResult = (result) => {
    if (!result?.ok) {
      setSavingStatus("Unsaved changes...");
      showMessage(result?.message || "Realtime save failed");
      return;
    }

    setSavingStatus("Saved");
    setTimeout(() => setSavingStatus(""), 1200);
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

    setSavingStatus("Saved");
    setTimeout(() => setSavingStatus(""), 1200);
    return true;
  };

  const flushPendingCellSave = () => {
    if (cellSaveTimerRef.current) {
      clearTimeout(cellSaveTimerRef.current);
      cellSaveTimerRef.current = null;
    }

    const payload = pendingCellSaveRef.current;
    pendingCellSaveRef.current = null;

    if (!payload) return;
    void saveCellPatchesHttp(payload);
  };

  const queueCellSocketSave = (payload) => {
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

  const applyCellPatchesLocal = (patches, options = {}) => {
    if (!canEdit) {
      showMessage("Viewer access: you cannot edit this sheet");
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

  const addRows = (count = ADD_ROWS_STEP) => {
    if (!selectedSheet || !canEdit) return;
    setRowsLoading(true);

    authFetch(API_URL + "/sheet/" + selectedSheet._id + "/rows", {
      method: "POST",
      body: JSON.stringify({ count }),
    })
      .then(async (res) => {
        const data = await res.json();

        if (!res.ok) {
          showMessage(data.message || "Failed to add rows");
          return;
        }

        const normalizedRows = normalizeData(data.rows || [], 0);

        setSelectedSheet((prev) => (
          prev
            ? {
                ...prev,
                data: [...prev.data, ...normalizedRows],
              }
            : prev
        ));
        setSheetRowTotal(data.total || 0);
        setVisibleRows((current) => current + normalizedRows.length);
        showMessage(`${normalizedRows.length} rows added`);
      })
      .catch(() => showMessage("Failed to add rows"))
      .finally(() => setRowsLoading(false));
  };

  const loadMoreRows = async () => {
    if (!selectedSheet || rowsLoading) return;

    const loadedRows = selectedSheet.data.length;

    if (visibleRows < loadedRows) {
      setVisibleRows((count) => Math.min(count + ROW_LOAD_STEP, loadedRows));
      return;
    }

    if (loadedRows >= sheetRowTotal) return;

    setRowsLoading(true);

    const res = await authFetch(
      API_URL + "/sheet/" + selectedSheet._id + "/rows?start=" + loadedRows + "&limit=" + ROW_LOAD_STEP
    );
    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to load rows");
      setRowsLoading(false);
      return;
    }

    const normalizedRows = normalizeData(data.rows || [], 0);

    setSelectedSheet((prev) => (
      prev
        ? {
            ...prev,
            data: [...prev.data, ...normalizedRows],
          }
        : prev
    ));
    setSheetRowTotal(data.total || sheetRowTotal);
    setVisibleRows((count) => Math.min(count + normalizedRows.length, loadedRows + normalizedRows.length));
    setRowsLoading(false);
  };

  const ensureRowsLoaded = async (targetCount) => {
    if (!selectedSheet || targetCount <= selectedSheet.data.length) return;

    setRowsLoading(true);

    let loadedRows = selectedSheet.data.length;
    const neededRows = Math.min(targetCount, sheetRowTotal);
    const loadedChunks = [];

    while (loadedRows < neededRows) {
      const limit = Math.min(ROW_LOAD_STEP, neededRows - loadedRows);
      const res = await authFetch(
        API_URL + "/sheet/" + selectedSheet._id + "/rows?start=" + loadedRows + "&limit=" + limit
      );
      const data = await res.json();

      if (!res.ok) {
        showMessage(data.message || "Failed to load rows");
        break;
      }

      const normalizedRows = normalizeData(data.rows || [], 0);
      if (normalizedRows.length === 0) break;

      loadedChunks.push(...normalizedRows);
      loadedRows += normalizedRows.length;
      setSheetRowTotal(data.total || sheetRowTotal);
    }

    if (loadedChunks.length > 0) {
      setSelectedSheet((prev) => (
        prev
          ? {
              ...prev,
              data: [...prev.data, ...loadedChunks],
            }
          : prev
      ));
    }

    setRowsLoading(false);
  };

  const loadAllRowsForExport = async () => {
    if (!selectedSheet) return [];

    const allRows = [...selectedSheet.data];
    let loadedRows = allRows.length;

    while (loadedRows < sheetRowTotal) {
      const limit = Math.min(ADD_ROWS_STEP, sheetRowTotal - loadedRows);
      const res = await authFetch(
        API_URL + "/sheet/" + selectedSheet._id + "/rows?start=" + loadedRows + "&limit=" + limit
      );
      const data = await res.json();

      if (!res.ok) {
        showMessage(data.message || "Failed to load rows for export");
        break;
      }

      const normalizedRows = normalizeData(data.rows || [], 0);
      if (normalizedRows.length === 0) break;

      allRows.push(...normalizedRows);
      loadedRows += normalizedRows.length;
    }

    if (allRows.length > selectedSheet.data.length) {
      setSelectedSheet((prev) => (prev ? { ...prev, data: allRows } : prev));
      setVisibleRows((count) => Math.max(count, Math.min(allRows.length, sheetRowTotal)));
    }

    return allRows;
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
    const cell = normalizeCell(selectedSheet.data[selectedCell.rowIndex][selectedCell.colIndex]);
    await navigator.clipboard.writeText(cell.formula || cell.value || "");
  };

  const pasteToCell = async () => {
    if (!selectedCell || !canEdit) return;
    const text = await navigator.clipboard.readText();
    updateCell(selectedCell.rowIndex, selectedCell.colIndex, text);
  };

  const dragFillDown = () => {
    if (!selectedCell || !selectedSheet || !canEdit) return;

    const { rowIndex, colIndex } = selectedCell;
    const source = normalizeCell(selectedSheet.data[rowIndex][colIndex]);
    const patches = [];

    for (let r = rowIndex + 1; r < Math.min(rowIndex + 6, selectedSheet.data.length); r++) {
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

  const exportExcel = async () => {
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

  const exportPDF = async () => {
    if (!selectedSheet) return;

    const { jsPDF, autoTable } = await loadPdfTools();
    const doc = new jsPDF({ orientation: "landscape" });
    const allRows = await loadAllRowsForExport();
    const exportRows = allRows.length > 0 ? allRows : selectedSheet.data;
    const rows = exportRows.map((row) =>
      row.slice(0, 16).map((cell) => normalizeCell(cell).value)
    );

    autoTable(doc, {
      head: [Array.from({ length: COLS }, (_, i) => colName(i))],
      body: rows,
    });

    doc.save((selectedSheet.name || "sheet") + ".pdf");
  };

  const uploadExcel = async (event) => {
    const file = event.target.files?.[0];
    if (!file || !canEdit || !selectedSheet) return;

    const buffer = await file.arrayBuffer();
    const XLSX = await loadExcelTools();
    const wb = XLSX.read(buffer);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const maxCols = rows.reduce((max, row) => Math.max(max, row.length), 0);

    if (maxCols > COLS) {
      const ok = window.confirm(
        "This file has more than 16 columns. Extra columns will be ignored. Continue?"
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
    setSheetRowTotal(importedData.length);
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

    const hasAllRowsLoaded = sheet.data.length >= sheetRowTotal;
    const payload = hasAllRowsLoaded
      ? { data: sheet.data, meta: sheet.meta }
      : { meta: sheet.meta };

    const res = await authFetch(API_URL + "/sheet/" + sheet._id, {
      method: "PUT",
      body: JSON.stringify(payload),
    });

    const data = await res.json();

    if (!res.ok) {
      setSavingStatus("");
      showMessage(data.message || "Failed to save sheet");
      return;
    }

    const normalized = normalizeSheet(data.sheet, {
      minRows: hasAllRowsLoaded ? MIN_SHEET_ROWS : 0,
    });

    setSelectedSheet((prev) => (
      hasAllRowsLoaded
        ? normalized
        : prev
          ? { ...prev, meta: normalized.meta, analytics: normalized.analytics }
          : normalized
    ));
    setSheetRowTotal(hasAllRowsLoaded ? normalized.data.length : sheetRowTotal);
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
    if (!selectedSheet || !canManage) return;

    const res = await authFetch(API_URL + "/sheet/" + selectedSheet._id + "/share", {
      method: "POST",
      body: JSON.stringify({ email: shareEmail, role: shareRole }),
    });

    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to share sheet");
      return;
    }

    mergeSheetMetadata(data.sheet);
    setShareEmail("");
  };

  const removeCollaborator = async (userId) => {
    if (!selectedSheet || !canManage) return;

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
    if (!selectedSheet || !canManage) return;

    const collaborators = (selectedSheet.collaborators || []).filter(
      (user) => user.role !== "owner"
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
    setAuthFormKey((key) => key + 1);
    setSheets([]);
    setSelectedSheet(null);
    setSheetRowTotal(0);
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

  const selectedCellStyle = selectedCell && selectedSheet
    ? normalizeCell(selectedSheet.data?.[selectedCell.rowIndex]?.[selectedCell.colIndex]).style
    : defaultCellStyle;

  const isBold = selectedCellStyle.fontWeight === "bold";
  const isItalic = selectedCellStyle.fontStyle === "italic";
  const isUnderline = selectedCellStyle.textDecoration === "underline";
  const visibleCellSearchIndex =
    cellSearchIndex >= 0 && cellSearchIndex < cellSearchMatches.length ? cellSearchIndex : -1;
  const virtualRowStart = Math.max(
    0,
    Math.floor(gridScrollTop / DEFAULT_ROW_HEIGHT) - VIRTUAL_ROW_BUFFER
  );
  const virtualRowEnd = Math.min(visibleRows, virtualRowStart + VIRTUAL_ROW_WINDOW);
  const renderedRows = selectedSheet?.data?.slice(virtualRowStart, virtualRowEnd) || [];
  const topSpacerHeight = virtualRowStart * DEFAULT_ROW_HEIGHT;
  const bottomSpacerHeight = Math.max(0, visibleRows - virtualRowEnd) * DEFAULT_ROW_HEIGHT;

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

    socketRef.current.on("cell-change", ({ rowIndex, colIndex, value, formula, patches }) => {
      setSelectedSheet((prev) => {
        if (!prev) return prev;
        const incomingPatches = Array.isArray(patches) && patches.length > 0
          ? patches
          : [{ rowIndex, colIndex, value, formula }];
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

        return { ...prev, data: shouldRecalculate ? recalculateData(data) : data };
      });
    });

    socketRef.current.on("cell-style-change", ({ rowIndex, colIndex, style }) => {
      setSelectedSheet((prev) => {
        if (!prev) return prev;
        const data = [...prev.data];
        data[rowIndex] = [...(data[rowIndex] || [])];
        const target = normalizeCell(data[rowIndex]?.[colIndex]);
        data[rowIndex][colIndex] = {
          ...target,
          style: { ...target.style, ...style },
        };
        return { ...prev, data };
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
    const handler = (e) => {
      if (e.ctrlKey && e.key.toLowerCase() === "s") {
        e.preventDefault();
        saveSheet(false);
      }
      if (e.ctrlKey && e.key.toLowerCase() === "c") {
        e.preventDefault();
        copySelection();
      }
      if (e.ctrlKey && e.key.toLowerCase() === "v") {
        e.preventDefault();
        pasteToCell();
      }
      if (e.key === "Delete" && selectedCell) {
        updateCell(selectedCell.rowIndex, selectedCell.colIndex, "");
      }
      if (e.ctrlKey && e.key.toLowerCase() === "b") {
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
        <div className="auth-card" key={authFormKey}>
          <h1>Sheet SaaS</h1>
          <p>ERP mini sheets with collaboration.</p>

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
          <p>{currentUser?.email}</p>

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

          {selectedSheet && canManage && (
            <div className="drawer-section">
              <h4>Owner Controls</h4>
              <button className="danger-btn" onClick={() => setDeleteConfirmOpen(true)}>
                Delete Sheet
              </button>
              <button onClick={() => setErpOptionsOpen(true)}>
                ERP Options Manager
              </button>

              <h4>Sheet Users</h4>
              {selectedSheet.collaborators?.some((user) => user.role !== "owner") && (
                <button className="danger-btn" onClick={stopAllSharing}>
                  Stop All Sharing
                </button>
              )}
              {selectedSheet.collaborators?.map((user) => (
                <div className="user-row" key={user.userId}>
                  <div>
                    <span>{user.email}</span>
                    <small>{user.role}</small>
                  </div>
                  {user.role !== "owner" && (
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

          {selectedSheet && canManage && (
            <div className="drawer-section">
              <h4>Share Sheet</h4>
              <input placeholder="email" value={shareEmail} onChange={(e) => setShareEmail(e.target.value)} />
              <select value={shareRole} onChange={(e) => setShareRole(e.target.value)}>
                <option value="viewer">Viewer</option>
                <option value="editor">Editor</option>
              </select>
              <button onClick={shareSheet}>Share</button>
            </div>
          )}

          {selectedSheet && (
            <div className="drawer-section">
              <h4>File & History</h4>
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={uploadExcel} hidden />
              <button onClick={() => fileInputRef.current.click()}>Upload Excel</button>
              <button onClick={exportExcel}>Export CSV</button>
              <button onClick={exportPDF}>Export PDF</button>
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

      {!selectedSheet ? (
        <main className="empty-excel">
          <h2>Open menu to create ERP Item Master sheet or select a sheet</h2>
        </main>
      ) : (
        <main className="sheet-fullscreen">
          <div className="formula-bar">
            <span>{selectedCell ? cellAddress(selectedCell.rowIndex, selectedCell.colIndex) : "A1"}</span>
            <input
              value={
                selectedCell
                  ? normalizeCell(selectedSheet.data[selectedCell.rowIndex][selectedCell.colIndex]).formula ||
                    normalizeCell(selectedSheet.data[selectedCell.rowIndex][selectedCell.colIndex]).value
                  : ""
              }
              onChange={(e) => {
                if (selectedCell) updateCell(selectedCell.rowIndex, selectedCell.colIndex, e.target.value);
              }}
            />
            <div className="online-users">
              {onlineUsers.map((u) => (
                <small key={u.socketId}>{u.email}</small>
              ))}
            </div>
          </div>

          <div className="grid-wrap-full" onScroll={(e) => setGridScrollTop(e.currentTarget.scrollTop)}>
            <table className="sheet-table-full">
              <colgroup>
                <col className="corner-col" />
                {selectedSheet.data[0]?.map((_, colIndex) => (
                  <col key={colIndex} style={{ width: getColumnWidth(colIndex) }} />
                ))}
              </colgroup>

              <thead>
                <tr>
                  <th className="corner-cell"></th>
                  {selectedSheet.data[0]?.map((_, colIndex) => (
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
                    <tr key={rowIndex} style={{ height: rowHeight }}>
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
                      const isCodeColumn = colIndex === 13;

                      const dropdownOptions = getDropdownOptions(rowIndex, colIndex);

                      return (
                        <td
                          key={rowIndex + "-" + colIndex}
                          rowSpan={mergeInfo ? mergeInfo.rowEnd - mergeInfo.rowStart + 1 : 1}
                          colSpan={mergeInfo ? mergeInfo.colEnd - mergeInfo.colStart + 1 : 1}
                          className={[
                            isInSelectedRange ? "selected-range-cell" : "",
                            isSelected ? "selected-cell" : "",
                          ].filter(Boolean).join(" ")}
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
                                disabled={!canEdit}
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
                            <textarea
                              value={normalizedCell.value}
                              disabled={!canEdit || (colIndex === 1 && rowIndex > 0)}
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

            {(visibleRows < selectedSheet.data.length || selectedSheet.data.length < sheetRowTotal) && (
              <div className="load-rows-bar">
                <button disabled={rowsLoading} onClick={loadMoreRows}>
                  {rowsLoading
                    ? "Loading rows..."
                    : `Load more rows (${Math.min(visibleRows, sheetRowTotal)}/${sheetRowTotal})`}
                </button>
              </div>
            )}

            {selectedSheet.data.length >= sheetRowTotal && (
              <div className="load-rows-bar">
                <button disabled={!canEdit || rowsLoading} onClick={() => addRows()}>
                  Add {ADD_ROWS_STEP} more rows ({sheetRowTotal} total)
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
            <select value={optionTarget} onChange={(e) => setOptionTarget(e.target.value)}>
              <option value="mainGroups">Main Groups</option>
              <option value="subGroups">Sub Groups</option>
              <option value="subSubGroups">Sub Sub Groups</option>
              <option value="supportGroups">Support Groups</option>
              <option value="detailedGroups">Detailed Groups</option>
              <option value="units">Unit of Measure</option>
              <option value="packages">Package Type</option>
              <option value="shelfLife">Shelf Life</option>
              <option value="sequence">Sequence</option>
              <option value="confirmation1">First Confirmation</option>
              <option value="confirmation2">Second Confirmation</option>
              <option value="confirmation3">Third Confirmation</option>
            </select>

            {["subGroups", "subSubGroups", "supportGroups", "detailedGroups"].includes(optionTarget) && (
              <>
                <label>Parent Value</label>
                <input
                  placeholder="Parent option name exactly"
                  value={optionParent}
                  onChange={(e) => setOptionParent(e.target.value)}
                />
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
