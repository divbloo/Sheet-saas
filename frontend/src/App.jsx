import { useEffect, useMemo, useRef, useState } from "react";
import { io } from "socket.io-client";
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
  textAlign: "left",
};

const defaultMeta = {
  colWidths: {
    0: 220,
    1: 280,
  },
  rowHeights: {},
  merges: [],
  versions: [],
};

const defaultErpOptions = {
  mainGroups: [],
  subGroups: {},
  subSubGroups: {},
  supportGroups: {},
  detailedGroups: {},
  units: [],
  packages: [],
  shelfLife: [],
  sequence: [],
  confirmation1: [],
  confirmation2: [],
  confirmation3: [],
};

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

const COLS = 20;
const ROWS = 500;

function App() {
  const [mode, setMode] = useState("login");
  const [email, setEmail] = useState("test@test.com");
  const [password, setPassword] = useState("123456");
  const [token, setToken] = useState(sessionStorage.getItem("token") || "");
  const [currentUser, setCurrentUser] = useState(null);

  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState(null);
  const [selectedCell, setSelectedCell] = useState(null);
  const [selectedRange, setSelectedRange] = useState(null);
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
  const fileInputRef = useRef(null);

  const canEdit = role === "owner" || role === "editor";
  const canManage = role === "owner";

  const normalizeCell = (cell) => {
    if (typeof cell === "object" && cell !== null && "value" in cell) {
      return {
        value: cell.value || "",
        formula: cell.formula || "",
        style: { ...defaultCellStyle, ...(cell.style || {}) },
      };
    }

    return {
      value: cell || "",
      formula: "",
      style: { ...defaultCellStyle },
    };
  };

  const normalizeData = (data = []) => {
    return Array.from({ length: Math.max(ROWS, data.length) }, (_, r) => {
      const row = data[r] || [];
      return Array.from({ length: Math.max(COLS, row.length) }, (_, c) =>
        normalizeCell(row[c] || "")
      );
    });
  };

  const normalizeSheet = (sheet) => ({
    ...sheet,
    data: normalizeData(sheet.data),
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

  const getColumnWidth = (colIndex) => {
    return selectedSheet?.meta?.colWidths?.[colIndex] || defaultMeta.colWidths[colIndex] || 120;
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
    if (socketRef.current && selectedSheet?._id) {
      socketRef.current.emit("leave-sheet", selectedSheet._id);
    }

    const res = await authFetch(API_URL + "/sheet/" + id);
    const data = await res.json();

    if (!res.ok) {
      showMessage(data.message || "Failed to open sheet");
      return;
    }

    const normalized = normalizeSheet(data.sheet);

    setSelectedSheet(normalized);
    setErpOptions(normalized.erpOptions || defaultErpOptions);
    setRole(data.role);
    setSelectedCell(null);
    setSelectedRange(null);
    setMenuOpen(false);

    await loadErpOptions(id);

    if (socketRef.current) {
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

  const updateSheetData = (updater, options = {}) => {
    if (!canEdit) {
      showMessage("Viewer access: you cannot edit this sheet");
      return;
    }

    setSelectedSheet((prev) => {
      if (!prev) return prev;

      const newData = prev.data.map((row) => row.map((cell) => ({ ...normalizeCell(cell) })));
      const updatedData = updater(newData);
      const recalculated = recalculateData(updatedData);

      return { ...prev, data: recalculated };
    });

    if (options.queueSave !== false) {
      setSavingStatus("Unsaved changes...");
    }
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

    updateSheetData((data) => {
      const oldCell = normalizeCell(data[rowIndex][colIndex]);

      data[rowIndex][colIndex] = {
        ...oldCell,
        value: String(inputValue).startsWith("=") ? evaluateFormula(inputValue, data) : inputValue,
        formula: String(inputValue).startsWith("=") ? inputValue : "",
      };

      if (colIndex === 0 && rowIndex > 0) {
        const descCell = normalizeCell(data[rowIndex][1]);
        data[rowIndex][1] = {
          ...descCell,
          value: inputValue,
          formula: "",
        };
      }

      if (colIndex === 2 && rowIndex > 0) {
        data[rowIndex][3] = { ...normalizeCell(data[rowIndex][3]), value: "" };
        data[rowIndex][4] = { ...normalizeCell(data[rowIndex][4]), value: "" };
        data[rowIndex][5] = { ...normalizeCell(data[rowIndex][5]), value: "" };
        data[rowIndex][6] = { ...normalizeCell(data[rowIndex][6]), value: "" };
      }

      if (colIndex === 3 && rowIndex > 0) {
        data[rowIndex][4] = { ...normalizeCell(data[rowIndex][4]), value: "" };
        data[rowIndex][5] = { ...normalizeCell(data[rowIndex][5]), value: "" };
        data[rowIndex][6] = { ...normalizeCell(data[rowIndex][6]), value: "" };
      }

      if (colIndex === 4 && rowIndex > 0) {
        data[rowIndex][5] = { ...normalizeCell(data[rowIndex][5]), value: "" };
        data[rowIndex][6] = { ...normalizeCell(data[rowIndex][6]), value: "" };
      }

      if (colIndex === 5 && rowIndex > 0) {
        data[rowIndex][6] = { ...normalizeCell(data[rowIndex][6]), value: "" };
      }

      return data;
    }, { queueSave: !socketRef.current });

    if (socketRef.current && selectedSheet) {
      setSavingStatus("Saving...");
      socketRef.current.emit("cell-change", {
        sheetId: selectedSheet._id,
        rowIndex,
        colIndex,
        value: patches[0].value,
        formula: patches[0].formula,
        patches,
      }, handleSocketSaveResult);
    }
  };

  const updateCellStyle = (styleKey, styleValue) => {
    if (!selectedCell || !selectedSheet || !canEdit) return;

    const { rowIndex, colIndex } = selectedCell;

    updateSheetData((data) => {
      const target = normalizeCell(data[rowIndex][colIndex]);
      data[rowIndex][colIndex] = {
        ...target,
        style: { ...target.style, [styleKey]: styleValue },
      };
      return data;
    }, { queueSave: !socketRef.current });

    if (socketRef.current) {
      setSavingStatus("Saving...");
      socketRef.current.emit("cell-style-change", {
        sheetId: selectedSheet._id,
        rowIndex,
        colIndex,
        style: { [styleKey]: styleValue },
      }, handleSocketSaveResult);
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

    setSavingStatus("Unsaved changes...");
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

    setSavingStatus("Unsaved changes...");
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

  const resizeCol = (colIndex) => {
    const width = prompt("Column width", getColumnWidth(colIndex));
    if (!width) return;

    setSelectedSheet((prev) => ({
      ...prev,
      meta: {
        ...prev.meta,
        colWidths: { ...(prev.meta?.colWidths || {}), [colIndex]: Number(width) },
      },
    }));

    setSavingStatus("Unsaved changes...");
  };

  const resizeRow = (rowIndex) => {
    const height = prompt("Row height", selectedSheet?.meta?.rowHeights?.[rowIndex] || 36);
    if (!height) return;

    setSelectedSheet((prev) => ({
      ...prev,
      meta: {
        ...prev.meta,
        rowHeights: { ...(prev.meta?.rowHeights || {}), [rowIndex]: Number(height) },
      },
    }));

    setSavingStatus("Unsaved changes...");
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
    if (!selectedCell || !canEdit) return;

    const { rowIndex, colIndex } = selectedCell;
    const source = normalizeCell(selectedSheet.data[rowIndex][colIndex]);

    updateSheetData((data) => {
      for (let r = rowIndex + 1; r < Math.min(rowIndex + 6, data.length); r++) {
        data[r][colIndex] = { ...source };
      }
      return data;
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

    setSavingStatus("Unsaved changes...");
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

    setSavingStatus("Unsaved changes...");
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
    const XLSX = await loadExcelTools();
    const plain = selectedSheet.data.map((row) => row.map((cell) => normalizeCell(cell).value));
    const ws = XLSX.utils.aoa_to_sheet(plain);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, selectedSheet.name || "Sheet");
    XLSX.writeFile(wb, (selectedSheet.name || "sheet") + ".xlsx");
  };

  const exportPDF = async () => {
    if (!selectedSheet) return;

    const { jsPDF, autoTable } = await loadPdfTools();
    const doc = new jsPDF({ orientation: "landscape" });
    const rows = selectedSheet.data.slice(0, 30).map((row) =>
      row.slice(0, 16).map((cell) => normalizeCell(cell).value)
    );

    autoTable(doc, {
      head: [selectedSheet.data[0].slice(0, 16).map((_, i) => colName(i))],
      body: rows,
    });

    doc.save((selectedSheet.name || "sheet") + ".pdf");
  };

  const uploadExcel = async (event) => {
    const file = event.target.files?.[0];
    if (!file || !canEdit) return;

    const buffer = await file.arrayBuffer();
    const XLSX = await loadExcelTools();
    const wb = XLSX.read(buffer);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const importedData = normalizeData(
      rows.map((row) => row.map((value) => ({ value: value || "", style: defaultCellStyle })))
    );

    setSelectedSheet((prev) => ({ ...prev, data: importedData }));
    setSavingStatus("Unsaved changes...");
  };

  const saveSheet = async (silent = false) => {
    const sheet = selectedSheetRef.current;
    if (!sheet || !canEdit) return;

    setSavingStatus("Saving...");

    const res = await authFetch(API_URL + "/sheet/" + sheet._id, {
      method: "PUT",
      body: JSON.stringify({ data: sheet.data, meta: sheet.meta }),
    });

    const data = await res.json();

    if (!res.ok) {
      setSavingStatus("");
      showMessage(data.message || "Failed to save sheet");
      return;
    }

    setSelectedSheet(normalizeSheet(data.sheet));
    setSavingStatus("Saved");
    if (!silent) showMessage("Sheet saved");
    await loadAnalytics();

    setTimeout(() => setSavingStatus(""), 1500);
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
      setSelectedSheet(normalizeSheet(data.sheet));
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

    setSelectedSheet(normalizeSheet(data.sheet));
    setShareEmail("");
  };

  const logout = () => {
    sessionStorage.removeItem("token");
    setToken("");
    setCurrentUser(null);
    setSheets([]);
    setSelectedSheet(null);
    setMenuOpen(false);
  };

  const dashboardStats = useMemo(() => {
    return {
      sheets: sheets.length,
      collaborators: selectedSheet?.collaborators?.length || 0,
      versions: selectedSheet?.meta?.versions?.length || 0,
    };
  }, [sheets, selectedSheet]);

  useEffect(() => {
    selectedSheetRef.current = selectedSheet;
  }, [selectedSheet]);

  useEffect(() => {
    if (!token) return;

    socketRef.current = io(API_URL, { auth: { token } });

    socketRef.current.on("presence-updated", setOnlineUsers);

    socketRef.current.on("cell-change", ({ rowIndex, colIndex, value, formula, patches }) => {
      setSelectedSheet((prev) => {
        if (!prev) return prev;
        const data = prev.data.map((row) => row.map((cell) => ({ ...normalizeCell(cell) })));
        const incomingPatches = Array.isArray(patches) && patches.length > 0
          ? patches
          : [{ rowIndex, colIndex, value, formula }];

        incomingPatches.forEach((patch) => {
          const target = normalizeCell(data[patch.rowIndex]?.[patch.colIndex]);
          const patchFormula = patch.formula || "";

          data[patch.rowIndex][patch.colIndex] = {
            ...target,
            value: patchFormula ? evaluateFormula(patchFormula, data) : patch.value,
            formula: patchFormula,
          };
        });

        return { ...prev, data: recalculateData(data) };
      });
    });

    socketRef.current.on("cell-style-change", ({ rowIndex, colIndex, style }) => {
      setSelectedSheet((prev) => {
        if (!prev) return prev;
        const data = prev.data.map((row) => row.map((cell) => ({ ...normalizeCell(cell) })));
        data[rowIndex][colIndex] = {
          ...data[rowIndex][colIndex],
          style: { ...data[rowIndex][colIndex].style, ...style },
        };
        return { ...prev, data };
      });
    });

    socketRef.current.on("sheet-saved", (sheet) => {
      setSelectedSheet(normalizeSheet(sheet));
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

    return () => socketRef.current.disconnect();
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
  }, [selectedSheet, savingStatus, canEdit]);

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
        <div className="auth-card">
          <h1>Sheet SaaS</h1>
          <p>ERP mini sheets with collaboration.</p>

          <label>Email</label>
          <input value={email} onChange={(e) => setEmail(e.target.value)} />

          <label>Password</label>
          <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} />

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
          <span>{savingStatus || "Ready"}</span>
        </div>

        {selectedSheet && (
          <div className="excel-toolbar">
            <button onClick={() => updateCellStyle("fontWeight", "bold")}>B</button>
            <button onClick={() => updateCellStyle("fontStyle", "italic")}>I</button>
            <button onClick={() => updateCellStyle("textDecoration", "underline")}>U</button>

            <select onChange={(e) => updateCellStyle("fontFamily", e.target.value)}>
              <option>Arial</option>
              <option>Calibri</option>
              <option>Tahoma</option>
              <option>Times New Roman</option>
              <option>Verdana</option>
            </select>

            <select onChange={(e) => updateCellStyle("fontSize", e.target.value)}>
              <option value="12px">12</option>
              <option value="14px">14</option>
              <option value="16px">16</option>
              <option value="18px">18</option>
              <option value="22px">22</option>
            </select>

            <input type="color" onChange={(e) => updateCellStyle("color", e.target.value)} />
            <input type="color" onChange={(e) => updateCellStyle("backgroundColor", e.target.value)} />

            <button onClick={() => updateCellStyle("textAlign", "left")}>Left</button>
            <button onClick={() => updateCellStyle("textAlign", "center")}>Center</button>
            <button onClick={() => updateCellStyle("textAlign", "right")}>Right</button>

            <button onClick={mergeCells}>Merge</button>
            <button onClick={unmergeCells}>Unmerge</button>
            <button onClick={dragFillDown}>Fill Down</button>
            <button onClick={() => saveSheet(false)}>Save</button>
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
              <option value="inventory">Inventory</option>
              <option value="sales">Sales</option>
              <option value="finance">Finance</option>
              <option value="purchasing">Purchasing</option>
              <option value="hr">HR</option>
              <option value="crm">CRM</option>
              <option value="offers">Offers</option>
              <option value="warehouse">Warehouse</option>
            </select>

            <button onClick={createSheet}>Create Sheet</button>
          </div>

          <div className="drawer-section">
            <h4>Sheets</h4>
            {sheets.map((s) => (
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
              {selectedSheet.collaborators?.map((user) => (
                <div className="user-row" key={user.userId}>
                  <span>{user.email}</span>
                  <small>{user.role}</small>
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
              <button onClick={exportExcel}>Export Excel</button>
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
          <button className="close-menu" onClick={() => setChangesPanelOpen(false)}>×</button>
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

          <div className="grid-wrap-full">
            <table className="sheet-table-full">
              <thead>
                <tr>
                  <th className="corner-cell"></th>
                  {selectedSheet.data[0]?.map((_, colIndex) => (
                    <th
                      key={colIndex}
                      onDoubleClick={() => resizeCol(colIndex)}
                      style={{ width: getColumnWidth(colIndex) }}
                    >
                      {colName(colIndex)}
                    </th>
                  ))}
                </tr>
              </thead>

              <tbody>
                {selectedSheet.data.map((row, rowIndex) => (
                  <tr
                    key={rowIndex}
                    style={{ height: selectedSheet.meta?.rowHeights?.[rowIndex] || 36 }}
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
                            e.preventDefault();
                            setSelectedCell({ rowIndex, colIndex });
                            setContextMenu({ x: e.clientX, y: e.clientY });
                          }}
                        >
                          {dropdownOptions ? (
                            <select
                              value={normalizedCell.value}
                              disabled={!canEdit}
                              style={{
                                ...normalizedCell.style,
                                width: getColumnWidth(colIndex),
                                height: selectedSheet.meta?.rowHeights?.[rowIndex] || 36,
                              }}
                              onFocus={() => setSelectedCell({ rowIndex, colIndex })}
                              onContextMenu={(e) => {
                                e.preventDefault();
                                e.stopPropagation();
                                setSelectedCell({ rowIndex, colIndex });
                                setContextMenu({ x: e.clientX, y: e.clientY });
                              }}
                              onChange={(e) => updateCell(rowIndex, colIndex, e.target.value)}
                            >
                              <option value=""></option>
                              {dropdownOptions.map((option) => (
                                <option key={option} value={option}>
                                  {option}
                                </option>
                              ))}
                            </select>
                          ) : (
                            <input
                              value={normalizedCell.value}
                              disabled={!canEdit || (colIndex === 1 && rowIndex > 0)}
                              style={{
                                ...normalizedCell.style,
                                width: getColumnWidth(colIndex),
                                height: selectedSheet.meta?.rowHeights?.[rowIndex] || 36,
                                backgroundColor:
                                  colIndex === 1 && rowIndex > 0
                                    ? "#f8fafc"
                                    : normalizedCell.style.backgroundColor,
                              }}
                              onFocus={() => setSelectedCell({ rowIndex, colIndex })}
                              onContextMenu={(e) => {
                                e.preventDefault();
                                e.stopPropagation();
                                setSelectedCell({ rowIndex, colIndex });
                                setContextMenu({ x: e.clientX, y: e.clientY });
                              }}
                              onChange={(e) => updateCell(rowIndex, colIndex, e.target.value)}
                            />
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {contextMenu && (
            <div
              className="context-menu"
              style={{ top: contextMenu.y, left: contextMenu.x }}
              onClick={(e) => e.stopPropagation()}
            >
              <button onClick={copySelection}>Copy</button>
              <button onClick={pasteToCell}>Paste</button>
              <button onClick={mergeCells}>Merge Cells</button>
              <button onClick={unmergeCells}>Unmerge</button>
              <button onClick={dragFillDown}>Fill Down</button>
              <button onClick={showCellChanges}>Show Changes</button>
              <button onClick={() => updateCell(selectedCell.rowIndex, selectedCell.colIndex, "")}>
                Clear Cell
              </button>
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
