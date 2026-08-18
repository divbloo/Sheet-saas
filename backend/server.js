require("dotenv").config();

const express = require("express");
const fs = require("fs");
const http = require("http");
const path = require("path");
const mongoose = require("mongoose");
const cors = require("cors");
const jwt = require("jsonwebtoken");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");
const { Server } = require("socket.io");

const User = require("./models/User");
const Sheet = require("./models/Sheet");
const SheetRow = require("./models/SheetRow");
const Workspace = require("./models/Workspace");
const ChangeLog = require("./models/ChangeLog");
const authRoutes = require("./routes/authRoutes");
const itemCodingRoutes = require("./routes/itemCodingRoutes");
const profileRoutes = require("./routes/profileRoutes");
const workspaceRoutes = require("./routes/workspaceRoutes");
const {
  createCorsOptions,
  createHelmetOptions,
  getFrontendUrls,
  verifyProductionSecurity,
} = require("./config/security");
const {
  ALLOWED_SHEET_TYPES,
  DEFAULT_ROW_PAGE_SIZE,
  DEFAULT_SHEET_ROWS,
  FIRST_CONFIRMATION_COLUMN_INDEX,
  MAX_ROW_PAGE_SIZE,
  ROW_LOCK_LAST_COLUMN_INDEX,
  TAX_DETAILED_GROUP_COLUMN_INDEX,
  TAX_ITEM_CODE_COLUMN_INDEX,
  TAX_ITEM_NAME_COLUMN_INDEX,
  TAX_MAIN_GROUP_COLUMN_INDEX,
  TAX_SUB_GROUP_COLUMN_INDEX,
  TAX_SUB_SUB_GROUP_COLUMN_INDEX,
  TAX_SUPPORT_GROUP_COLUMN_INDEX,
} = require("./config/sheetConstants");
const { isValidCellIndex, isValidObjectId } = require("./utils/validation");
const {
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
} = require("./utils/sheetHelpers");
const {
  DEFAULT_SHEET_COLS,
  buildRowSearchText,
  buildRowSearchTokens,
  buildSearchQueryTokens,
  defaultCellStyle,
  normalizeRowCells,
} = require("./utils/sheetRows");
const { createERPTemplateData } = require("./utils/erpTemplates");
const { auth } = require("./utils/userHelpers");
const {
  canAssignSheetRole,
  canBypassRowLocks,
  canEdit,
  canManage,
  canManageSheetUsers,
  canRead,
  canRemoveSheetCollaborator,
  canUseWorkspace,
  getUserRole,
  getWorkspaceRole,
} = require("./utils/permissions");

const app = express();
const server = http.createServer(app);

const PORT = process.env.PORT || 5000;
const FRONTEND_URLS = getFrontendUrls();
const corsOptions = createCorsOptions(FRONTEND_URLS);

const getCellText = (cells = [], colIndex) => {
  const cell = cells[colIndex] || {};
  if (cell && typeof cell === "object") return String(cell.formula || cell.value || "").trim();
  return String(cell || "").trim();
};

const rowNeedsCode = (cells = []) => {
  const normalizedCells = normalizeRowCells(cells);
  const firstConfirmation = getCellText(normalizedCells, FIRST_CONFIRMATION_COLUMN_INDEX);
  const itemCode = getCellText(normalizedCells, TAX_ITEM_CODE_COLUMN_INDEX);

  return Boolean(firstConfirmation && !itemCode);
};

const io = new Server(server, {
  cors: {
    origin: FRONTEND_URLS,
    methods: ["GET", "POST", "PUT", "PATCH", "DELETE"],
  },
});

app.use(helmet(createHelmetOptions()));
app.use(cors(corsOptions));
app.use(express.json({ limit: "15mb" }));

app.use(
  rateLimit({
    windowMs: 15 * 60 * 1000,
    max: 500,
  })
);

const perfLogThresholdMs = Number(process.env.PERF_LOG_THRESHOLD_MS || 250);
const perfLogPaths = [
  /^\/sheet\/[^/]+$/,
  /^\/sheet\/[^/]+\/rows$/,
  /^\/sheet\/[^/]+\/search$/,
  /^\/sheet\/[^/]+\/pending-code-rows$/,
  /^\/sheet\/[^/]+\/tax-export-rows$/,
  /^\/item-coding-options$/,
];

app.use((req, res, next) => {
  const start = process.hrtime.bigint();

  res.on("finish", () => {
    if (!perfLogPaths.some((pattern) => pattern.test(req.path))) return;

    const elapsedMs = Number(process.hrtime.bigint() - start) / 1e6;
    if (elapsedMs >= perfLogThresholdMs) {
      console.log(`[perf] ${req.method} ${req.originalUrl} ${res.statusCode} ${elapsedMs.toFixed(1)}ms`);
    }
  });

  next();
});

const rejectInvalidId = (res, name) => {
  return res.status(400).json({ message: `Invalid ${name}` });
};

app.param("id", (req, res, next, id) => {
  if (!isValidObjectId(id)) {
    return rejectInvalidId(res, "id");
  }

  next();
});

app.param("userId", (req, res, next, userId) => {
  if (!isValidObjectId(userId)) {
    return rejectInvalidId(res, "user id");
  }

  next();
});

if (!process.env.MONGO_URI) {
  console.error("Missing MONGO_URI in .env file");
  process.exit(1);
}

if (!process.env.JWT_SECRET) {
  console.error("Missing JWT_SECRET in .env file");
  process.exit(1);
}

verifyProductionSecurity();

const connectToDatabase = async () => {
  try {
    await mongoose.connect(process.env.MONGO_URI, {
      serverSelectionTimeoutMS: 15000,
    });
    console.log("DB Connected");
  } catch (error) {
    console.error("DB Connection Error:", error);
    process.exit(1);
  }
};

const isTransactionUnsupportedError = (error) => {
  const message = String(error?.message || "");

  return (
    error?.code === 20 ||
    error?.codeName === "IllegalOperation" ||
    message.includes("Transaction numbers are only allowed") ||
    message.includes("transactions are not supported")
  );
};

const runWithOptionalTransaction = async (operation) => {
  const session = await mongoose.startSession();

  try {
    let result;

    await session.withTransaction(async () => {
      result = await operation(session);
    });

    return result;
  } catch (error) {
    if (!isTransactionUnsupportedError(error)) {
      throw error;
    }

    return operation(null);
  } finally {
    await session.endSession();
  }
};

const rowHasContent = (row = []) => normalizeRowCells(row)
  .some((cell) => String(cell.value ?? "").trim() || String(cell.formula || "").trim());

const rowHasProtectedContent = (row = []) => normalizeRowCells(row)
  .slice(0, ROW_LOCK_LAST_COLUMN_INDEX + 1)
  .some((cell) => String(cell.value ?? "").trim() || String(cell.formula || "").trim());

const buildRowSearchFields = (cells = []) => ({
  searchText: buildRowSearchText(cells),
  searchTokens: buildRowSearchTokens(cells),
});

const updateSheetAnalytics = async (sheet, userEmail, options = {}) => {
  const lastEditedAt = options.lastEditedAt || new Date();
  const totalMergedCells = options.totalMergedCells ?? sheet.meta?.merges?.length ?? 0;
  const queryOptions = options.session ? { session: options.session } : undefined;

  if (Number.isFinite(options.totalFormulaCells)) {
    await Sheet.updateOne(
      { _id: sheet._id },
      {
        $inc: { "analytics.totalEdits": 1 },
        $set: {
          "analytics.totalFormulaCells": Math.max(0, options.totalFormulaCells),
          "analytics.totalMergedCells": totalMergedCells,
          "analytics.lastEditedBy": userEmail,
          "analytics.lastEditedAt": lastEditedAt,
        },
      },
      queryOptions
    );
    return;
  }

  const formulaCountDelta = Number(options.formulaCountDelta || 0);

  await Sheet.updateOne(
    { _id: sheet._id },
    [
      {
        $set: {
          "analytics.totalEdits": { $add: [{ $ifNull: ["$analytics.totalEdits", 0] }, 1] },
          "analytics.totalFormulaCells": {
            $max: [
              0,
              { $add: [{ $ifNull: ["$analytics.totalFormulaCells", 0] }, formulaCountDelta] },
            ],
          },
          "analytics.totalMergedCells": totalMergedCells,
          "analytics.lastEditedBy": userEmail,
          "analytics.lastEditedAt": lastEditedAt,
        },
      },
    ],
    queryOptions
  );
};

const findSheetForUser = async (sheetId, userId) => {
  const sheet = await Sheet.findById(sheetId);

  if (!sheet) {
    return { sheet: null, role: null };
  }

  const role = getUserRole(sheet, userId);

  return { sheet, role };
};

const hydrateCollaboratorUsernames = async (sheet) => {
  if (!sheet?.collaborators?.some((collaborator) => !collaborator.username)) return false;

  const userIds = sheet.collaborators.map((collaborator) => collaborator.userId).filter(Boolean);
  const users = await User.find({ _id: { $in: userIds } }).select("email username").lean();
  const userMap = new Map(users.map((user) => [user._id.toString(), user]));
  let changed = false;

  sheet.collaborators.forEach((collaborator) => {
    const user = userMap.get(collaborator.userId?.toString());
    if (!user) return;

    if (!collaborator.username) {
      collaborator.username = user.username || user.email;
      changed = true;
    }
    if (collaborator.email !== user.email) {
      collaborator.email = user.email;
      changed = true;
    }
  });

  return changed;
};

const countFormulaCellsInRows = async (sheetId, session = null) => {
  const aggregate = SheetRow.aggregate([
    { $match: { sheetId: new mongoose.Types.ObjectId(sheetId) } },
    { $unwind: "$cells" },
    { $match: { "cells.formula": { $nin: [null, ""] } } },
    { $count: "count" },
  ]);
  if (session) aggregate.session(session);

  const [result] = await aggregate;

  return result?.count || 0;
};

const getFormulaCellCount = async (sheet) => {
  const cachedCount = Number(sheet.analytics?.totalFormulaCells);

  return Number.isFinite(cachedCount)
    ? cachedCount
    : countFormulaCellsInRows(sheet._id);
};

const migrateSheetRowsIfNeeded = async (sheet) => {
  if (!sheet) return;

  const sourceRows = Array.isArray(sheet.data) ? sheet.data : [];
  if (sourceRows.length === 0) return;

  const existingRowCount = await SheetRow.countDocuments({ sheetId: sheet._id });
  if (existingRowCount > 0) return;

  const owner = sheet.collaborators?.find((collaborator) => collaborator.role === "owner");

  try {
    await SheetRow.insertMany(
      sourceRows.map((row, rowIndex) => {
        const hasContent = rowHasProtectedContent(row);

        return {
          sheetId: sheet._id,
          rowIndex,
          cells: normalizeRowCells(row),
          ...buildRowSearchFields(row),
          ownerId: hasContent ? sheet.createdBy : null,
          ownerEmail: hasContent ? owner?.email || "" : "",
          ownerUsername: hasContent ? owner?.username || owner?.email || "" : "",
        };
      }),
      { ordered: false }
    );
  } catch (error) {
    if (error?.code !== 11000 && error?.name !== "MongoBulkWriteError") {
      throw error;
    }
  }

  sheet.data = [];
  sheet.markModified("data");
  await sheet.save();
};

const ensureSheetRowsSearchText = async (sheetId) => {
  const missingSearchFieldsFilter = {
    sheetId,
    $or: [
      { searchText: { $exists: false } },
      {
        $and: [
          { searchText: { $nin: [null, ""] } },
          {
            $or: [
              { searchTokens: { $exists: false } },
              { searchTokens: { $size: 0 } },
            ],
          },
        ],
      },
    ],
  };
  const missingCount = await SheetRow.countDocuments({
    ...missingSearchFieldsFilter,
  });

  if (missingCount === 0) return;

  const cursor = SheetRow.find(missingSearchFieldsFilter)
    .select("cells")
    .cursor();
  const operations = [];

  for await (const row of cursor) {
    operations.push({
      updateOne: {
        filter: { _id: row._id },
        update: { $set: buildRowSearchFields(row.cells) },
      },
    });

    if (operations.length >= 500) {
      await SheetRow.bulkWrite(operations);
      operations.length = 0;
    }
  }

  if (operations.length > 0) {
    await SheetRow.bulkWrite(operations);
  }
};

const searchBackfillsInFlight = new Set();

const scheduleSheetRowsSearchBackfill = (sheetId) => {
  const key = sheetId.toString();
  if (searchBackfillsInFlight.has(key)) return;

  searchBackfillsInFlight.add(key);
  setImmediate(async () => {
    try {
      await ensureSheetRowsSearchText(sheetId);
    } catch (error) {
      console.error("Failed to backfill sheet row search tokens", error);
    } finally {
      searchBackfillsInFlight.delete(key);
    }
  });
};

const getRowsForSheet = async (sheetId, start = 0, limit = DEFAULT_ROW_PAGE_SIZE) => {
  const rows = await SheetRow.find({ sheetId, rowIndex: { $gte: start, $lt: start + limit } })
    .sort({ rowIndex: 1 })
    .select("rowIndex cells ownerId ownerEmail ownerUsername")
    .lean();
  const rowMap = new Map(rows.map((row) => [row.rowIndex, normalizeRowCells(row.cells)]));
  const rowOwners = Object.fromEntries(
    rows
      .filter((row) => row.ownerId)
      .map((row) => [
        row.rowIndex,
        {
          userId: row.ownerId.toString(),
          email: row.ownerEmail || "",
          username: row.ownerUsername || row.ownerEmail || "",
        },
      ])
  );

  return {
    data: Array.from({ length: limit }, (_, offset) => (
      rowMap.get(start + offset) || normalizeRowCells([])
    )),
    rowOwners,
  };
};

const getSheetRowCount = async (sheetId) => {
  const lastRow = await SheetRow.findOne({ sheetId }).sort({ rowIndex: -1 }).select("rowIndex").lean();
  return Math.max(DEFAULT_SHEET_ROWS, lastRow ? lastRow.rowIndex + 1 : 0);
};

const getLastDataRowIndex = async (sheetId) => {
  const indexedRow = await SheetRow.findOne({
    sheetId,
    searchText: { $regex: "\\S" },
  })
    .sort({ rowIndex: -1 })
    .select("rowIndex")
    .lean();

  if (indexedRow) return indexedRow.rowIndex;

  const totalRows = await getSheetRowCount(sheetId);
  const pageSize = 500;

  for (let end = totalRows; end > 0; end -= pageSize) {
    const start = Math.max(0, end - pageSize);
    const rows = await SheetRow.find({
      sheetId,
      rowIndex: { $gte: start, $lt: end },
    })
      .sort({ rowIndex: -1 })
      .select("rowIndex cells")
      .lean();

    const lastDataRow = rows.find((row) => rowHasContent(row.cells));
    if (lastDataRow) return lastDataRow.rowIndex;
  }

  return 0;
};

const ensureSheetRow = async (sheetId, rowIndex) => {
  const row = await SheetRow.findOneAndUpdate(
    { sheetId, rowIndex },
    {
      $setOnInsert: {
        sheetId,
        rowIndex,
        cells: normalizeRowCells([]),
        searchText: "",
        searchTokens: [],
      },
    },
    { returnDocument: "after", upsert: true }
  );

  row.cells = normalizeRowCells(row.cells);
  return row;
};

const createRowLockedError = (row) => {
  const error = new Error(
    `Row ${row.rowIndex + 1} is locked by ${row.ownerUsername || row.ownerEmail || "another user"}`
  );
  error.statusCode = 403;
  return error;
};

const ensureRowEditAccess = async (sheet, user, role, rowIndex) => {
  let row = await ensureSheetRow(sheet._id, rowIndex);
  const isPrivileged = canBypassRowLocks(role);
  const isOwner = row.ownerId?.toString() === user.id.toString();

  if (row.ownerId && !isOwner && !isPrivileged) {
    throw createRowLockedError(row);
  }

  if (!row.ownerId) {
    const claimedRow = await SheetRow.findOneAndUpdate(
      {
        _id: row._id,
        $or: [{ ownerId: null }, { ownerId: { $exists: false } }],
      },
      {
        $set: {
          ownerId: user.id,
          ownerEmail: user.email,
          ownerUsername: user.username || user.email,
        },
      },
      { returnDocument: "after" }
    );

    row = claimedRow || await ensureSheetRow(sheet._id, rowIndex);

    if (
      row.ownerId?.toString() !== user.id.toString() &&
      !isPrivileged
    ) {
      throw createRowLockedError(row);
    }
  }

  return row;
};

const getRowOwnershipMap = (rows = []) => Object.fromEntries(
  rows
    .filter(Boolean)
    .map((row) => [
      row.rowIndex,
      row.ownerId
        ? {
            userId: row.ownerId.toString(),
            email: row.ownerEmail || "",
            username: row.ownerUsername || row.ownerEmail || "",
          }
        : null,
    ])
);

const applyCellPatchesToRows = async (sheet, user, role, patches) => {
  const patchesByRow = new Map();

  patches.forEach((patch) => {
    const rowIndex = Number(patch.rowIndex);
    const colIndex = Number(patch.colIndex);

    if (!isValidCellIndex(rowIndex) || !isValidCellIndex(colIndex)) {
      return;
    }

    const rowPatches = patchesByRow.get(rowIndex) || [];
    rowPatches.push({ ...patch, rowIndex, colIndex });
    patchesByRow.set(rowIndex, rowPatches);
  });

  const rowIndexes = Array.from(patchesByRow.keys());
  const isPrivileged = canBypassRowLocks(role);
  const existingRows = await SheetRow.find({
    sheetId: sheet._id,
    rowIndex: { $in: rowIndexes },
  });
  const existingRowsByIndex = new Map(existingRows.map((row) => [row.rowIndex, row]));
  const protectedRowIndexes = rowIndexes.filter((rowIndex) => (
    patchesByRow.get(rowIndex).some((patch) => patch.colIndex <= ROW_LOCK_LAST_COLUMN_INDEX)
  ));

  for (const rowIndex of protectedRowIndexes) {
    const row = existingRowsByIndex.get(rowIndex);
    const isOwner = row?.ownerId?.toString() === user.id.toString();

    if (row?.ownerId && !isOwner && !isPrivileged) {
      throw createRowLockedError(row);
    }
  }

  const missingRowIndexes = rowIndexes.filter((rowIndex) => !existingRowsByIndex.has(rowIndex));

  if (missingRowIndexes.length > 0) {
    try {
      await SheetRow.bulkWrite(
        missingRowIndexes.map((rowIndex) => ({
          updateOne: {
            filter: { sheetId: sheet._id, rowIndex },
            update: {
              $setOnInsert: {
                sheetId: sheet._id,
                rowIndex,
                cells: normalizeRowCells([]),
                searchText: "",
                searchTokens: [],
              },
            },
            upsert: true,
          },
        })),
        { ordered: false }
      );
    } catch (error) {
      if (error?.code !== 11000 && error?.name !== "MongoBulkWriteError") {
        throw error;
      }
    }
  }

  const rows = await SheetRow.find({
    sheetId: sheet._id,
    rowIndex: { $in: rowIndexes },
  });
  const rowsByIndex = new Map(rows.map((row) => {
    row.cells = normalizeRowCells(row.cells);
    return [row.rowIndex, row];
  }));
  const unownedProtectedRowIndexes = protectedRowIndexes.filter((rowIndex) => !rowsByIndex.get(rowIndex)?.ownerId);

  if (unownedProtectedRowIndexes.length > 0) {
    await SheetRow.bulkWrite(
      unownedProtectedRowIndexes.map((rowIndex) => ({
        updateOne: {
          filter: {
            sheetId: sheet._id,
            rowIndex,
            $or: [{ ownerId: null }, { ownerId: { $exists: false } }],
          },
          update: {
            $set: {
              ownerId: user.id,
              ownerEmail: user.email,
              ownerUsername: user.username || user.email,
            },
          },
        },
      }))
    );

    const claimedRows = await SheetRow.find({
      sheetId: sheet._id,
      rowIndex: { $in: unownedProtectedRowIndexes },
    });

    claimedRows.forEach((row) => {
      row.cells = normalizeRowCells(row.cells);
      rowsByIndex.set(row.rowIndex, row);
    });
  }

  for (const rowIndex of protectedRowIndexes) {
    const row = rowsByIndex.get(rowIndex);
    const isOwner = row?.ownerId?.toString() === user.id.toString();

    if (row?.ownerId && !isOwner && !isPrivileged) {
      throw createRowLockedError(row);
    }
  }

  const updatedRows = [];
  const changeLogs = [];
  const rowOperations = [];
  let formulaCountDelta = 0;

  for (const [rowIndex, rowPatches] of patchesByRow.entries()) {
    const editsProtectedColumns = rowPatches.some(
      (patch) => patch.colIndex <= ROW_LOCK_LAST_COLUMN_INDEX
    );
    const row = rowsByIndex.get(rowIndex);
    if (!row) {
      throw new Error(`Failed to load row ${rowIndex + 1}`);
    }

    rowPatches.forEach((patch) => {
      const currentCell = row.cells[patch.colIndex] || createCell("");
      const previousHasFormula = Boolean(currentCell && typeof currentCell === "object" && currentCell.formula);
      const nextHasFormula = Boolean(patch.formula);
      const oldValue = currentCell && typeof currentCell === "object"
        ? currentCell.formula || currentCell.value || ""
        : currentCell;
      const newValue = patch.formula || patch.value || "";

      formulaCountDelta += Number(nextHasFormula) - Number(previousHasFormula);

      if (String(oldValue ?? "") !== String(newValue ?? "")) {
        changeLogs.push({
          sheetId: sheet._id,
          workspaceId: sheet.workspaceId || null,
          userId: user.id,
          userEmail: user.email,
          rowIndex,
          colIndex: patch.colIndex,
          cellAddress: cellAddress(rowIndex, patch.colIndex),
          oldValue,
          newValue,
          changeType: patch.formula || String(patch.value || "").startsWith("=") ? "formula" : "value",
        });
      }

      row.cells[patch.colIndex] = {
        ...currentCell,
        value: patch.value ?? "",
        formula: patch.formula || "",
        style: {
          ...defaultCellStyle,
          ...(currentCell.style || {}),
          ...(patch.style || {}),
        },
      };
    });

    const searchFields = buildRowSearchFields(row.cells);
    row.searchText = searchFields.searchText;
    row.searchTokens = searchFields.searchTokens;

    if (editsProtectedColumns && !rowHasProtectedContent(row.cells)) {
      row.ownerId = null;
      row.ownerEmail = "";
      row.ownerUsername = "";
    }

    rowOperations.push({
      updateOne: {
        filter: { _id: row._id },
        update: {
          $set: {
            cells: row.cells,
            searchText: row.searchText,
            searchTokens: row.searchTokens,
            ownerId: row.ownerId || null,
            ownerEmail: row.ownerEmail || "",
            ownerUsername: row.ownerUsername || "",
          },
        },
      },
    });
    updatedRows.push(row);
  }

  if (rowOperations.length > 0) {
    await SheetRow.bulkWrite(rowOperations);
  }

  if (changeLogs.length > 0) {
    await ChangeLog.insertMany(changeLogs);
  }

  const pendingCodeUpdates = updatedRows.map((row) => ({
    rowIndex: row.rowIndex,
    pending: rowNeedsCode(row.cells),
  }));

  return {
    rowOwners: getRowOwnershipMap(updatedRows),
    pendingCodeUpdates,
    formulaCountDelta,
  };
};

const applyCellPatchesForUser = async ({ sheet, user, rowIndex, colIndex, value, formula, patches }) => {
  const normalizedPatches = Array.isArray(patches) && patches.length > 0
    ? patches.slice(0, 50)
    : [{ rowIndex, colIndex, value, formula }];
  const role = getUserRole(sheet, user.id);

  await migrateSheetRowsIfNeeded(sheet);

  if (
    normalizedPatches.some((patch) => Number(patch.colIndex) === FIRST_CONFIRMATION_COLUMN_INDEX) &&
    !isFirstConfirmationEditOpen()
  ) {
    throw createTimeLockedColumnError();
  }

  const { rowOwners, pendingCodeUpdates, formulaCountDelta } = await applyCellPatchesToRows(sheet, user, role, normalizedPatches);

  await updateSheetAnalytics(sheet, user.email, { formulaCountDelta });

  return { patches: normalizedPatches, rowOwners, pendingCodeUpdates };
};

const createChangeLog = async ({
  sheet,
  user,
  rowIndex,
  colIndex,
  oldValue,
  newValue,
  changeType,
}) => {
  return ChangeLog.create({
    sheetId: sheet._id,
    workspaceId: sheet.workspaceId || null,
    userId: user.id,
    userEmail: user.email,
    rowIndex,
    colIndex,
    cellAddress: cellAddress(rowIndex, colIndex),
    oldValue,
    newValue,
    changeType,
  });
};

app.get("/api/health", (req, res) => {
  res.json({ message: "Sheet SaaS API is running" });
});

app.use(authRoutes);
app.use(profileRoutes);
app.use(workspaceRoutes);
app.use(itemCodingRoutes);

/* SHEETS */

app.post("/sheet", auth, async (req, res) => {
  try {
    const workspaceId = req.body.workspaceId || null;

    if (workspaceId) {
      if (!isValidObjectId(workspaceId)) {
        return rejectInvalidId(res, "workspace id");
      }

      const workspace = await Workspace.findById(workspaceId);

      if (!workspace) {
        return res.status(404).json({ message: "Workspace not found" });
      }

      const workspaceRole = getWorkspaceRole(workspace, req.user.id);

      if (!canUseWorkspace(workspaceRole)) {
        return res.status(403).json({ message: "Workspace access denied" });
      }
    }

    const erpType = req.body.erpType || "custom";

    if (!ALLOWED_SHEET_TYPES.has(erpType)) {
      return res.status(400).json({ message: "Invalid sheet type" });
    }

    const isERP = erpType !== "custom";
    const initialRows = isERP ? createERPTemplateData(erpType) : [];

    const sheet = await Sheet.create({
      name: req.body.name || "New Sheet",
      workspaceId,
      createdBy: req.user.id,
      data: [],
      meta: createDefaultMeta(),
      erpTemplate: {
        enabled: isERP,
        type: erpType,
        moduleName: isERP ? erpType : "",
      },
      erpOptions: createDefaultErpOptions(),
      analytics: {
        totalEdits: 0,
        totalFormulaCells: 0,
        totalMergedCells: 0,
        lastEditedBy: "",
        lastEditedAt: null,
        activeUsers: 0,
      },
      collaborators: [
        {
          userId: req.user.id,
          email: req.user.email,
          username: req.user.username || req.user.email,
          role: "owner",
        },
      ],
    });

    const initialRowDocuments = initialRows
      .map((row, rowIndex) => ({
        sheetId: sheet._id,
        rowIndex,
        cells: normalizeRowCells(row),
        ...buildRowSearchFields(row),
        ownerId: req.user.id,
        ownerEmail: req.user.email,
        ownerUsername: req.user.username || req.user.email,
      }))
      .filter((row) => row.searchText);

    if (initialRowDocuments.length > 0) {
      await SheetRow.insertMany(initialRowDocuments);
    }

    res.status(201).json(sheet);
  } catch (error) {
    res.status(500).json({ message: "Failed to create sheet" });
  }
});

app.get("/sheets", auth, async (req, res) => {
  try {
    const filter = {
      "collaborators.userId": req.user.id,
    };

    if (req.query.workspaceId) {
      if (!isValidObjectId(req.query.workspaceId)) {
        return rejectInvalidId(res, "workspace id");
      }

      filter.workspaceId = req.query.workspaceId;
    }

    const sheets = await Sheet.find(filter)
      .select("-data")
      .sort({ updatedAt: -1 });

    res.json(sheets);
  } catch (error) {
    res.status(500).json({ message: "Failed to get sheets" });
  }
});

app.get("/sheet/:id", auth, async (req, res) => {
  try {
    const rowLimit = Math.min(
      Number.parseInt(req.query.rowLimit, 10) || DEFAULT_ROW_PAGE_SIZE,
      MAX_ROW_PAGE_SIZE
    );
    const focusLastDataRow = req.query.focus === "lastData";
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canRead(role)) {
      return res.status(404).json({ message: "Sheet not found or access denied" });
    }

    await migrateSheetRowsIfNeeded(sheet);

    let sheetChanged = false;
    if (!sheet.meta) {
      sheet.meta = createDefaultMeta();
      sheetChanged = true;
    }
    if (!sheet.erpOptions) {
      sheet.erpOptions = createDefaultErpOptions();
      sheetChanged = true;
    }
    sheetChanged = await hydrateCollaboratorUsernames(sheet) || sheetChanged;

    if (sheetChanged) {
      await sheet.save();
    }

    const totalRows = await getSheetRowCount(req.params.id);
    const lastDataRowIndex = focusLastDataRow ? await getLastDataRowIndex(sheet._id) : 0;
    const rowStart = focusLastDataRow
      ? Math.max(0, Math.min(lastDataRowIndex, totalRows - 1) - Math.floor(rowLimit / 2))
      : Math.max(0, Number.parseInt(req.query.rowStart, 10) || 0);
    const rowsResult = await getRowsForSheet(
      sheet._id,
      rowStart,
      Math.max(0, Math.min(rowLimit, totalRows - rowStart))
    );
    const sheetObject = sheet.toObject();
    sheetObject.data = rowsResult.data;
    sheetObject.rowOwners = rowsResult.rowOwners;

    sendJson(req, res, {
      sheet: sheetObject,
      role,
      rows: {
        start: rowStart,
        count: rowsResult.data.length,
        total: totalRows,
        lastDataRowIndex,
      },
    });
  } catch (error) {
    res.status(500).json({ message: "Failed to get sheet" });
  }
});

app.get("/sheet/:id/rows", auth, async (req, res) => {
  try {
    const start = Math.max(0, Number.parseInt(req.query.start, 10) || 0);
    const limit = Math.min(
      Math.max(1, Number.parseInt(req.query.limit, 10) || DEFAULT_ROW_PAGE_SIZE),
      MAX_ROW_PAGE_SIZE
    );

    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canRead(role)) {
      return res.status(404).json({ message: "Sheet not found or access denied" });
    }

    await migrateSheetRowsIfNeeded(sheet);

    const totalRows = await getSheetRowCount(req.params.id);
    const rowsResult = await getRowsForSheet(
      sheet._id,
      start,
      Math.max(0, Math.min(limit, totalRows - start))
    );

    sendJson(req, res, {
      rows: rowsResult.data,
      rowOwners: rowsResult.rowOwners,
      start,
      count: rowsResult.data.length,
      total: totalRows,
    });
  } catch (error) {
    res.status(500).json({ message: "Failed to get sheet rows" });
  }
});

app.get("/sheet/:id/search", auth, async (req, res) => {
  try {
    const query = String(req.query.q || "").trim().toLowerCase();
    const limit = Math.min(Math.max(1, Number.parseInt(req.query.limit, 10) || 100), 500);

    if (!query) {
      return sendJson(req, res, { matches: [], total: 0 });
    }

    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canRead(role)) {
      return res.status(404).json({ message: "Sheet not found or access denied" });
    }

    await migrateSheetRowsIfNeeded(sheet);
    scheduleSheetRowsSearchBackfill(sheet._id);

    const queryTokens = query.length >= 2 ? buildSearchQueryTokens(query).slice(0, 20) : [];
    const rowFilter = queryTokens.length > 0
      ? { sheetId: sheet._id, searchTokens: { $all: queryTokens } }
      : { sheetId: sheet._id, searchText: { $regex: escapeRegex(query), $options: "i" } };

    const rowDocs = await SheetRow.find(rowFilter)
      .sort({ rowIndex: 1 })
      .limit(limit)
      .select("rowIndex cells")
      .lean();
    const matches = [];

    for (const row of rowDocs) {
      for (let colIndex = 0; colIndex < DEFAULT_SHEET_COLS; colIndex += 1) {
        const cell = row.cells?.[colIndex] || {};
        const text = `${cell.value ?? ""} ${cell.formula || ""}`.toLowerCase();

        if (text.includes(query)) {
          matches.push({ rowIndex: row.rowIndex, colIndex });
          if (matches.length >= limit) {
            return sendJson(req, res, { matches, total: matches.length, truncated: true });
          }
        }
      }
    }

    sendJson(req, res, { matches, total: matches.length, truncated: false });
  } catch (error) {
    res.status(500).json({ message: "Failed to search sheet" });
  }
});

app.get("/sheet/:id/pending-code-rows", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canRead(role)) {
      return res.status(404).json({ message: "Sheet not found or access denied" });
    }

    await migrateSheetRowsIfNeeded(sheet);

    const firstConfirmationValuePath = `cells.${FIRST_CONFIRMATION_COLUMN_INDEX}.value`;
    const firstConfirmationFormulaPath = `cells.${FIRST_CONFIRMATION_COLUMN_INDEX}.formula`;
    const candidateRows = await SheetRow.find({
      sheetId: sheet._id,
      $or: [
        { [firstConfirmationValuePath]: { $exists: true, $nin: ["", null] } },
        { [firstConfirmationFormulaPath]: { $exists: true, $nin: ["", null] } },
      ],
    })
      .sort({ rowIndex: 1 })
      .select("rowIndex cells")
      .lean();

    const rowIndexes = candidateRows
      .filter((row) => rowNeedsCode(row.cells))
      .map((row) => row.rowIndex);

    sendJson(req, res, { rowIndexes });
  } catch (error) {
    res.status(500).json({ message: "Failed to load pending code rows" });
  }
});

app.get("/sheet/:id/tax-export-rows", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canRead(role)) {
      return res.status(404).json({ message: "Sheet not found or access denied" });
    }

    const fromDate = parseRequiredDate(req.query.from);
    const toDate = parseRequiredDate(req.query.to);

    if (!fromDate || !toDate || fromDate > toDate) {
      return res.status(400).json({ message: "Choose a valid tax export date range" });
    }

    await migrateSheetRowsIfNeeded(sheet);

    const codeChanges = await ChangeLog.find({
      sheetId: sheet._id,
      colIndex: TAX_ITEM_CODE_COLUMN_INDEX,
      changeType: { $in: ["value", "formula"] },
      createdAt: { $gte: fromDate, $lte: toDate },
    })
      .sort({ createdAt: 1 })
      .select("rowIndex newValue createdAt")
      .lean();

    const rowIndexes = Array.from(new Set(
      codeChanges
        .filter((change) => String(change.newValue || "").trim())
        .map((change) => change.rowIndex)
    ));

    if (!rowIndexes.length) {
      return sendJson(req, res, { rows: [] });
    }

    const changedAtByRow = new Map();
    codeChanges.forEach((change) => {
      if (!String(change.newValue || "").trim()) return;
      changedAtByRow.set(change.rowIndex, change.createdAt);
    });

    const rows = await SheetRow.find({
      sheetId: sheet._id,
      rowIndex: { $in: rowIndexes },
    })
      .select("rowIndex cells")
      .lean();

    const taxRows = rows
      .map((row) => ({
        rowIndex: row.rowIndex,
        changedAt: changedAtByRow.get(row.rowIndex),
        cells: normalizeRowCells(row.cells),
      }))
      .filter((row) => {
        const itemName = getCellText(row.cells, TAX_ITEM_NAME_COLUMN_INDEX);
        const firstConfirmation = getCellText(row.cells, FIRST_CONFIRMATION_COLUMN_INDEX);
        const itemCode = getCellText(row.cells, TAX_ITEM_CODE_COLUMN_INDEX);

        return itemName && firstConfirmation && itemCode;
      })
      .sort((left, right) => {
        const sortColumns = [
          TAX_MAIN_GROUP_COLUMN_INDEX,
          TAX_SUB_GROUP_COLUMN_INDEX,
          TAX_SUB_SUB_GROUP_COLUMN_INDEX,
          TAX_SUPPORT_GROUP_COLUMN_INDEX,
          TAX_DETAILED_GROUP_COLUMN_INDEX,
          TAX_ITEM_NAME_COLUMN_INDEX,
        ];

        for (const colIndex of sortColumns) {
          const result = compareText(getCellText(left.cells, colIndex), getCellText(right.cells, colIndex));
          if (result !== 0) return result;
        }

        return left.rowIndex - right.rowIndex;
      });

    sendJson(req, res, { rows: taxRows });
  } catch (error) {
    res.status(500).json({ message: "Failed to load tax export rows" });
  }
});

app.get("/sheet/:id/export.csv", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canRead(role)) {
      return res.status(404).json({ message: "Sheet not found or access denied" });
    }

    await migrateSheetRowsIfNeeded(sheet);

    const safeName = String(sheet.name || "sheet").replace(/[^\w.-]+/g, "_");
    const cursor = SheetRow.find({ sheetId: sheet._id })
      .sort({ rowIndex: 1 })
      .select("cells")
      .lean()
      .cursor();

    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Content-Disposition", `attachment; filename="${safeName}.csv"`);
    res.write("\ufeff");
    res.write(Array.from({ length: DEFAULT_SHEET_COLS }, (_, index) => escapeCsvCell(cellAddress(0, index).replace("1", ""))).join(",") + "\n");

    for await (const row of cursor) {
      const values = Array.from({ length: DEFAULT_SHEET_COLS }, (_, colIndex) => {
        const cell = row.cells?.[colIndex] || {};
        return escapeCsvCell(cell.formula || cell.value || "");
      });
      res.write(values.join(",") + "\n");
    }

    res.end();
  } catch (error) {
    res.status(500).json({ message: "Failed to export sheet" });
  }
});

app.put("/sheet/:id", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canEdit(role)) {
      return res.status(403).json({ message: "You do not have edit permission" });
    }

    sheet.meta = req.body.meta || sheet.meta || createDefaultMeta();

    await migrateSheetRowsIfNeeded(sheet);

    await sheet.save();
    await updateSheetAnalytics(sheet, req.user.email, {
      totalFormulaCells: await countFormulaCellsInRows(sheet._id),
    });

    await ChangeLog.create({
      sheetId: sheet._id,
      workspaceId: sheet.workspaceId || null,
      userId: req.user.id,
      userEmail: req.user.email,
      rowIndex: 0,
      colIndex: 0,
      cellAddress: "FULL_SHEET",
      oldValue: "Previous sheet state",
      newValue: "Updated sheet state",
      changeType: "import",
    });

    const updatedSheet = await Sheet.findById(sheet._id);
    const sheetObject = updatedSheet.toObject();
    sheetObject.data = [];

    io.to(req.params.id).emit("sheet-saved", sheetObject);

    res.json({ ok: true, sheet: sheetObject });
  } catch (error) {
    res.status(500).json({ message: "Failed to update sheet" });
  }
});

app.post("/sheet/:id/import-rows", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canBypassRowLocks(role)) {
      return res.status(403).json({ message: "Only owner or admin can import rows" });
    }

    const start = Math.max(0, Number.parseInt(req.body.start, 10) || 0);
    const rows = Array.isArray(req.body.rows) ? req.body.rows.slice(0, MAX_ROW_PAGE_SIZE) : [];
    const rowOperations = rows.map((row, offset) => {
      const hasContent = rowHasProtectedContent(row);

      return {
        updateOne: {
          filter: { sheetId: sheet._id, rowIndex: start + offset },
          update: {
            $set: {
              sheetId: sheet._id,
              rowIndex: start + offset,
              cells: normalizeRowCells(row),
              ...buildRowSearchFields(row),
              ownerId: hasContent ? req.user.id : null,
              ownerEmail: hasContent ? req.user.email : "",
              ownerUsername: hasContent ? req.user.username || req.user.email : "",
            },
          },
          upsert: true,
        },
      };
    });

    await runWithOptionalTransaction(async (session) => {
      const queryOptions = session ? { session } : undefined;

      if (rowOperations.length > 0) {
        await SheetRow.bulkWrite(rowOperations, queryOptions);
      }

      if (req.body.reset === true) {
        const deleteFilter = rows.length > 0
          ? {
              sheetId: sheet._id,
              $or: [
                { rowIndex: { $lt: start } },
                { rowIndex: { $gte: start + rows.length } },
              ],
            }
          : { sheetId: sheet._id };

        await SheetRow.deleteMany(deleteFilter, queryOptions);
      }

      await Sheet.updateOne({ _id: sheet._id }, { $set: { data: [] } }, queryOptions);
      await updateSheetAnalytics(sheet, req.user.email, {
        totalFormulaCells: await countFormulaCellsInRows(sheet._id, session),
        session,
      });
    });

    res.json({
      ok: true,
      start,
      count: rows.length,
      total: await getSheetRowCount(sheet._id),
    });
  } catch (error) {
    res.status(500).json({ message: "Failed to import rows" });
  }
});

app.patch("/sheet/:id/cells", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canEdit(role)) {
      return res.status(403).json({ message: "You do not have edit permission" });
    }

    const { rowIndex, colIndex, value, formula, patches } = req.body;

    const firstPatch = Array.isArray(patches) ? patches[0] : null;
    const targetRowIndex = rowIndex ?? firstPatch?.rowIndex;
    const targetColIndex = colIndex ?? firstPatch?.colIndex;

    if (!isValidCellIndex(Number(targetRowIndex)) || !isValidCellIndex(Number(targetColIndex))) {
      return res.status(400).json({ message: "Invalid cell position" });
    }

    const editResult = await applyCellPatchesForUser({
      sheet,
      user: req.user,
      rowIndex: targetRowIndex,
      colIndex: targetColIndex,
      value,
      formula,
      patches,
    });

    io.to(req.params.id).emit("cell-change", {
      rowIndex: targetRowIndex,
      colIndex: targetColIndex,
      value,
      formula,
      patches: editResult.patches,
      rowOwners: editResult.rowOwners,
      pendingCodeUpdates: editResult.pendingCodeUpdates,
      updatedBy: req.user.email,
    });

    res.json({
      ok: true,
      rowOwners: editResult.rowOwners,
      pendingCodeUpdates: editResult.pendingCodeUpdates,
    });
  } catch (error) {
    console.error("Failed to update cell", error);
    res.status(error.statusCode || 500).json({ message: error.message || "Failed to update cell" });
  }
});

app.patch("/sheet/:id/cell-style", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canEdit(role)) {
      return res.status(403).json({ message: "You do not have edit permission" });
    }

    const { rowIndex, colIndex, style } = req.body;

    if (!isValidCellIndex(Number(rowIndex)) || !isValidCellIndex(Number(colIndex))) {
      return res.status(400).json({ message: "Invalid cell position" });
    }

    if (Number(colIndex) === FIRST_CONFIRMATION_COLUMN_INDEX && !isFirstConfirmationEditOpen()) {
      throw createTimeLockedColumnError();
    }

    await migrateSheetRowsIfNeeded(sheet);
    const row = Number(colIndex) <= ROW_LOCK_LAST_COLUMN_INDEX
      ? await ensureRowEditAccess(sheet, req.user, role, Number(rowIndex))
      : await ensureSheetRow(sheet._id, Number(rowIndex));
    const cell = row.cells[Number(colIndex)] || createCell("");
    const previousStyle = { ...(cell.style || {}) };

    row.cells[Number(colIndex)] = {
      value: cell.value ?? "",
      formula: cell.formula || "",
      style: {
        ...defaultCellStyle,
        ...previousStyle,
        ...(style || {}),
      },
    };
    row.markModified("cells");
    await row.save();

    await createChangeLog({
      sheet,
      user: req.user,
      rowIndex: Number(rowIndex),
      colIndex: Number(colIndex),
      oldValue: previousStyle,
      newValue: style,
      changeType: "style",
    });

    await updateSheetAnalytics(sheet, req.user.email);

    io.to(req.params.id).emit("cell-style-change", {
      rowIndex,
      colIndex,
      style,
      rowOwners: getRowOwnershipMap([row]),
      updatedBy: req.user.email,
    });

    res.json({ ok: true, rowOwners: getRowOwnershipMap([row]) });
  } catch (error) {
    res.status(error.statusCode || 500).json({ message: error.message || "Failed to update cell style" });
  }
});

app.patch("/sheet/:id/name", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canManage(role)) {
      return res.status(403).json({ message: "Only owner can rename sheet" });
    }

    sheet.name = req.body.name || sheet.name;
    await sheet.save();

    io.to(req.params.id).emit("sheet-renamed", {
      sheetId: sheet._id,
      name: sheet.name,
    });

    res.json({ ok: true, sheet });
  } catch (error) {
    res.status(500).json({ message: "Failed to rename sheet" });
  }
});

app.post("/sheet/:id/share", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canManageSheetUsers(role)) {
      return res.status(403).json({ message: "Only sheet admins can share sheet" });
    }

    const identifier = String(req.body.identifier || req.body.email || "").trim();
    const email = identifier.toLowerCase();
    const newRole = String(req.body.role || "viewer");

    if (!identifier || !["admin", "editor", "viewer"].includes(newRole)) {
      return res.status(400).json({ message: "Valid username/email and role are required" });
    }

    const userToShare = await User.findOne({
      $or: [
        { email },
        { username: identifier },
      ],
    });

    if (!userToShare) {
      return res.status(404).json({ message: "User must sign up before sharing" });
    }

    const existing = sheet.collaborators.find(
      (item) => item.userId.toString() === userToShare._id.toString()
    );

    if (existing) {
      if (!canAssignSheetRole(role, existing.role, newRole)) {
        return res.status(403).json({ message: "You cannot assign this role" });
      }

      existing.role = newRole;
      existing.email = userToShare.email;
      existing.username = userToShare.username || userToShare.email;
    } else {
      if (!canAssignSheetRole(role, null, newRole)) {
        return res.status(403).json({ message: "You cannot assign this role" });
      }

      sheet.collaborators.push({
        userId: userToShare._id,
        email: userToShare.email,
        username: userToShare.username || userToShare.email,
        role: newRole,
      });
    }

    await sheet.save();

    io.to(req.params.id).emit("collaborators-updated", sheet.collaborators);

    res.json({ ok: true, sheet });
  } catch (error) {
    res.status(500).json({ message: "Failed to share sheet" });
  }
});

app.delete("/sheet/:id/collaborator/:userId", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canManageSheetUsers(role)) {
      return res.status(403).json({ message: "Only sheet admins can remove collaborators" });
    }

    sheet.collaborators = sheet.collaborators.filter(
      (item) =>
        item.userId.toString() !== req.params.userId ||
        !canRemoveSheetCollaborator(role, item.role)
    );

    await sheet.save();

    io.to(req.params.id).emit("collaborators-updated", sheet.collaborators);

    res.json({ ok: true, sheet });
  } catch (error) {
    res.status(500).json({ message: "Failed to remove collaborator" });
  }
});

app.delete("/sheet/:id", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canManage(role)) {
      return res.status(403).json({ message: "Only owner can delete sheet" });
    }

    await ChangeLog.deleteMany({ sheetId: sheet._id });
    await SheetRow.deleteMany({ sheetId: sheet._id });
    await Sheet.deleteOne({ _id: sheet._id });

    io.to(req.params.id).emit("sheet-deleted", {
      sheetId: req.params.id,
    });

    res.json({ ok: true, message: "Sheet deleted successfully" });
  } catch (error) {
    res.status(500).json({ message: "Failed to delete sheet" });
  }
});

/* ERP OPTIONS */

app.get("/sheet/:id/erp-options", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canRead(role)) {
      return res.status(403).json({ message: "Access denied" });
    }

    res.json(sheet.erpOptions || createDefaultErpOptions());
  } catch (error) {
    res.status(500).json({ message: "Failed to get ERP options" });
  }
});

app.patch("/sheet/:id/erp-options", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canManage(role)) {
      return res.status(403).json({ message: "Only owner can edit ERP options" });
    }

    sheet.erpOptions = {
      ...(sheet.erpOptions || createDefaultErpOptions()),
      ...(req.body.erpOptions || {}),
    };

    await sheet.save();

    io.to(req.params.id).emit("erp-options-updated", sheet.erpOptions);

    res.json({ ok: true, erpOptions: sheet.erpOptions });
  } catch (error) {
    res.status(500).json({ message: "Failed to update ERP options" });
  }
});

/* CHANGE LOGS */

app.get("/sheet/:id/changes", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canRead(role)) {
      return res.status(403).json({ message: "Access denied" });
    }

    const filter = { sheetId: req.params.id };

    if (req.query.rowIndex !== undefined && req.query.colIndex !== undefined) {
      filter.rowIndex = Number(req.query.rowIndex);
      filter.colIndex = Number(req.query.colIndex);
    }

    const requestedLimit = Number(req.query.limit);
    const limit = Number.isInteger(requestedLimit)
      ? Math.min(Math.max(requestedLimit, 1), 500)
      : 250;
    const changes = await ChangeLog.find(filter).sort({ createdAt: -1 }).limit(limit);

    res.json(changes);
  } catch (error) {
    res.status(500).json({ message: "Failed to get changes" });
  }
});

/* ADMIN ANALYTICS */

app.get("/admin/analytics", auth, async (req, res) => {
  try {
    const workspaces = await Workspace.find({ "members.userId": req.user.id });
    const workspaceIds = workspaces.map((workspace) => workspace._id);

    const sheets = await Sheet.find({
      $or: [
        { "collaborators.userId": req.user.id },
        { workspaceId: { $in: workspaceIds } },
      ],
    });

    const totalChanges = await ChangeLog.countDocuments({
      $or: [
        { userId: req.user.id },
        { workspaceId: { $in: workspaceIds } },
      ],
    });

    const totalUsers = new Set();

    sheets.forEach((sheet) => {
      sheet.collaborators.forEach((collaborator) => {
        totalUsers.add(collaborator.email);
      });
    });

    res.json({
      totalWorkspaces: workspaces.length,
      totalSheets: sheets.length,
      totalChanges,
      totalUsers: totalUsers.size,
      erpSheets: sheets.filter((sheet) => sheet.erpTemplate?.enabled).length,
      recentSheets: sheets
        .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt))
        .slice(0, 10),
    });
  } catch (error) {
    res.status(500).json({ message: "Failed to get analytics" });
  }
});

/* SOCKET */

const onlineUsersBySheet = new Map();

const getOnlineUsers = (sheetId) => {
  return Array.from(onlineUsersBySheet.get(sheetId)?.values() || []);
};

const broadcastPresence = async (sheetId) => {
  if (!isValidObjectId(sheetId)) {
    return;
  }

  const users = getOnlineUsers(sheetId);

  await Sheet.findByIdAndUpdate(sheetId, {
    "analytics.activeUsers": users.length,
  });

  io.to(sheetId).emit("presence-updated", users);
};

io.use((socket, next) => {
  try {
    const token = socket.handshake.auth?.token;

    if (!token) {
      return next(new Error("Unauthorized socket"));
    }

    socket.user = jwt.verify(token, process.env.JWT_SECRET);
    next();
  } catch (error) {
    next(new Error("Invalid socket token"));
  }
});

io.on("connection", (socket) => {
  socket.on("join-sheet", async (sheetId) => {
    try {
      if (!isValidObjectId(sheetId)) {
        socket.emit("socket-error", "Invalid sheet id");
        return;
      }

      const { sheet, role } = await findSheetForUser(sheetId, socket.user.id);

      if (!sheet || !canRead(role)) {
        socket.emit("socket-error", "Access denied");
        return;
      }

      socket.join(sheetId);
      socket.currentSheetId = sheetId;

      if (!onlineUsersBySheet.has(sheetId)) {
        onlineUsersBySheet.set(sheetId, new Map());
      }

      onlineUsersBySheet.get(sheetId).set(socket.id, {
        socketId: socket.id,
        userId: socket.user.id,
        email: socket.user.email,
        username: socket.user.username || socket.user.email,
        role,
      });

      await broadcastPresence(sheetId);
    } catch (error) {
      socket.emit("socket-error", "Failed to join sheet");
    }
  });

  socket.on("leave-sheet", async (sheetId) => {
    if (!isValidObjectId(sheetId)) {
      socket.emit("socket-error", "Invalid sheet id");
      return;
    }

    socket.leave(sheetId);

    if (onlineUsersBySheet.has(sheetId)) {
      onlineUsersBySheet.get(sheetId).delete(socket.id);
      await broadcastPresence(sheetId);
    }

    if (socket.currentSheetId === sheetId) {
      socket.currentSheetId = null;
    }
  });

  socket.on("cell-change", async ({ sheetId, rowIndex, colIndex, value, formula, patches }, ack) => {
    try {
      if (!isValidObjectId(sheetId)) {
        socket.emit("socket-error", "Invalid sheet id");
        if (typeof ack === "function") ack({ ok: false, message: "Invalid sheet id" });
        return;
      }

      const { sheet, role } = await findSheetForUser(sheetId, socket.user.id);

      if (!sheet || !canEdit(role)) {
        socket.emit("socket-error", "You do not have edit permission");
        if (typeof ack === "function") ack({ ok: false, message: "You do not have edit permission" });
        return;
      }

      const firstPatch = Array.isArray(patches) ? patches[0] : null;
      const targetRowIndex = rowIndex ?? firstPatch?.rowIndex;
      const targetColIndex = colIndex ?? firstPatch?.colIndex;

      if (!isValidCellIndex(Number(targetRowIndex)) || !isValidCellIndex(Number(targetColIndex))) {
        socket.emit("socket-error", "Invalid cell position");
        if (typeof ack === "function") ack({ ok: false, message: "Invalid cell position" });
        return;
      }

      const editResult = await applyCellPatchesForUser({
        sheet,
        user: socket.user,
        rowIndex: targetRowIndex,
        colIndex: targetColIndex,
        value,
        formula,
        patches,
      });

      socket.to(sheetId).emit("cell-change", {
        rowIndex: targetRowIndex,
        colIndex: targetColIndex,
        value,
        formula,
        patches: editResult.patches,
        rowOwners: editResult.rowOwners,
        pendingCodeUpdates: editResult.pendingCodeUpdates,
        updatedBy: socket.user.email,
      });

      if (typeof ack === "function") {
        ack({
          ok: true,
          rowOwners: editResult.rowOwners,
          pendingCodeUpdates: editResult.pendingCodeUpdates,
        });
      }
    } catch (error) {
      console.error("Failed to update cell", error);
      socket.emit("socket-error", error.message || "Failed to update cell");
      if (typeof ack === "function") ack({ ok: false, message: error.message || "Failed to update cell" });
    }
  });

  socket.on("cell-style-change", async ({ sheetId, rowIndex, colIndex, style }, ack) => {
    try {
      if (!isValidObjectId(sheetId)) {
        socket.emit("socket-error", "Invalid sheet id");
        if (typeof ack === "function") ack({ ok: false, message: "Invalid sheet id" });
        return;
      }

      const { sheet, role } = await findSheetForUser(sheetId, socket.user.id);

      if (!sheet || !canEdit(role)) {
        socket.emit("socket-error", "You do not have edit permission");
        if (typeof ack === "function") ack({ ok: false, message: "You do not have edit permission" });
        return;
      }

      if (!isValidCellIndex(Number(rowIndex)) || !isValidCellIndex(Number(colIndex))) {
        socket.emit("socket-error", "Invalid cell position");
        if (typeof ack === "function") ack({ ok: false, message: "Invalid cell position" });
        return;
      }

      if (Number(colIndex) === FIRST_CONFIRMATION_COLUMN_INDEX && !isFirstConfirmationEditOpen()) {
        throw createTimeLockedColumnError();
      }

      await migrateSheetRowsIfNeeded(sheet);
      const row = Number(colIndex) <= ROW_LOCK_LAST_COLUMN_INDEX
        ? await ensureRowEditAccess(sheet, socket.user, role, Number(rowIndex))
        : await ensureSheetRow(sheet._id, Number(rowIndex));
      const cell = row.cells[Number(colIndex)] || createCell("");
      const previousStyle = { ...(cell.style || {}) };
      row.cells[Number(colIndex)] = {
        value: cell.value ?? "",
        formula: cell.formula || "",
        style: {
          ...defaultCellStyle,
          ...previousStyle,
          ...(style || {}),
        },
      };
      row.markModified("cells");
      await row.save();

      await createChangeLog({
        sheet,
        user: socket.user,
        rowIndex: Number(rowIndex),
        colIndex: Number(colIndex),
        oldValue: previousStyle,
        newValue: style,
        changeType: "style",
      });

      await updateSheetAnalytics(sheet, socket.user.email);

      socket.to(sheetId).emit("cell-style-change", {
        rowIndex,
        colIndex,
        style,
        rowOwners: getRowOwnershipMap([row]),
        updatedBy: socket.user.email,
      });

      if (typeof ack === "function") ack({ ok: true, rowOwners: getRowOwnershipMap([row]) });
    } catch (error) {
      socket.emit("socket-error", error.message || "Failed to update cell style");
      if (typeof ack === "function") ack({ ok: false, message: error.message || "Failed to update cell style" });
    }
  });

  socket.on("cursor-change", ({ sheetId, rowIndex, colIndex }) => {
    if (!isValidObjectId(sheetId)) {
      socket.emit("socket-error", "Invalid sheet id");
      return;
    }

    if (socket.currentSheetId !== sheetId || !socket.rooms.has(sheetId)) {
      socket.emit("socket-error", "Access denied");
      return;
    }

    if (!isValidCellIndex(Number(rowIndex)) || !isValidCellIndex(Number(colIndex))) {
      socket.emit("socket-error", "Invalid cell position");
      return;
    }

    socket.to(sheetId).emit("cursor-change", {
      userId: socket.user.id,
      email: socket.user.email,
      rowIndex: Number(rowIndex),
      colIndex: Number(colIndex),
    });
  });

  socket.on("disconnect", async () => {
    const sheetId = socket.currentSheetId;

    if (sheetId && onlineUsersBySheet.has(sheetId)) {
      onlineUsersBySheet.get(sheetId).delete(socket.id);
      await broadcastPresence(sheetId);
    }
  });
});

if (process.env.NODE_ENV === "production") {
  const frontendDistPath = path.resolve(__dirname, "../frontend/dist");

  if (fs.existsSync(frontendDistPath)) {
    app.use(express.static(frontendDistPath));

    app.use((req, res, next) => {
      if (req.method !== "GET" || !req.accepts("html")) {
        return next();
      }

      return res.sendFile(path.join(frontendDistPath, "index.html"));
    });
  }
}

connectToDatabase().then(() => {
  server.listen(PORT, () => {
    console.log("SaaS API running on port " + PORT);
  });
});
