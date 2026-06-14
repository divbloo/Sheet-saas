require("dotenv").config();

const express = require("express");
const fs = require("fs");
const http = require("http");
const path = require("path");
const zlib = require("zlib");
const mongoose = require("mongoose");
const cors = require("cors");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");
const { Server } = require("socket.io");

const User = require("./models/User");
const Sheet = require("./models/Sheet");
const SheetRow = require("./models/SheetRow");
const Workspace = require("./models/Workspace");
const ChangeLog = require("./models/ChangeLog");
const {
  authLimiter,
  createCorsOptions,
  createHelmetOptions,
  getFrontendUrls,
  verifyProductionSecurity,
} = require("./config/security");
const defaultErpOptions = require("./config/defaultErpOptions.json");
const { isValidCellIndex, isValidObjectId } = require("./utils/validation");
const {
  DEFAULT_SHEET_COLS,
  buildRowSearchText,
  createCell,
  defaultCellStyle,
  normalizeRowCells,
} = require("./utils/sheetRows");

const app = express();
const server = http.createServer(app);

const PORT = process.env.PORT || 5000;
const FRONTEND_URLS = getFrontendUrls();
const corsOptions = createCorsOptions(FRONTEND_URLS);

const escapeCsvCell = (value) => {
  const text = String(value ?? "");
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
};

const escapeRegex = (value) => String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

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

const DEFAULT_COLUMN_WIDTHS = {
  0: 300,
  1: 300,
  2: 220,
  3: 220,
  4: 145,
  5: 220,
  6: 145,
  7: 82,
  8: 60,
  9: 60,
  10: 150,
  11: 92,
  12: 150,
  13: 92,
  14: 220,
};

const DEFAULT_SHEET_ROWS = 5000;
const DEFAULT_ROW_PAGE_SIZE = 50;
const MAX_ROW_PAGE_SIZE = 500;
const ROW_LOCK_LAST_COLUMN_INDEX = 10;
const ALLOWED_SHEET_TYPES = new Set(["custom", "item-master"]);

const createEmptySheetData = (rows = DEFAULT_SHEET_ROWS, cols = DEFAULT_SHEET_COLS) => {
  return Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => createCell(""))
  );
};

const rowHasProtectedContent = (row = []) => normalizeRowCells(row)
  .slice(0, ROW_LOCK_LAST_COLUMN_INDEX + 1)
  .some((cell) => String(cell.value ?? "").trim() || String(cell.formula || "").trim());

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

const visibleItemMasterHeaders = [
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

const createERPTemplateData = (type) => {
  const headersByType = {
    "item-master": [
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
    ],
    inventory: [
      "Item Code",
      "Item Name",
      "Category",
      "Warehouse",
      "Opening Qty",
      "In",
      "Out",
      "Balance",
      "Min Stock",
      "Status",
    ],
    sales: [
      "Date",
      "Customer",
      "Sales Person",
      "Brand",
      "Item",
      "Qty",
      "Unit Price",
      "Total",
      "Status",
      "Notes",
    ],
    finance: [
      "Date",
      "Account",
      "Description",
      "Debit",
      "Credit",
      "Balance",
      "Cost Center",
      "Status",
    ],
    purchasing: [
      "PR No.",
      "Supplier",
      "Item",
      "Qty",
      "Requested By",
      "Unit Price",
      "Delivery Date",
      "Status",
    ],
    hr: [
      "Employee ID",
      "Name",
      "Department",
      "Position",
      "Join Date",
      "Salary",
      "Leave Balance",
      "Status",
    ],
    crm: [
      "Lead",
      "Company",
      "Contact",
      "Phone",
      "Email",
      "Stage",
      "Next Action",
      "Owner",
    ],
    offers: [
      "Offer No.",
      "Customer",
      "Subject",
      "Item",
      "Qty",
      "Cost",
      "Selling Price",
      "Margin",
      "Approval",
      "Status",
    ],
    warehouse: [
      "Transaction Date",
      "Item Code",
      "Item Name",
      "Location",
      "Received",
      "Issued",
      "Balance",
      "Handled By",
    ],
  };

  const headers = type === "item-master"
    ? visibleItemMasterHeaders
    : headersByType[type] || headersByType["item-master"];
  const data = createEmptySheetData(1);

  headers.slice(0, DEFAULT_SHEET_COLS).forEach((header, index) => {
    data[0][index] = {
      value: header,
      formula: "",
      style: {
        ...defaultCellStyle,
        fontWeight: "bold",
        backgroundColor: "#ccfbf1",
        color: "#075985",
        textAlign: "center",
      },
    };
  });

  return data;
};

const signToken = (user) => {
  return jwt.sign(
    {
      id: user._id.toString(),
      email: user.email,
      username: user.username || user.email,
      role: user.role || "user",
    },
    process.env.JWT_SECRET,
    { expiresIn: "7d" }
  );
};

const toUserResponse = (user) => ({
  id: user._id,
  email: user.email,
  username: user.username || user.email,
  role: user.role || "user",
  avatarUrl: user.avatarUrl || "",
});

const normalizeUsername = (value) => {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .slice(0, 60);
};

const getDefaultUsername = (email) => String(email || "").split("@")[0] || "user";

const normalizeAvatarUrl = (value) => {
  const text = String(value || "").trim();
  if (!text) return "";

  if (!/^data:image\/(png|jpe?g|webp);base64,/i.test(text)) {
    return null;
  }

  return text.length <= 2_500_000 ? text : null;
};

const createUniqueUsername = async (email, requestedUsername = "") => {
  const base = normalizeUsername(requestedUsername) || getDefaultUsername(email);
  let username = base;
  let suffix = 2;

  while (await User.exists({ username })) {
    username = `${base}${suffix}`;
    suffix += 1;
  }

  return username;
};

const auth = (req, res, next) => {
  try {
    const header = req.headers.authorization;

    if (!header || !header.startsWith("Bearer ")) {
      return res.status(401).json({ message: "Unauthorized" });
    }

    const token = header.split(" ")[1];
    const decoded = jwt.verify(token, process.env.JWT_SECRET);

    req.user = decoded;
    next();
  } catch (error) {
    return res.status(401).json({ message: "Invalid token" });
  }
};

const getUserRole = (sheet, userId) => {
  const collaborator = sheet.collaborators.find(
    (item) => item.userId.toString() === userId.toString()
  );

  return collaborator ? collaborator.role : null;
};

const canRead = (role) => ["owner", "admin", "editor", "viewer"].includes(role);
const canEdit = (role) => ["owner", "admin", "editor"].includes(role);
const canManage = (role) => role === "owner";
const canBypassRowLocks = (role) => role === "owner" || role === "admin";
const canManageSheetUsers = (role) => ["owner", "admin"].includes(role);
const canAssignSheetRole = (actorRole, currentRole, nextRole) => {
  if (currentRole === "owner" || nextRole === "owner") return false;
  if (actorRole === "owner") return ["admin", "editor", "viewer"].includes(nextRole);
  if (actorRole === "admin") {
    return currentRole !== "admin" && ["editor", "viewer"].includes(nextRole);
  }
  return false;
};
const canRemoveSheetCollaborator = (actorRole, targetRole) => {
  if (targetRole === "owner") return false;
  if (actorRole === "owner") return true;
  return actorRole === "admin" && targetRole !== "admin";
};

const findSheetForUser = async (sheetId, userId) => {
  const sheet = await Sheet.findById(sheetId);

  if (!sheet) {
    return { sheet: null, role: null };
  }

  const role = getUserRole(sheet, userId);

  return { sheet, role };
};

const findSheetForUserProjected = async (sheetId, userId, projection) => {
  const sheet = await Sheet.findById(sheetId, projection);

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

const getWorkspaceRole = (workspace, userId) => {
  const member = workspace.members.find(
    (item) => item.userId.toString() === userId.toString()
  );

  return member ? member.role : null;
};

const canManageWorkspace = (role) => role === "admin";
const canUseWorkspace = (role) => ["admin", "member", "viewer"].includes(role);

const countFormulaCells = (data = []) => {
  let count = 0;

  data.forEach((row) => {
    row.forEach((cell) => {
      if (cell && typeof cell === "object" && cell.formula) {
        count += 1;
      }
    });
  });

  return count;
};

const countFormulaCellsInRows = async (sheetId) => {
  const [result] = await SheetRow.aggregate([
    { $match: { sheetId: new mongoose.Types.ObjectId(sheetId) } },
    { $unwind: "$cells" },
    { $match: { "cells.formula": { $nin: [null, ""] } } },
    { $count: "count" },
  ]);

  return result?.count || 0;
};

const migrateSheetRowsIfNeeded = async (sheet) => {
  if (!sheet) return;

  const existingRowCount = await SheetRow.countDocuments({ sheetId: sheet._id });
  if (existingRowCount > 0) return;

  const sourceRows = Array.isArray(sheet.data) ? sheet.data : [];
  if (sourceRows.length === 0) return;

  if (sourceRows.length > 0) {
    const owner = sheet.collaborators?.find((collaborator) => collaborator.role === "owner");

    try {
      await SheetRow.insertMany(
      sourceRows.map((row, rowIndex) => {
        const hasContent = rowHasProtectedContent(row);

        return {
          sheetId: sheet._id,
          rowIndex,
          cells: normalizeRowCells(row),
          searchText: buildRowSearchText(row),
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
  }

  sheet.data = [];
  sheet.markModified("data");
  await sheet.save();
};

const ensureSheetRowsSearchText = async (sheetId) => {
  const missingCount = await SheetRow.countDocuments({
    sheetId,
    $or: [{ searchText: { $exists: false } }, { searchText: "" }],
  });

  if (missingCount === 0) return;

  const cursor = SheetRow.find({
    sheetId,
    $or: [{ searchText: { $exists: false } }, { searchText: "" }],
  }).cursor();
  const operations = [];

  for await (const row of cursor) {
    operations.push({
      updateOne: {
        filter: { _id: row._id },
        update: { $set: { searchText: buildRowSearchText(row.cells) } },
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

const getRowsForSheet = async (sheetId, start = 0, limit = DEFAULT_ROW_PAGE_SIZE) => {
  const rows = await SheetRow.find({ sheetId, rowIndex: { $gte: start, $lt: start + limit } })
    .sort({ rowIndex: 1 })
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

const ensureSheetRow = async (sheetId, rowIndex) => {
  const row = await SheetRow.findOneAndUpdate(
    { sheetId, rowIndex },
    {
      $setOnInsert: {
        sheetId,
        rowIndex,
        cells: normalizeRowCells([]),
        searchText: "",
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
    .filter((row) => row?.ownerId)
    .map((row) => [
      row.rowIndex,
      {
        userId: row.ownerId.toString(),
        email: row.ownerEmail || "",
        username: row.ownerUsername || row.ownerEmail || "",
      },
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

  const updatedCells = [];
  const updatedRows = [];

  for (const [rowIndex, rowPatches] of patchesByRow.entries()) {
    const editsProtectedColumns = rowPatches.some(
      (patch) => patch.colIndex <= ROW_LOCK_LAST_COLUMN_INDEX
    );
    const row = editsProtectedColumns
      ? await ensureRowEditAccess(sheet, user, role, rowIndex)
      : await ensureSheetRow(sheet._id, rowIndex);

    rowPatches.forEach((patch) => {
      const currentCell = row.cells[patch.colIndex] || createCell("");

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

      updatedCells.push({
        rowIndex,
        colIndex: patch.colIndex,
        cell: row.cells[patch.colIndex],
      });
    });

    row.searchText = buildRowSearchText(row.cells);
    row.markModified("cells");
    row.markModified("searchText");
    await row.save();
    updatedRows.push(row);
  }

  return { updatedCells, rowOwners: getRowOwnershipMap(updatedRows) };
};

const applyCellPatchesForUser = async ({ sheet, user, rowIndex, colIndex, value, formula, patches }) => {
  const normalizedPatches = Array.isArray(patches) && patches.length > 0
    ? patches.slice(0, 50)
    : [{ rowIndex, colIndex, value, formula }];
  const primaryPatch = normalizedPatches[0] || { rowIndex, colIndex, value, formula };
  const numericRowIndex = Number(rowIndex);
  const numericColIndex = Number(colIndex);
  const newValue = primaryPatch.formula || primaryPatch.value || "";
  const role = getUserRole(sheet, user.id);

  await migrateSheetRowsIfNeeded(sheet);

  const editsProtectedColumns = normalizedPatches.some(
    (patch) => Number(patch.colIndex) <= ROW_LOCK_LAST_COLUMN_INDEX
  );
  const oldRow = editsProtectedColumns
    ? await ensureRowEditAccess(sheet, user, role, numericRowIndex)
    : await ensureSheetRow(sheet._id, numericRowIndex);
  const oldCell = oldRow.cells?.[numericColIndex] || "";
  const oldValue =
    oldCell && typeof oldCell === "object" ? oldCell.value : oldCell;

  await createChangeLog({
    sheet,
    user,
    rowIndex: numericRowIndex,
    colIndex: numericColIndex,
    oldValue,
    newValue,
    changeType: primaryPatch.formula || String(primaryPatch.value || "").startsWith("=") ? "formula" : "value",
  });

  const { rowOwners } = await applyCellPatchesToRows(sheet, user, role, normalizedPatches);

  sheet.analytics = {
    ...(sheet.analytics || {}),
    totalEdits: ((sheet.analytics && sheet.analytics.totalEdits) || 0) + 1,
    totalFormulaCells: await countFormulaCellsInRows(sheet._id),
    totalMergedCells: sheet.meta?.merges?.length || 0,
    lastEditedBy: user.email,
    lastEditedAt: new Date(),
  };

  await sheet.save();

  return { patches: normalizedPatches, rowOwners };
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

/* AUTH */

app.post("/signup", authLimiter, async (req, res) => {
  try {
    const email = String(req.body.email || "").toLowerCase().trim();
    const password = String(req.body.password || "");
    const username = normalizeUsername(req.body.username) || await createUniqueUsername(email);

    if (!email || !password) {
      return res.status(400).json({ message: "Email and password are required" });
    }

    if (password.length < 6) {
      return res.status(400).json({ message: "Password must be at least 6 characters" });
    }

    const existingUser = await User.findOne({ email });

    if (existingUser) {
      return res.status(409).json({ message: "User already exists" });
    }

    const hashedPassword = await bcrypt.hash(password, 10);

    const user = await User.create({
      email,
      password: hashedPassword,
      username,
      role: "user",
    });

    const token = signToken(user);

    res.status(201).json({
      message: "User created successfully",
      token,
      user: toUserResponse(user),
    });
  } catch (error) {
    if (error?.code === 11000) {
      return res.status(409).json({ message: "Email or username already exists" });
    }
    res.status(500).json({ message: "Signup failed" });
  }
});

app.post("/login", authLimiter, async (req, res) => {
  try {
    const email = String(req.body.email || "").toLowerCase().trim();
    const password = String(req.body.password || "");

    if (!email || !password) {
      return res.status(400).json({ message: "Email and password are required" });
    }

    const user = await User.findOne({ email });

    if (!user) {
      return res.status(401).json({ message: "Invalid email or password" });
    }

    const isPasswordCorrect = await bcrypt.compare(password, user.password);

    if (!isPasswordCorrect) {
      return res.status(401).json({ message: "Invalid email or password" });
    }

    const token = signToken(user);

    res.json({
      message: "Login successful",
      token,
      user: toUserResponse(user),
    });
  } catch (error) {
    res.status(500).json({ message: "Login failed" });
  }
});

app.get("/me", auth, async (req, res) => {
  const user = await User.findById(req.user.id);

  if (!user) {
    return res.status(404).json({ message: "User not found" });
  }

  res.json({ user: toUserResponse(user) });
});

app.patch("/me", auth, async (req, res) => {
  try {
    const user = await User.findById(req.user.id);

    if (!user) {
      return res.status(404).json({ message: "User not found" });
    }

    const email = String(req.body.email || user.email).toLowerCase().trim();
    const username = normalizeUsername(req.body.username) || user.username || getDefaultUsername(email);
    const avatarUrl = normalizeAvatarUrl(req.body.avatarUrl);

    if (!email || !email.includes("@")) {
      return res.status(400).json({ message: "Valid email is required" });
    }

    if (username.length < 2) {
      return res.status(400).json({ message: "Username must be at least 2 characters" });
    }

    if (avatarUrl === null) {
      return res.status(400).json({ message: "Profile image must be PNG, JPG, or WEBP and under 2 MB" });
    }

    const existingEmail = await User.findOne({ email, _id: { $ne: user._id } });
    if (existingEmail) {
      return res.status(409).json({ message: "Email already exists" });
    }

    const existingUsername = await User.findOne({ username, _id: { $ne: user._id } });
    if (existingUsername) {
      return res.status(409).json({ message: "Username already exists" });
    }

    const oldEmail = user.email;
    const oldUsername = user.username;
    user.email = email;
    user.username = username;
    user.avatarUrl = avatarUrl;
    await user.save();

    if (oldEmail !== email || oldUsername !== username) {
      await Promise.all([
        Sheet.updateMany(
          { "collaborators.userId": user._id },
          {
            $set: {
              "collaborators.$[member].email": email,
              "collaborators.$[member].username": username,
            },
          },
          { arrayFilters: [{ "member.userId": user._id }] }
        ),
        Workspace.updateMany(
          { "members.userId": user._id },
          { $set: { "members.$[member].email": email } },
          { arrayFilters: [{ "member.userId": user._id }] }
        ),
      ]);
    }

    res.json({
      user: toUserResponse(user),
      token: signToken(user),
    });
  } catch (error) {
    res.status(500).json({ message: "Failed to update profile" });
  }
});

app.patch("/me/password", authLimiter, auth, async (req, res) => {
  try {
    const currentPassword = String(req.body.currentPassword || "");
    const newPassword = String(req.body.newPassword || "");

    if (!currentPassword || !newPassword) {
      return res.status(400).json({ message: "Current and new password are required" });
    }

    if (newPassword.length < 6) {
      return res.status(400).json({ message: "Password must be at least 6 characters" });
    }

    const user = await User.findById(req.user.id);

    if (!user) {
      return res.status(404).json({ message: "User not found" });
    }

    const isPasswordCorrect = await bcrypt.compare(currentPassword, user.password);

    if (!isPasswordCorrect) {
      return res.status(401).json({ message: "Current password is incorrect" });
    }

    user.password = await bcrypt.hash(newPassword, 10);
    await user.save();

    res.json({ ok: true });
  } catch (error) {
    res.status(500).json({ message: "Failed to change password" });
  }
});

/* WORKSPACES */

app.post("/workspaces", auth, async (req, res) => {
  try {
    const workspace = await Workspace.create({
      name: req.body.name || "My Workspace",
      ownerId: req.user.id,
      members: [
        {
          userId: req.user.id,
          email: req.user.email,
          role: "admin",
        },
      ],
    });

    res.status(201).json(workspace);
  } catch (error) {
    res.status(500).json({ message: "Failed to create workspace" });
  }
});

app.get("/workspaces", auth, async (req, res) => {
  try {
    const workspaces = await Workspace.find({
      "members.userId": req.user.id,
    }).sort({ updatedAt: -1 });

    res.json(workspaces);
  } catch (error) {
    res.status(500).json({ message: "Failed to get workspaces" });
  }
});

app.post("/workspaces/:id/members", auth, async (req, res) => {
  try {
    const workspace = await Workspace.findById(req.params.id);

    if (!workspace) {
      return res.status(404).json({ message: "Workspace not found" });
    }

    const role = getWorkspaceRole(workspace, req.user.id);

    if (!canManageWorkspace(role)) {
      return res.status(403).json({ message: "Only workspace admin can add members" });
    }

    const email = String(req.body.email || "").toLowerCase().trim();
    const memberRole = req.body.role || "member";

    if (!["admin", "member", "viewer"].includes(memberRole)) {
      return res.status(400).json({ message: "Invalid role" });
    }

    const user = await User.findOne({ email });

    if (!user) {
      return res.status(404).json({ message: "User must sign up first" });
    }

    const existing = workspace.members.find(
      (item) => item.userId.toString() === user._id.toString()
    );

    if (existing) {
      existing.role = memberRole;
      existing.email = user.email;
    } else {
      workspace.members.push({
        userId: user._id,
        email: user.email,
        role: memberRole,
      });
    }

    await workspace.save();

    res.json(workspace);
  } catch (error) {
    res.status(500).json({ message: "Failed to add workspace member" });
  }
});

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
        searchText: buildRowSearchText(row),
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
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canRead(role)) {
      return res.status(404).json({ message: "Sheet not found or access denied" });
    }

    await migrateSheetRowsIfNeeded(sheet);

    if (!sheet.meta) sheet.meta = createDefaultMeta();
    if (!sheet.erpOptions) sheet.erpOptions = createDefaultErpOptions();
    await hydrateCollaboratorUsernames(sheet);

    await sheet.save();

    const totalRows = await getSheetRowCount(req.params.id);
    const rowsResult = await getRowsForSheet(sheet._id, 0, Math.min(rowLimit, totalRows));
    const sheetObject = sheet.toObject();
    sheetObject.data = rowsResult.data;
    sheetObject.rowOwners = rowsResult.rowOwners;

    sendJson(req, res, {
      sheet: sheetObject,
      role,
      rows: {
        start: 0,
        count: rowsResult.data.length,
        total: totalRows,
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
    await ensureSheetRowsSearchText(sheet._id);

    const rowDocs = await SheetRow.find({
      sheetId: sheet._id,
      searchText: { $regex: escapeRegex(query), $options: "i" },
    })
      .sort({ rowIndex: 1 })
      .limit(limit)
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

    sheet.analytics = {
      ...(sheet.analytics || {}),
      totalEdits: ((sheet.analytics && sheet.analytics.totalEdits) || 0) + 1,
      totalFormulaCells: await countFormulaCellsInRows(sheet._id),
      totalMergedCells: sheet.meta?.merges?.length || 0,
      lastEditedBy: req.user.email,
      lastEditedAt: new Date(),
    };

    await sheet.save();

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

    const sheetObject = sheet.toObject();
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

    if (req.body.reset === true) {
      await SheetRow.deleteMany({ sheetId: sheet._id });
    }

    if (rows.length > 0) {
      await SheetRow.bulkWrite(
        rows.map((row, offset) => {
          const hasContent = rowHasProtectedContent(row);

          return {
            updateOne: {
              filter: { sheetId: sheet._id, rowIndex: start + offset },
              update: {
                $set: {
                  sheetId: sheet._id,
                  rowIndex: start + offset,
                  cells: normalizeRowCells(row),
                  searchText: buildRowSearchText(row),
                  ownerId: hasContent ? req.user.id : null,
                  ownerEmail: hasContent ? req.user.email : "",
                  ownerUsername: hasContent ? req.user.username || req.user.email : "",
                },
              },
              upsert: true,
            },
          };
        })
      );
    }

    sheet.data = [];
    sheet.markModified("data");
    sheet.analytics = {
      ...(sheet.analytics || {}),
      totalEdits: ((sheet.analytics && sheet.analytics.totalEdits) || 0) + 1,
      totalFormulaCells: await countFormulaCellsInRows(sheet._id),
      totalMergedCells: sheet.meta?.merges?.length || 0,
      lastEditedBy: req.user.email,
      lastEditedAt: new Date(),
    };

    await sheet.save();

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
      updatedBy: req.user.email,
    });

    res.json({ ok: true, rowOwners: editResult.rowOwners });
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
    row.searchText = buildRowSearchText(row.cells);
    row.markModified("cells");
    row.markModified("searchText");
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

    sheet.analytics = {
      ...(sheet.analytics || {}),
      totalEdits: ((sheet.analytics && sheet.analytics.totalEdits) || 0) + 1,
      totalFormulaCells: await countFormulaCellsInRows(sheet._id),
      totalMergedCells: sheet.meta?.merges?.length || 0,
      lastEditedBy: req.user.email,
      lastEditedAt: new Date(),
    };

    await sheet.save();

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

    const changes = await ChangeLog.find(filter).sort({ createdAt: -1 }).limit(100);

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
        updatedBy: socket.user.email,
      });

      if (typeof ack === "function") ack({ ok: true, rowOwners: editResult.rowOwners });
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

      sheet.analytics = {
        ...(sheet.analytics || {}),
        totalEdits: ((sheet.analytics && sheet.analytics.totalEdits) || 0) + 1,
        totalFormulaCells: await countFormulaCellsInRows(sheet._id),
        totalMergedCells: sheet.meta?.merges?.length || 0,
        lastEditedBy: socket.user.email,
        lastEditedAt: new Date(),
      };

      await sheet.save();

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

    socket.to(sheetId).emit("cursor-change", {
      userId: socket.user.id,
      email: socket.user.email,
      rowIndex,
      colIndex,
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
