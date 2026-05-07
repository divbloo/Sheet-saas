require("dotenv").config();

const express = require("express");
const fs = require("fs");
const http = require("http");
const path = require("path");
const mongoose = require("mongoose");
const cors = require("cors");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");
const { Server } = require("socket.io");

const User = require("./models/User");
const Sheet = require("./models/Sheet");
const Workspace = require("./models/Workspace");
const ChangeLog = require("./models/ChangeLog");
const {
  authLimiter,
  createCorsOptions,
  createHelmetOptions,
  getFrontendUrls,
  verifyProductionSecurity,
} = require("./config/security");
const { isValidCellIndex, isValidObjectId } = require("./utils/validation");

const app = express();
const server = http.createServer(app);

const PORT = process.env.PORT || 5000;
const FRONTEND_URLS = getFrontendUrls();
const corsOptions = createCorsOptions(FRONTEND_URLS);

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

mongoose
  .connect(process.env.MONGO_URI)
  .then(() => console.log("DB Connected"))
  .catch((err) => console.error("DB Connection Error:", err));

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

const createCell = (value = "") => ({
  value,
  formula: "",
  style: { ...defaultCellStyle },
});

const createEmptySheetData = (rows = 60, cols = 20) => {
  return Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => createCell(""))
  );
};

const ensureCell = (sheet, rowIndex, colIndex) => {
  while (sheet.data.length <= rowIndex) {
    sheet.data.push([]);
  }

  while (sheet.data[rowIndex].length <= colIndex) {
    sheet.data[rowIndex].push(createCell(""));
  }

  const cell = sheet.data[rowIndex][colIndex];

  if (!cell || typeof cell !== "object") {
    sheet.data[rowIndex][colIndex] = createCell(cell || "");
  }

  return sheet.data[rowIndex][colIndex];
};

const applyCellPatch = (sheet, patch) => {
  const rowIndex = Number(patch.rowIndex);
  const colIndex = Number(patch.colIndex);

  if (!isValidCellIndex(rowIndex) || !isValidCellIndex(colIndex)) {
    return null;
  }

  const cell = ensureCell(sheet, rowIndex, colIndex);

  cell.value = patch.value ?? "";
  cell.formula = patch.formula || "";
  cell.style = {
    ...defaultCellStyle,
    ...(cell.style || {}),
    ...(patch.style || {}),
  };

  return { rowIndex, colIndex, cell };
};

const createDefaultMeta = () => ({
  colWidths: {},
  rowHeights: {},
  merges: [],
  versions: [],
});

const createDefaultErpOptions = () => ({
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
});

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
      "العبوة",
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

  const headers = headersByType[type] || headersByType["item-master"];
  const data = createEmptySheetData();

  headers.forEach((header, index) => {
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
    },
    process.env.JWT_SECRET,
    { expiresIn: "7d" }
  );
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

const canRead = (role) => ["owner", "editor", "viewer"].includes(role);
const canEdit = (role) => ["owner", "editor"].includes(role);
const canManage = (role) => role === "owner";

const findSheetForUser = async (sheetId, userId) => {
  const sheet = await Sheet.findById(sheetId);

  if (!sheet) {
    return { sheet: null, role: null };
  }

  const role = getUserRole(sheet, userId);

  return { sheet, role };
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

app.get("/", (req, res) => {
  res.json({ message: "Sheet SaaS API is running" });
});

/* AUTH */

app.post("/signup", authLimiter, async (req, res) => {
  try {
    const email = String(req.body.email || "").toLowerCase().trim();
    const password = String(req.body.password || "");

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
    });

    const token = signToken(user);

    res.status(201).json({
      message: "User created successfully",
      token,
      user: {
        id: user._id,
        email: user.email,
      },
    });
  } catch (error) {
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
      user: {
        id: user._id,
        email: user.email,
      },
    });
  } catch (error) {
    res.status(500).json({ message: "Login failed" });
  }
});

app.get("/me", auth, async (req, res) => {
  res.json({
    user: {
      id: req.user.id,
      email: req.user.email,
    },
  });
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
    const isERP = erpType !== "custom";

    const sheet = await Sheet.create({
      name: req.body.name || "New Sheet",
      workspaceId,
      createdBy: req.user.id,
      data: isERP ? createERPTemplateData(erpType) : createEmptySheetData(),
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
          role: "owner",
        },
      ],
    });

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

    const sheets = await Sheet.find(filter).sort({ updatedAt: -1 });

    res.json(sheets);
  } catch (error) {
    res.status(500).json({ message: "Failed to get sheets" });
  }
});

app.get("/sheet/:id", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canRead(role)) {
      return res.status(404).json({ message: "Sheet not found or access denied" });
    }

    if (!sheet.meta) sheet.meta = createDefaultMeta();
    if (!sheet.erpOptions) sheet.erpOptions = createDefaultErpOptions();

    await sheet.save();

    res.json({ sheet, role });
  } catch (error) {
    res.status(500).json({ message: "Failed to get sheet" });
  }
});

app.put("/sheet/:id", auth, async (req, res) => {
  try {
    const { sheet, role } = await findSheetForUser(req.params.id, req.user.id);

    if (!sheet || !canEdit(role)) {
      return res.status(403).json({ message: "You do not have edit permission" });
    }

    sheet.data = req.body.data || sheet.data;
    sheet.meta = req.body.meta || sheet.meta || createDefaultMeta();

    sheet.analytics = {
      ...(sheet.analytics || {}),
      totalEdits: ((sheet.analytics && sheet.analytics.totalEdits) || 0) + 1,
      totalFormulaCells: countFormulaCells(sheet.data),
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

    io.to(req.params.id).emit("sheet-saved", sheet);

    res.json({ ok: true, sheet });
  } catch (error) {
    res.status(500).json({ message: "Failed to update sheet" });
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

    if (!sheet || !canManage(role)) {
      return res.status(403).json({ message: "Only owner can share sheet" });
    }

    const email = String(req.body.email || "").toLowerCase().trim();
    const newRole = String(req.body.role || "viewer");

    if (!email || !["editor", "viewer"].includes(newRole)) {
      return res.status(400).json({ message: "Valid email and role are required" });
    }

    const userToShare = await User.findOne({ email });

    if (!userToShare) {
      return res.status(404).json({ message: "User must sign up before sharing" });
    }

    const existing = sheet.collaborators.find(
      (item) => item.userId.toString() === userToShare._id.toString()
    );

    if (existing) {
      if (existing.role === "owner") {
        return res.status(400).json({ message: "Owner role cannot be changed" });
      }

      existing.role = newRole;
      existing.email = userToShare.email;
    } else {
      sheet.collaborators.push({
        userId: userToShare._id,
        email: userToShare.email,
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

    if (!sheet || !canManage(role)) {
      return res.status(403).json({ message: "Only owner can remove collaborators" });
    }

    sheet.collaborators = sheet.collaborators.filter(
      (item) =>
        item.userId.toString() !== req.params.userId ||
        item.role === "owner"
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

      if (!isValidCellIndex(Number(rowIndex)) || !isValidCellIndex(Number(colIndex))) {
        socket.emit("socket-error", "Invalid cell position");
        if (typeof ack === "function") ack({ ok: false, message: "Invalid cell position" });
        return;
      }

      const normalizedPatches = Array.isArray(patches) && patches.length > 0
        ? patches.slice(0, 50)
        : [{ rowIndex, colIndex, value, formula }];
      const numericRowIndex = Number(rowIndex);
      const numericColIndex = Number(colIndex);

      const oldCell = sheet.data?.[rowIndex]?.[colIndex] || "";
      const oldValue =
        oldCell && typeof oldCell === "object" ? oldCell.value : oldCell;

      await createChangeLog({
        sheet,
        user: socket.user,
        rowIndex: numericRowIndex,
        colIndex: numericColIndex,
        oldValue,
        newValue: formula || value,
        changeType: formula || String(value).startsWith("=") ? "formula" : "value",
      });

      normalizedPatches.forEach((patch) => applyCellPatch(sheet, patch));
      sheet.markModified("data");

      sheet.analytics = {
        ...(sheet.analytics || {}),
        totalEdits: ((sheet.analytics && sheet.analytics.totalEdits) || 0) + 1,
        totalFormulaCells: countFormulaCells(sheet.data),
        totalMergedCells: sheet.meta?.merges?.length || 0,
        lastEditedBy: socket.user.email,
        lastEditedAt: new Date(),
      };

      await sheet.save();

      socket.to(sheetId).emit("cell-change", {
        rowIndex,
        colIndex,
        value,
        formula,
        patches: normalizedPatches,
        updatedBy: socket.user.email,
      });

      if (typeof ack === "function") ack({ ok: true });
    } catch (error) {
      socket.emit("socket-error", "Failed to update cell");
      if (typeof ack === "function") ack({ ok: false, message: "Failed to update cell" });
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

      const cell = ensureCell(sheet, Number(rowIndex), Number(colIndex));
      const previousStyle = { ...(cell.style || {}) };
      cell.style = {
        ...defaultCellStyle,
        ...previousStyle,
        ...(style || {}),
      };
      sheet.markModified("data");

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
        totalFormulaCells: countFormulaCells(sheet.data),
        totalMergedCells: sheet.meta?.merges?.length || 0,
        lastEditedBy: socket.user.email,
        lastEditedAt: new Date(),
      };

      await sheet.save();

      socket.to(sheetId).emit("cell-style-change", {
        rowIndex,
        colIndex,
        style,
        updatedBy: socket.user.email,
      });

      if (typeof ack === "function") ack({ ok: true });
    } catch (error) {
      socket.emit("socket-error", "Failed to update cell style");
      if (typeof ack === "function") ack({ ok: false, message: "Failed to update cell style" });
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

server.listen(PORT, () => {
  console.log("SaaS API running on port " + PORT);
});
