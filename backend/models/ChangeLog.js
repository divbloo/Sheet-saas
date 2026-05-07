const mongoose = require("mongoose");

const ChangeLogSchema = new mongoose.Schema(
  {
    sheetId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Sheet",
      required: true,
      index: true,
    },
    workspaceId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Workspace",
      default: null,
    },
    userId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "User",
      required: true,
    },
    userEmail: {
      type: String,
      required: true,
    },
    rowIndex: {
      type: Number,
      required: true,
    },
    colIndex: {
      type: Number,
      required: true,
    },
    cellAddress: {
      type: String,
      required: true,
    },
    oldValue: {
      type: mongoose.Schema.Types.Mixed,
      default: "",
    },
    newValue: {
      type: mongoose.Schema.Types.Mixed,
      default: "",
    },
    changeType: {
      type: String,
      enum: ["value", "style", "merge", "resize", "import", "formula"],
      default: "value",
    },
  },
  { timestamps: true }
);

ChangeLogSchema.index({ sheetId: 1, createdAt: -1 });
ChangeLogSchema.index({ workspaceId: 1, createdAt: -1 });
ChangeLogSchema.index({ userId: 1, createdAt: -1 });
ChangeLogSchema.index({ sheetId: 1, rowIndex: 1, colIndex: 1, createdAt: -1 });

module.exports = mongoose.model("ChangeLog", ChangeLogSchema);
