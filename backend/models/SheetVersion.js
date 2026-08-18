const mongoose = require("mongoose");

const SheetVersionSchema = new mongoose.Schema(
  {
    sheetId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Sheet",
      required: true,
      index: true,
    },
    createdBy: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "User",
      default: null,
    },
    createdByEmail: {
      type: String,
      default: "",
    },
    rowCount: {
      type: Number,
      default: 0,
    },
    sizeBytes: {
      type: Number,
      default: 0,
    },
    legacyKey: {
      type: String,
      unique: true,
      sparse: true,
    },
    snapshot: {
      type: Buffer,
      required: true,
      select: false,
    },
  },
  { timestamps: true }
);

SheetVersionSchema.index({ sheetId: 1, createdAt: -1 });

module.exports = mongoose.model("SheetVersion", SheetVersionSchema);
