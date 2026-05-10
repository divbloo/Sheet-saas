const mongoose = require("mongoose");

const CellSchema = new mongoose.Schema(
  {
    value: {
      type: mongoose.Schema.Types.Mixed,
      default: "",
    },

    formula: {
      type: String,
      default: "",
    },

    style: {
      type: Object,
      default: {},
    },
  },
  { _id: false }
);

const SheetRowSchema = new mongoose.Schema(
  {
    sheetId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Sheet",
      required: true,
      index: true,
    },

    rowIndex: {
      type: Number,
      required: true,
    },

    cells: {
      type: [CellSchema],
      default: [],
    },

    searchText: {
      type: String,
      default: "",
      index: true,
    },
  },
  {
    timestamps: true,
  }
);

SheetRowSchema.index({ sheetId: 1, rowIndex: 1 }, { unique: true });
SheetRowSchema.index({ sheetId: 1, searchText: 1 });

module.exports = mongoose.model("SheetRow", SheetRowSchema);
