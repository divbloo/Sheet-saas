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

const SheetSchema = new mongoose.Schema(
  {
    name: {
      type: String,
      required: true,
    },

    owner: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "User",
      required: true,
      index: true,
    },

    workspaceId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Workspace",
      default: null,
      index: true,
    },

    erpType: {
      type: String,
      default: "custom",
    },

    data: {
      type: [[CellSchema]],
      default: [],
    },

    collaborators: [
      {
        userId: {
          type: mongoose.Schema.Types.ObjectId,
          ref: "User",
        },

        email: String,

        role: {
          type: String,
          enum: ["viewer", "editor"],
          default: "viewer",
        },
      },
    ],

    meta: {
      colWidths: {
        type: Object,
        default: {},
      },

      rowHeights: {
        type: Object,
        default: {},
      },

      merges: {
        type: Array,
        default: [],
      },

      versions: {
        type: Array,
        default: [],
      },
    },

    erpOptions: {
      type: Object,
      default: {},
    },
  },
  {
    timestamps: true,
  }
);

module.exports = mongoose.model("Sheet", SheetSchema);