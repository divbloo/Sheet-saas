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
      trim: true,
    },

    createdBy: {
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

    erpTemplate: {
      enabled: {
        type: Boolean,
        default: false,
      },
      type: {
        type: String,
        default: "custom",
      },
      moduleName: {
        type: String,
        default: "",
      },
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
          enum: ["owner", "editor", "viewer"],
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

    analytics: {
      totalEdits: {
        type: Number,
        default: 0,
      },
      totalFormulaCells: {
        type: Number,
        default: 0,
      },
      totalMergedCells: {
        type: Number,
        default: 0,
      },
      lastEditedBy: {
        type: String,
        default: "",
      },
      lastEditedAt: {
        type: Date,
        default: null,
      },
      activeUsers: {
        type: Number,
        default: 0,
      },
    },
  },
  {
    timestamps: true,
  }
);

SheetSchema.index({ "collaborators.userId": 1, updatedAt: -1 });
SheetSchema.index({ workspaceId: 1, updatedAt: -1 });
SheetSchema.index({ createdBy: 1, updatedAt: -1 });

module.exports = mongoose.model("Sheet", SheetSchema);
