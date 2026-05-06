const mongoose = require("mongoose");

const WorkspaceMemberSchema = new mongoose.Schema(
  {
    userId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "User",
      required: true,
    },
    email: {
      type: String,
      required: true,
      lowercase: true,
      trim: true,
    },
    role: {
      type: String,
      enum: ["admin", "member", "viewer"],
      default: "member",
    },
  },
  { _id: false }
);

const WorkspaceSchema = new mongoose.Schema(
  {
    name: {
      type: String,
      required: true,
      trim: true,
    },
    ownerId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "User",
      required: true,
    },
    members: {
      type: [WorkspaceMemberSchema],
      default: [],
    },
  },
  { timestamps: true }
);

WorkspaceSchema.index({ "members.userId": 1 });

module.exports = mongoose.model("Workspace", WorkspaceSchema);