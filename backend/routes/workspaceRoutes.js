const express = require("express");

const User = require("../models/User");
const Workspace = require("../models/Workspace");
const { auth } = require("../utils/userHelpers");
const { canManageWorkspace, getWorkspaceRole } = require("../utils/permissions");

const router = express.Router();

router.post("/workspaces", auth, async (req, res) => {
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

router.get("/workspaces", auth, async (req, res) => {
  try {
    const workspaces = await Workspace.find({
      "members.userId": req.user.id,
    }).sort({ updatedAt: -1 });

    res.json(workspaces);
  } catch (error) {
    res.status(500).json({ message: "Failed to get workspaces" });
  }
});

router.post("/workspaces/:id/members", auth, async (req, res) => {
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

module.exports = router;
