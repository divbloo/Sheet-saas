const express = require("express");
const bcrypt = require("bcryptjs");

const Sheet = require("../models/Sheet");
const User = require("../models/User");
const Workspace = require("../models/Workspace");
const { authLimiter } = require("../config/security");
const {
  auth,
  getDefaultUsername,
  normalizeAvatarUrl,
  normalizeUsername,
  signToken,
  toUserResponse,
} = require("../utils/userHelpers");

const router = express.Router();

router.get("/me", auth, async (req, res) => {
  const user = await User.findById(req.user.id);

  if (!user) {
    return res.status(404).json({ message: "User not found" });
  }

  res.json({ user: toUserResponse(user) });
});

router.patch("/me", auth, async (req, res) => {
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

router.patch("/me/password", authLimiter, auth, async (req, res) => {
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

module.exports = router;
