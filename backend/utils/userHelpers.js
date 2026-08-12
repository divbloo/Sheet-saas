const jwt = require("jsonwebtoken");

const User = require("../models/User");

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

module.exports = {
  auth,
  createUniqueUsername,
  getDefaultUsername,
  normalizeAvatarUrl,
  normalizeUsername,
  signToken,
  toUserResponse,
};
