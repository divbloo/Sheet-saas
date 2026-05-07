const rateLimit = require("express-rate-limit");

const getFrontendUrls = () => {
  return (process.env.FRONTEND_URL || "http://localhost:5173")
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean);
};

const createCorsOptions = (allowedOrigins) => ({
  origin(origin, callback) {
    if (!origin || allowedOrigins.includes(origin)) {
      return callback(null, true);
    }

    return callback(new Error("Not allowed by CORS"));
  },
});

const createHelmetOptions = () => ({
  contentSecurityPolicy: process.env.NODE_ENV === "production" ? undefined : false,
});

const authLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 25,
  message: { message: "Too many authentication attempts, please try again later" },
});

const verifyProductionSecurity = () => {
  if (process.env.NODE_ENV === "production" && process.env.JWT_SECRET.length < 32) {
    console.error("JWT_SECRET must be at least 32 characters in production");
    process.exit(1);
  }
};

module.exports = {
  authLimiter,
  createCorsOptions,
  createHelmetOptions,
  getFrontendUrls,
  verifyProductionSecurity,
};
