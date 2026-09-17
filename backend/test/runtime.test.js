const assert = require("node:assert/strict");
const test = require("node:test");

const {
  createGracefulShutdown,
  getReadiness,
  validateRuntimeEnvironment,
} = require("../utils/runtime");

test("validateRuntimeEnvironment rejects missing required secrets", () => {
  assert.throws(
    () => validateRuntimeEnvironment({ NODE_ENV: "production" }),
    /MONGO_URI/
  );
});

test("validateRuntimeEnvironment enforces production JWT strength and a valid port", () => {
  const baseEnvironment = {
    NODE_ENV: "production",
    MONGO_URI: "mongodb://database.example/sheet-saas",
    FRONTEND_URL: "https://sheets.example.com",
  };

  assert.throws(
    () => validateRuntimeEnvironment({ ...baseEnvironment, JWT_SECRET: "too-short" }),
    /at least 32 characters/
  );
  assert.throws(
    () => validateRuntimeEnvironment({
      ...baseEnvironment,
      JWT_SECRET: "a".repeat(32),
      PORT: "70000",
    }),
    /PORT/
  );
});

test("validateRuntimeEnvironment returns normalized runtime settings", () => {
  const settings = validateRuntimeEnvironment({
    NODE_ENV: "production",
    MONGO_URI: " mongodb://database.example/sheet-saas ",
    JWT_SECRET: "a".repeat(32),
    FRONTEND_URL: "https://sheets.example.com",
    PORT: "8080",
    TRUST_PROXY: "2",
  });

  assert.equal(settings.port, 8080);
  assert.equal(settings.trustProxy, 2);
  assert.equal(settings.mongoUri, "mongodb://database.example/sheet-saas");
});

test("getReadiness only reports ready for a connected database", () => {
  assert.deepEqual(getReadiness(1), { isReady: true, status: "ok" });
  assert.deepEqual(getReadiness(0), { isReady: false, status: "unavailable" });
  assert.deepEqual(getReadiness(2), { isReady: false, status: "unavailable" });
});

test("createGracefulShutdown closes realtime, HTTP, and database resources once", async () => {
  const calls = [];
  const shutdown = createGracefulShutdown({
    closeRealtime: async () => calls.push("realtime"),
    closeHttp: async () => calls.push("http"),
    closeDatabase: async () => calls.push("database"),
    logger: { info: () => {}, error: () => {} },
  });

  await Promise.all([shutdown("SIGTERM"), shutdown("SIGINT")]);

  assert.deepEqual(calls, ["realtime", "http", "database"]);
});
