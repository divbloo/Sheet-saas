const test = require("node:test");
const assert = require("node:assert/strict");
const { createHelmetOptions, getFrontendUrls } = require("../config/security");

test("getFrontendUrls supports comma-separated origins", () => {
  const previous = process.env.FRONTEND_URL;
  process.env.FRONTEND_URL = "http://localhost:5173, https://app.example.com ";

  assert.deepEqual(getFrontendUrls(), [
    "http://localhost:5173",
    "https://app.example.com",
  ]);

  if (previous === undefined) {
    delete process.env.FRONTEND_URL;
  } else {
    process.env.FRONTEND_URL = previous;
  }
});

test("createHelmetOptions does not force HTTPS upgrades by default", () => {
  const previousNodeEnv = process.env.NODE_ENV;
  const previousForceHttps = process.env.FORCE_HTTPS;

  process.env.NODE_ENV = "production";
  delete process.env.FORCE_HTTPS;

  const options = createHelmetOptions();

  assert.equal(options.contentSecurityPolicy.directives.upgradeInsecureRequests, null);

  if (previousNodeEnv === undefined) {
    delete process.env.NODE_ENV;
  } else {
    process.env.NODE_ENV = previousNodeEnv;
  }

  if (previousForceHttps === undefined) {
    delete process.env.FORCE_HTTPS;
  } else {
    process.env.FORCE_HTTPS = previousForceHttps;
  }
});
