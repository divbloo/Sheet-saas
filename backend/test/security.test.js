const test = require("node:test");
const assert = require("node:assert/strict");
const { getFrontendUrls } = require("../config/security");

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
