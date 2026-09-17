const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");

const { findCaseSensitiveTarget } = require("./checkImportCase");

test("findCaseSensitiveTarget accepts the exact Linux filename casing", () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "sheet-saas-import-case-"));
  fs.mkdirSync(path.join(root, "components"));
  fs.writeFileSync(path.join(root, "components", "SheetGrid.jsx"), "");

  assert.equal(
    findCaseSensitiveTarget(root, "./components/SheetGrid"),
    path.join(root, "components", "SheetGrid.jsx")
  );
});

test("findCaseSensitiveTarget rejects casing that only works on Windows", () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "sheet-saas-import-case-"));
  fs.mkdirSync(path.join(root, "components"));
  fs.writeFileSync(path.join(root, "components", "SheetGrid.jsx"), "");

  assert.equal(findCaseSensitiveTarget(root, "./components/sheetgrid"), null);
});
