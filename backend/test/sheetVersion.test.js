const test = require("node:test");
const assert = require("node:assert/strict");
const SheetVersion = require("../models/SheetVersion");

test("SheetVersion keeps compressed snapshots out of normal projections", () => {
  assert.equal(SheetVersion.schema.path("snapshot").options.select, false);
});

test("SheetVersion has a per-sheet newest-first index", () => {
  const indexes = SheetVersion.schema.indexes();
  const versionIndex = indexes.find(([fields]) => (
    fields.sheetId === 1 && fields.createdAt === -1
  ));

  assert.ok(versionIndex);
});
