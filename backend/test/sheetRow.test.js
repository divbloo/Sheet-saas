const test = require("node:test");
const assert = require("node:assert/strict");
const SheetRow = require("../models/SheetRow");

test("SheetRow has unique per-sheet row index", () => {
  const indexes = SheetRow.schema.indexes();
  const rowIndex = indexes.find(([fields]) => fields.sheetId === 1 && fields.rowIndex === 1);

  assert.ok(rowIndex);
  assert.equal(rowIndex[1].unique, true);
});

test("SheetRow has a searchable text index shape", () => {
  const indexes = SheetRow.schema.indexes();
  const searchIndex = indexes.find(([fields]) => fields.sheetId === 1 && fields.searchText === 1);

  assert.ok(searchIndex);
});
