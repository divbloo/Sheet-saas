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

test("SheetRow has a searchable token index shape", () => {
  const indexes = SheetRow.schema.indexes();
  const searchIndex = indexes.find(([fields]) => (
    fields.sheetId === 1 &&
    fields.searchTokens === 1 &&
    fields.rowIndex === 1
  ));

  assert.ok(searchIndex);
});

test("SheetRow stores the user who first claimed the row", () => {
  assert.ok(SheetRow.schema.path("ownerId"));
  assert.ok(SheetRow.schema.path("ownerEmail"));
  assert.ok(SheetRow.schema.path("ownerUsername"));
  assert.ok(SheetRow.schema.path("searchTokens"));
  assert.ok(SheetRow.schema.path("hasContent"));
  assert.ok(SheetRow.schema.path("needsCode"));
});

test("SheetRow has a fast pending-code index", () => {
  const indexes = SheetRow.schema.indexes();
  const pendingCodeIndex = indexes.find(([fields]) => (
    fields.sheetId === 1 && fields.needsCode === 1 && fields.rowIndex === 1
  ));

  assert.ok(pendingCodeIndex);
});

test("SheetRow has a fast last-content-row index", () => {
  const indexes = SheetRow.schema.indexes();
  const contentIndex = indexes.find(([fields]) => (
    fields.sheetId === 1 && fields.hasContent === 1 && fields.rowIndex === -1
  ));

  assert.ok(contentIndex);
});
