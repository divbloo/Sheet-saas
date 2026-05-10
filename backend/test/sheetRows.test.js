const test = require("node:test");
const assert = require("node:assert/strict");
const { DEFAULT_SHEET_COLS, buildRowSearchText, normalizeRowCells } = require("../utils/sheetRows");

test("normalizeRowCells pads rows to the configured sheet width", () => {
  const row = normalizeRowCells([{ value: "A" }]);

  assert.equal(row.length, DEFAULT_SHEET_COLS);
  assert.equal(row[0].value, "A");
  assert.equal(row[1].value, "");
});

test("buildRowSearchText includes cell values and formulas", () => {
  const text = buildRowSearchText([{ value: "Item A" }, { formula: "=SUM(A1:A2)" }]);

  assert.match(text, /item a/);
  assert.match(text, /=sum\(a1:a2\)/);
});
