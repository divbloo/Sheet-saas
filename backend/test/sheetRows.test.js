const test = require("node:test");
const assert = require("node:assert/strict");
const {
  DEFAULT_SHEET_COLS,
  buildRowSearchText,
  buildRowSearchTokens,
  buildSearchQueryTokens,
  compactRowCells,
  compactStoredRowCells,
  normalizeRowCells,
  rowHasStoredData,
  rowNeedsCode,
} = require("../utils/sheetRows");

test("normalizeRowCells pads rows to the configured sheet width", () => {
  const row = normalizeRowCells([{ value: "A" }]);

  assert.equal(row.length, DEFAULT_SHEET_COLS);
  assert.equal(row[0].value, "A");
  assert.equal(row[1].value, "");
});

test("normalizeRowCells drops the legacy package column", () => {
  const row = normalizeRowCells(
    Array.from({ length: DEFAULT_SHEET_COLS + 1 }, (_, index) => ({ value: `C${index}` }))
  );

  assert.equal(row.length, DEFAULT_SHEET_COLS);
  assert.equal(row[7].value, "C7");
  assert.equal(row[8].value, "C9");
  assert.equal(row[13].value, "C14");
});

test("compactRowCells removes default styles and trailing empty cells", () => {
  assert.deepEqual(compactRowCells([]), []);
  assert.deepEqual(compactRowCells([{ value: "Item" }]), ["Item"]);
  assert.deepEqual(
    compactRowCells([{ value: "", style: { backgroundColor: "#ff0000" } }]),
    [{ value: "", formula: "", style: { backgroundColor: "#ff0000" } }]
  );
});

test("compactStoredRowCells strips default styles before database writes", () => {
  assert.deepEqual(compactStoredRowCells([]), []);
  assert.deepEqual(
    compactStoredRowCells([{ value: "Item" }]),
    [{ value: "Item", formula: "", style: {} }]
  );
  assert.deepEqual(
    compactStoredRowCells([{ value: "Item", style: { color: "#ff0000" } }]),
    [{ value: "Item", formula: "", style: { color: "#ff0000" } }]
  );
});

test("buildRowSearchText includes cell values and formulas", () => {
  const text = buildRowSearchText([{ value: "Item A" }, { formula: "=SUM(A1:A2)" }]);

  assert.match(text, /item a/);
  assert.match(text, /=sum\(a1:a2\)/);
});


test("buildRowSearchText returns empty text for empty rows", () => {
  const text = buildRowSearchText([]);

  assert.equal(text, "");
});

test("rowHasStoredData ignores default empty cells but keeps custom styles", () => {
  assert.equal(rowHasStoredData([]), false);
  assert.equal(rowHasStoredData([{ value: "Item" }]), true);
  assert.equal(rowHasStoredData([{ style: { backgroundColor: "#ff0000" } }]), true);
});

test("rowNeedsCode only flags confirmed rows without an item code", () => {
  const pending = [];
  pending[11] = { value: "confirmed" };
  assert.equal(rowNeedsCode(pending), true);

  pending[12] = { value: "ITEM-1" };
  assert.equal(rowNeedsCode(pending), false);
  assert.equal(rowNeedsCode([]), false);
});

test("buildRowSearchTokens supports partial indexed search", () => {
  const tokens = buildRowSearchTokens([{ value: "Invoice Total" }]);

  assert.ok(tokens.includes("invoice"));
  assert.ok(tokens.includes("nvo"));
  assert.ok(tokens.includes("to"));
});

test("buildSearchQueryTokens matches row search token shape", () => {
  const rowTokens = buildRowSearchTokens([{ value: "Invoice Total" }]);
  const queryTokens = buildSearchQueryTokens("voi");

  assert.ok(queryTokens.every((token) => rowTokens.includes(token)));
});
