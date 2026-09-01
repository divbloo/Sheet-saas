import test from "node:test";
import assert from "node:assert/strict";

import {
  EXCEL_IMPORT_OPTIONS,
  getCellClipboardText,
  getSheetTableWidth,
} from "./gridLayout.js";

test("sheet width preserves every configured column instead of shrinking with viewport zoom", () => {
  assert.equal(getSheetTableWidth([300, 220, 150]), 712);
});

test("copying a code preserves its complete text and leading zeros", () => {
  assert.equal(
    getCellClipboardText({ value: "00012345678901234567890", formula: "" }),
    "00012345678901234567890"
  );
  assert.equal(getCellClipboardText({ value: 0, formula: "" }), "0");
});

test("Excel imports use displayed text so item codes are not coerced to numbers", () => {
  assert.deepEqual(EXCEL_IMPORT_OPTIONS, {
    header: 1,
    defval: "",
    raw: false,
  });
});
