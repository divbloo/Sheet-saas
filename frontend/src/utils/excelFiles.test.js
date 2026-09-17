import assert from "node:assert/strict";
import test from "node:test";

import {
  createWorkbook,
  getWorksheetRows,
  loadWorkbook,
  parseCsv,
  writeWorkbookBuffer,
} from "./excelFiles.js";

test("Excel helpers preserve displayed item codes with leading zeros", async () => {
  const { workbook, worksheet } = await createWorkbook("Codes", [[123]]);
  worksheet.getCell(1, 1).numFmt = "000000";

  const loaded = await loadWorkbook(await writeWorkbookBuffer(workbook));
  const rows = getWorksheetRows(loaded.getWorksheet("Codes"));

  assert.equal(rows[0][0], "000123");
});

test("Excel helpers preserve empty cells between populated columns", async () => {
  const { worksheet } = await createWorkbook("Sheet1", [["A", "", "C"]]);

  assert.deepEqual(getWorksheetRows(worksheet), [["A", "", "C"]]);
});

test("parseCsv preserves quoted commas, quotes, and leading zeros", () => {
  assert.deepEqual(
    parseCsv('Code,Description\r\n"000123","Valve, 2"" brass"'),
    [["Code", "Description"], ["000123", 'Valve, 2" brass']]
  );
});
