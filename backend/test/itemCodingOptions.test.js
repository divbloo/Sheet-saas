const assert = require("node:assert/strict");
const test = require("node:test");

const { buildOptionsFromSheetRows } = require("../utils/itemCodingOptions");

const makeRows = (name, item = "PVC PIPE") => ({
  name,
  rows: [
    ["ITEM", "MATERIAL", "DIAMETER", "PRESSURE", "UOM", "COMBINATION"],
    [item, "PVC", "110MM", "PN10", "M", `${item} PVC 110MM PN10 M`],
    [item, "PVC", "110MM", "PN10", "M", `${item} PVC 110MM PN10 M`],
  ],
});

test("buildOptionsFromSheetRows normalizes coding workbook sheets", () => {
  const { options, summary } = buildOptionsFromSheetRows([
    makeRows("Pipes"),
    makeRows("Fittings", "ELBOW"),
    makeRows("Tanks", "TANK"),
    makeRows("Valves", "VALVE"),
  ]);

  assert.equal(options.Pipes.label, "Pipes");
  assert.equal(options.Pipes.fields[0].key, "material");
  assert.equal(options.Pipes.rows.length, 1);
  assert.equal(options.Pipes.rows[0].combination, "PVC PIPE PVC 110MM PN10 M");
  assert.equal(summary.Pipes.duplicatesRemoved, 1);
});

test("buildOptionsFromSheetRows rejects missing required coding sheets", () => {
  assert.throws(
    () => buildOptionsFromSheetRows([makeRows("Pipes")]),
    /Missing required coding sheets/
  );
});

test("buildOptionsFromSheetRows matches coding sheets by name before order", () => {
  const { options } = buildOptionsFromSheetRows([
    makeRows("الخزانات", "TANK"),
    makeRows("المحابس", "VALVE"),
    makeRows("المواسير", "PIPE"),
    makeRows("القطع خاصة", "ELBOW"),
  ]);

  assert.equal(options.Tanks.rows[0].category, "TANK");
  assert.equal(options.Valves.rows[0].category, "VALVE");
  assert.equal(options.Pipes.rows[0].category, "PIPE");
  assert.equal(options.Fittings.rows[0].category, "ELBOW");
});
