const test = require("node:test");
const assert = require("node:assert/strict");
const mongoose = require("mongoose");
const { isValidCellIndex, isValidObjectId } = require("../utils/validation");

test("isValidObjectId accepts canonical Mongo ObjectIds", () => {
  const id = new mongoose.Types.ObjectId().toString();
  assert.equal(isValidObjectId(id), true);
});

test("isValidObjectId rejects malformed and non-canonical ids", () => {
  assert.equal(isValidObjectId("abc"), false);
  assert.equal(isValidObjectId("507f1f77bcf86cd799439011".toUpperCase()), false);
  assert.equal(isValidObjectId(null), false);
});

test("isValidCellIndex accepts bounded non-negative integers", () => {
  assert.equal(isValidCellIndex(0), true);
  assert.equal(isValidCellIndex(9999), true);
  assert.equal(isValidCellIndex(999999), true);
});

test("isValidCellIndex rejects unsafe indexes", () => {
  assert.equal(isValidCellIndex(-1), false);
  assert.equal(isValidCellIndex(1000000), false);
  assert.equal(isValidCellIndex(1.5), false);
  assert.equal(isValidCellIndex("1"), false);
});
