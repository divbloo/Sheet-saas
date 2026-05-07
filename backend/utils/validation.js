const mongoose = require("mongoose");

const isValidObjectId = (id) => {
  if (typeof id !== "string" || !mongoose.Types.ObjectId.isValid(id)) {
    return false;
  }

  return new mongoose.Types.ObjectId(id).toString() === id;
};

const isValidCellIndex = (index) => {
  return Number.isInteger(index) && index >= 0 && index < 10000;
};

module.exports = {
  isValidObjectId,
  isValidCellIndex,
};
