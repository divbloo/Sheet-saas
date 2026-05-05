const mongoose = require("mongoose");

const SheetSchema = new mongoose.Schema({
  userId: String,
  name: String,
  data: Array,
  createdAt: { type: Date, default: Date.now },
});

module.exports = mongoose.model("Sheet", SheetSchema);