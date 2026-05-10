require("dotenv").config();

const mongoose = require("mongoose");
const Sheet = require("../models/Sheet");
const SheetRow = require("../models/SheetRow");
const { buildRowSearchText, normalizeRowCells } = require("../utils/sheetRows");

const migrate = async () => {
  if (!process.env.MONGO_URI) {
    throw new Error("MONGO_URI is required");
  }

  await mongoose.connect(process.env.MONGO_URI);

  const sheets = await Sheet.find({ "data.0": { $exists: true } });
  let migratedSheets = 0;
  let migratedRows = 0;

  for (const sheet of sheets) {
    const existingRows = await SheetRow.countDocuments({ sheetId: sheet._id });

    if (existingRows === 0 && sheet.data.length > 0) {
      await SheetRow.insertMany(
        sheet.data.map((row, rowIndex) => ({
          sheetId: sheet._id,
          rowIndex,
          cells: normalizeRowCells(row),
          searchText: buildRowSearchText(row),
        })),
        { ordered: false }
      );
      migratedRows += sheet.data.length;
    }

    sheet.data = [];
    sheet.markModified("data");
    await sheet.save();
    migratedSheets += 1;
  }

  console.log(`Migrated ${migratedRows} rows from ${migratedSheets} sheets.`);
  await mongoose.disconnect();
};

migrate().catch(async (error) => {
  console.error(error);
  await mongoose.disconnect();
  process.exit(1);
});
