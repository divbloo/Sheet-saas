require("dotenv").config();

const mongoose = require("mongoose");
const SheetRow = require("../models/SheetRow");
const {
  buildRowSearchText,
  buildRowSearchTokens,
  compactStoredRowCells,
  rowHasStoredData,
  rowNeedsCode,
} = require("../utils/sheetRows");

const BATCH_SIZE = 500;

const optimize = async () => {
  if (!process.env.MONGO_URI) {
    throw new Error("MONGO_URI is required");
  }

  await mongoose.connect(process.env.MONGO_URI);
  await SheetRow.createIndexes();

  const cursor = SheetRow.find({}).select("cells updatedAt").lean().cursor();
  const operations = [];
  let updatedRows = 0;
  let removedRows = 0;

  const flush = async () => {
    if (operations.length === 0) return;
    await SheetRow.bulkWrite(operations, { ordered: false });
    operations.length = 0;
  };

  for await (const row of cursor) {
    if (!rowHasStoredData(row.cells)) {
      operations.push({
        deleteOne: { filter: { _id: row._id, updatedAt: row.updatedAt } },
      });
      removedRows += 1;
    } else {
      const searchText = buildRowSearchText(row.cells);
      operations.push({
        updateOne: {
          filter: { _id: row._id, updatedAt: row.updatedAt },
          update: {
            $set: {
              cells: compactStoredRowCells(row.cells),
              searchText,
              searchTokens: buildRowSearchTokens(row.cells),
              hasContent: Boolean(searchText),
              needsCode: rowNeedsCode(row.cells),
            },
          },
        },
      });
      updatedRows += 1;
    }

    if (operations.length >= BATCH_SIZE) await flush();
  }

  await flush();
  console.log(`Optimized ${updatedRows} rows and removed ${removedRows} empty rows.`);
  await mongoose.disconnect();
};

optimize().catch(async (error) => {
  console.error(error);
  await mongoose.disconnect();
  process.exit(1);
});
