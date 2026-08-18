require("dotenv").config();

const mongoose = require("mongoose");
const SheetRow = require("../models/SheetRow");
const { buildRowSearchText, buildRowSearchTokens, rowNeedsCode } = require("../utils/sheetRows");

const migrate = async () => {
  if (!process.env.MONGO_URI) {
    throw new Error("MONGO_URI is required");
  }

  await mongoose.connect(process.env.MONGO_URI);

  const cursor = SheetRow.find({})
    .select("cells")
    .cursor();
  const operations = [];
  let migratedRows = 0;

  for await (const row of cursor) {
    const searchText = buildRowSearchText(row.cells);

    operations.push({
      updateOne: {
        filter: { _id: row._id },
        update: {
          $set: {
            searchText,
            searchTokens: buildRowSearchTokens(row.cells),
            hasContent: Boolean(searchText),
            needsCode: rowNeedsCode(row.cells),
          },
        },
      },
    });

    if (operations.length >= 500) {
      await SheetRow.bulkWrite(operations);
      migratedRows += operations.length;
      operations.length = 0;
    }
  }

  if (operations.length > 0) {
    await SheetRow.bulkWrite(operations);
    migratedRows += operations.length;
  }

  console.log(`Backfilled search tokens for ${migratedRows} sheet rows.`);
  await mongoose.disconnect();
};

migrate().catch(async (error) => {
  console.error(error);
  await mongoose.disconnect();
  process.exit(1);
});
