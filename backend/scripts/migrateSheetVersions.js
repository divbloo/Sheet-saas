require("dotenv").config();

const mongoose = require("mongoose");
const zlib = require("node:zlib");
const { promisify } = require("node:util");
const Sheet = require("../models/Sheet");
const SheetVersion = require("../models/SheetVersion");
const { compactRowCells, rowHasStoredData } = require("../utils/sheetRows");

const gzipAsync = promisify(zlib.gzip);
const MAX_VERSION_SNAPSHOT_BYTES = 12 * 1024 * 1024;

const sanitizeMeta = (meta = {}) => ({
  colWidths: meta?.colWidths || {},
  rowHeights: meta?.rowHeights || {},
  merges: Array.isArray(meta?.merges) ? meta.merges : [],
  versions: [],
});

const migrate = async () => {
  if (!process.env.MONGO_URI) throw new Error("MONGO_URI is required");
  await mongoose.connect(process.env.MONGO_URI);

  const cursor = Sheet.find({ "meta.versions.0": { $exists: true } })
    .select("createdBy collaborators meta.versions")
    .cursor();
  let migratedSheets = 0;
  let migratedVersions = 0;
  let compressedBytes = 0;

  for await (const sheet of cursor) {
    const versions = Array.isArray(sheet.meta?.versions) ? sheet.meta.versions : [];
    const owner = sheet.collaborators?.find((collaborator) => collaborator.role === "owner");

    for (let index = 0; index < versions.length; index += 1) {
      const version = versions[index]?.toObject?.() || versions[index] || {};
      const sourceRows = Array.isArray(version.data) ? version.data : [];
      const rows = sourceRows
        .map((cells, rowIndex) => ({ rowIndex, cells }))
        .filter(({ cells }) => rowHasStoredData(cells))
        .map(({ rowIndex, cells }) => ({ rowIndex, cells: compactRowCells(cells) }));
      const snapshot = await gzipAsync(Buffer.from(JSON.stringify({
        meta: sanitizeMeta(version.meta),
        rows,
      })));

      if (snapshot.length > MAX_VERSION_SNAPSHOT_BYTES) {
        throw new Error(`Version ${index + 1} for sheet ${sheet._id} exceeds the snapshot limit`);
      }

      const createdAt = version.createdAt ? new Date(version.createdAt) : new Date();
      const legacyKey = `${sheet._id}:${createdAt.toISOString()}:${index}`;
      const result = await SheetVersion.updateOne(
        { legacyKey },
        {
          $setOnInsert: {
            sheetId: sheet._id,
            createdBy: sheet.createdBy || null,
            createdByEmail: owner?.email || "",
            rowCount: rows.length,
            sizeBytes: snapshot.length,
            snapshot,
            createdAt,
            updatedAt: createdAt,
            legacyKey,
          },
        },
        { upsert: true, timestamps: false }
      );
      if (result.upsertedCount > 0) migratedVersions += 1;
      compressedBytes += snapshot.length;
    }

    await Sheet.updateOne(
      { _id: sheet._id },
      { $set: { "meta.versions": [] } }
    );
    migratedSheets += 1;
    console.log(`Migrated ${versions.length} versions for sheet ${sheet._id}.`);
  }

  console.log(
    `Migrated ${migratedVersions} versions from ${migratedSheets} sheets ` +
    `(${(compressedBytes / 1024 / 1024).toFixed(2)} MB compressed).`
  );
  await mongoose.disconnect();
};

migrate().catch(async (error) => {
  console.error(error);
  await mongoose.disconnect();
  process.exit(1);
});
