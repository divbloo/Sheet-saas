const defaultItemCodingOptions = require("../config/itemDescriptionOptions.json");

const sheetConfigs = [
  { key: "Pipes", aliases: ["المواسير", "pipes"], fieldKeys: ["material", "diameter", "pressure", "uom"] },
  { key: "Fittings", aliases: ["القطع خاصة", "fittings"], fieldKeys: ["material", "diameter", "pressure", "uom"] },
  { key: "Tanks", aliases: ["الخزانات", "tanks"], fieldKeys: ["material", "capacity", "direction", "uom"] },
  { key: "Valves", aliases: ["المحابس", "valves"], fieldKeys: ["material", "diameter", "pressure", "uom"] },
];

const normalizeText = (value) => String(value ?? "").trim();
const normalizeLookupText = (value) => normalizeText(value).replace(/\s+/g, " ").toLowerCase();

const findSheetConfig = (sheet, sheetIndex, usedKeys) => {
  const normalizedName = normalizeLookupText(sheet?.name);
  const namedMatch = sheetConfigs.find(
    (config) =>
      !usedKeys.has(config.key) &&
      config.aliases.some((alias) => normalizeLookupText(alias) === normalizedName)
  );

  if (namedMatch) return namedMatch;

  const orderedConfig = sheetConfigs[sheetIndex];
  return orderedConfig && !usedKeys.has(orderedConfig.key) ? orderedConfig : null;
};

const normalizeOptionsShape = (options = {}) => {
  const normalized = {};

  Object.entries(options || {}).forEach(([key, value]) => {
    const rows = Array.isArray(value?.rows) ? value.rows : [];
    const fields = Array.isArray(value?.fields) ? value.fields : [];

    normalized[key] = {
      label: normalizeText(value?.label) || key,
      fields: fields.map((field) => ({
        key: normalizeText(field.key),
        label: normalizeText(field.label) || normalizeText(field.key),
      })).filter((field) => field.key),
      rows: rows.map((row) => ({
        ...row,
        category: normalizeText(row.category),
        itemName: normalizeText(row.itemName || row.category),
        combination: normalizeText(row.combination),
      })).filter((row) => row.category && row.combination),
    };
  });

  return normalized;
};

const buildOptionsFromSheetRows = (sheets = []) => {
  const options = {};
  const summary = {};
  const usedKeys = new Set();

  sheets.forEach((sheet, sheetIndex) => {
    const config = findSheetConfig(sheet, sheetIndex, usedKeys);
    if (!config) return;
    usedKeys.add(config.key);

    const sourceRows = Array.isArray(sheet.rows) ? sheet.rows : [];
    const headers = (sourceRows[0] || []).map(normalizeText);
    const itemColumn = 0;
    const combinationColumn = 5;
    const fieldColumns = [1, 2, 3, 4];
    const seenCombinations = new Set();
    const rows = [];
    let duplicatesRemoved = 0;

    sourceRows.slice(1).forEach((sourceRow) => {
      const itemName = normalizeText(sourceRow[itemColumn]);
      const combination = normalizeText(sourceRow[combinationColumn]);
      if (!itemName || !combination) return;

      if (seenCombinations.has(combination)) {
        duplicatesRemoved += 1;
        return;
      }
      seenCombinations.add(combination);

      const row = {
        category: itemName,
        itemName,
        combination,
      };

      fieldColumns.forEach((columnIndex, fieldIndex) => {
        row[config.fieldKeys[fieldIndex]] = normalizeText(sourceRow[columnIndex]);
      });

      rows.push(row);
    });

    options[config.key] = {
      label: normalizeText(sheet.name) || config.key,
      fields: fieldColumns.map((columnIndex, fieldIndex) => ({
        key: config.fieldKeys[fieldIndex],
        label: normalizeText(headers[columnIndex]) || config.fieldKeys[fieldIndex],
      })),
      rows,
    };

    if (!rows.length) {
      const error = new Error(`Coding sheet "${options[config.key].label}" has no valid rows`);
      error.statusCode = 400;
      throw error;
    }

    summary[options[config.key].label] = {
      key: config.key,
      sourceRows: sourceRows.length,
      rows: rows.length,
      duplicatesRemoved,
    };
  });

  const missing = sheetConfigs
    .map((config) => config.key)
    .filter((key) => !options[key]);

  if (missing.length) {
    const error = new Error("Missing required coding sheets: " + missing.join(", "));
    error.statusCode = 400;
    throw error;
  }

  return { options, summary };
};

module.exports = {
  buildOptionsFromSheetRows,
  defaultItemCodingOptions: normalizeOptionsShape(defaultItemCodingOptions),
  normalizeOptionsShape,
};
