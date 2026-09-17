export const MAX_EXCEL_FILE_SIZE_BYTES = 15 * 1024 * 1024;

const loadExcelJs = async () => {
  const module = await import("exceljs");
  return module.default;
};

export const assertSupportedExcelFile = (file) => {
  if (!file) throw new Error("No Excel file selected");
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    throw new Error("Only .xlsx files are supported");
  }
  if (file.size > MAX_EXCEL_FILE_SIZE_BYTES) {
    throw new Error("Excel file exceeds the 15 MB limit");
  }
};

export const assertSupportedSpreadsheetFile = (file) => {
  if (!file) throw new Error("No spreadsheet file selected");
  const extension = file.name.toLowerCase().match(/\.[^.]+$/)?.[0];
  if (![".xlsx", ".csv"].includes(extension)) {
    throw new Error("Only .xlsx and .csv files are supported");
  }
  if (file.size > MAX_EXCEL_FILE_SIZE_BYTES) {
    throw new Error("Spreadsheet file exceeds the 15 MB limit");
  }
};

export const parseCsv = (source) => {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];

    if (quoted) {
      if (character === '"' && source[index + 1] === '"') {
        cell += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        cell += character;
      }
      continue;
    }

    if (character === '"' && cell === "") {
      quoted = true;
    } else if (character === ",") {
      row.push(cell);
      cell = "";
    } else if (character === "\n" || character === "\r") {
      if (character === "\r" && source[index + 1] === "\n") index += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += character;
    }
  }

  if (cell !== "" || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
};

export const createWorkbook = async (sheetName, rows = []) => {
  const ExcelJS = await loadExcelJs();
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Sheet SaaS";
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet(sheetName);
  worksheet.addRows(rows);
  return { workbook, worksheet };
};

export const loadWorkbook = async (buffer) => {
  const ExcelJS = await loadExcelJs();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
};

const getCellDisplayText = (cell) => {
  const value = cell.value;
  const format = String(cell.numFmt || "");

  if (typeof value === "number" && /^0+$/.test(format)) {
    return String(value).padStart(format.length, "0");
  }
  if (value && typeof value === "object" && "result" in value) {
    return String(value.result ?? "");
  }
  return cell.text || "";
};

export const getWorksheetRows = (worksheet) => {
  if (!worksheet) return [];

  const rows = [];
  for (let rowNumber = 1; rowNumber <= worksheet.actualRowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    rows.push(Array.from(
      { length: row.cellCount },
      (_, columnIndex) => getCellDisplayText(row.getCell(columnIndex + 1))
    ));
  }
  return rows;
};

export const writeWorkbookBuffer = (workbook) => workbook.xlsx.writeBuffer();

export const downloadWorkbook = async (workbook, filename) => {
  const buffer = await writeWorkbookBuffer(workbook);
  const blob = new Blob(
    [buffer],
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
  );
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
};
