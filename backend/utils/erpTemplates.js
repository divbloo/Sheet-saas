const { DEFAULT_SHEET_COLS, createCell, defaultCellStyle } = require("./sheetRows");
const { DEFAULT_SHEET_ROWS } = require("../config/sheetConstants");

const createEmptySheetData = (rows = DEFAULT_SHEET_ROWS, cols = DEFAULT_SHEET_COLS) => {
  return Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => createCell(""))
  );
};

const visibleItemMasterHeaders = [
  "Ø§Ø³Ù… Ø§Ù„ØµÙ†Ù",
  "Ø§Ù„ÙˆØµÙ",
  "Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹Ø© Ø§Ù„Ø±Ø¦ÙŠØ³ÙŠØ©",
  "Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹Ø© Ø§Ù„ÙØ±Ø¹ÙŠØ©",
  "Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹Ø© ØªØ­Øª Ø§Ù„ÙØ±Ø¹ÙŠØ©",
  "Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹Ø© Ø§Ù„Ù…Ø³Ø§Ø¹Ø¯Ø©",
  "Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹Ø© Ø§Ù„ØªÙØµÙŠÙ„ÙŠØ©",
  "ÙˆØ­Ø¯Ø© Ø§Ù„Ù‚ÙŠØ§Ø³",
  "Ø§Ù„ØµÙ„Ø§Ø­ÙŠØ©",
  "Ø§Ù„ØªØ³Ù„Ø³Ù„",
  "Ù…Ù„Ø§Ø­Ø¸Ø§Øª",
  "Ø§Ù„ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ø£ÙˆÙ„",
  "Ø§Ù„ÙƒÙˆØ¯",
  "Ø§Ù„ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ø«Ø§Ù†ÙŠ",
  "Modified Description",
];

const headersByType = {
  "item-master": [
    "Ø§Ø³Ù… Ø§Ù„ØµÙ†Ù",
    "Ø§Ù„ÙˆØµÙ",
    "Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹Ø© Ø§Ù„Ø±Ø¦ÙŠØ³ÙŠØ©",
    "Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹Ø© Ø§Ù„ÙØ±Ø¹ÙŠØ©",
    "Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹Ø© ØªØ­Øª Ø§Ù„ÙØ±Ø¹ÙŠØ©",
    "Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹Ø© Ø§Ù„Ù…Ø³Ø§Ø¹Ø¯Ø©",
    "Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹Ø© Ø§Ù„ØªÙØµÙŠÙ„ÙŠØ©",
    "ÙˆØ­Ø¯Ø© Ø§Ù„Ù‚ÙŠØ§Ø³",
    "Ø§Ù„ØµÙ„Ø§Ø­ÙŠØ©",
    "Ø§Ù„ØªØ³Ù„Ø³Ù„",
    "Ù…Ù„Ø§Ø­Ø¸Ø§Øª",
    "Ø§Ù„ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ø£ÙˆÙ„",
    "Ø§Ù„ÙƒÙˆØ¯",
    "Ø§Ù„ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ø«Ø§Ù†ÙŠ",
    "Ø§Ù„ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ø«Ø§Ù„Ø«",
  ],
  inventory: [
    "Item Code",
    "Item Name",
    "Category",
    "Warehouse",
    "Opening Qty",
    "In",
    "Out",
    "Balance",
    "Min Stock",
    "Status",
  ],
  sales: [
    "Date",
    "Customer",
    "Sales Person",
    "Brand",
    "Item",
    "Qty",
    "Unit Price",
    "Total",
    "Status",
    "Notes",
  ],
  finance: [
    "Date",
    "Account",
    "Description",
    "Debit",
    "Credit",
    "Balance",
    "Cost Center",
    "Status",
  ],
  purchasing: [
    "PR No.",
    "Supplier",
    "Item",
    "Qty",
    "Requested By",
    "Unit Price",
    "Delivery Date",
    "Status",
  ],
  hr: [
    "Employee ID",
    "Name",
    "Department",
    "Position",
    "Join Date",
    "Salary",
    "Leave Balance",
    "Status",
  ],
  crm: [
    "Lead",
    "Company",
    "Contact",
    "Phone",
    "Email",
    "Stage",
    "Next Action",
    "Owner",
  ],
  offers: [
    "Offer No.",
    "Customer",
    "Subject",
    "Item",
    "Qty",
    "Cost",
    "Selling Price",
    "Margin",
    "Approval",
    "Status",
  ],
  warehouse: [
    "Transaction Date",
    "Item Code",
    "Item Name",
    "Location",
    "Received",
    "Issued",
    "Balance",
    "Handled By",
  ],
};

const createERPTemplateData = (type) => {
  const headers = type === "item-master"
    ? visibleItemMasterHeaders
    : headersByType[type] || headersByType["item-master"];
  const data = createEmptySheetData(1);

  headers.slice(0, DEFAULT_SHEET_COLS).forEach((header, index) => {
    data[0][index] = {
      value: header,
      formula: "",
      style: {
        ...defaultCellStyle,
        fontWeight: "bold",
        backgroundColor: "#ccfbf1",
        color: "#075985",
        textAlign: "center",
      },
    };
  });

  return data;
};

module.exports = {
  createERPTemplateData,
};
