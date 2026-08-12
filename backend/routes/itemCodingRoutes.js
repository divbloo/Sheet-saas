const express = require("express");

const ItemCodingOptions = require("../models/ItemCodingOptions");
const Sheet = require("../models/Sheet");
const { auth } = require("../utils/userHelpers");
const { isValidObjectId } = require("../utils/validation");
const {
  buildOptionsFromSheetRows,
  defaultItemCodingOptions,
  normalizeOptionsShape,
} = require("../utils/itemCodingOptions");

const router = express.Router();

const hasCodingOptionsAccess = async (user, sheetId) => {
  if (user?.role === "admin") return true;
  if (!isValidObjectId(sheetId)) return false;

  return Boolean(
    await Sheet.exists({
      _id: sheetId,
      collaborators: {
        $elemMatch: {
          userId: user.id,
          role: { $in: ["owner", "admin"] },
        },
      },
    })
  );
};

router.get("/item-coding-options", auth, async (_req, res) => {
  try {
    const savedOptions = await ItemCodingOptions.findOne({ key: "default" }).lean();
    const options = savedOptions?.options
      ? normalizeOptionsShape(savedOptions.options)
      : defaultItemCodingOptions;

    res.json({ options, updatedAt: savedOptions?.updatedAt || null });
  } catch (err) {
    console.error("Failed to load item coding options", err);
    res.status(500).json({ message: "Failed to load item coding options" });
  }
});

router.put("/item-coding-options", auth, async (req, res) => {
  try {
    if (!(await hasCodingOptionsAccess(req.user, req.body.sheetId))) {
      return res.status(403).json({ message: "Only sheet owners/admins can update item coding options" });
    }

    const { options, summary } = buildOptionsFromSheetRows(req.body.sheets || []);
    const savedOptions = await ItemCodingOptions.findOneAndUpdate(
      { key: "default" },
      {
        key: "default",
        options,
        uploadedBy: req.user.id,
        uploadedByEmail: req.user.email || "",
      },
      { new: true, upsert: true, setDefaultsOnInsert: true }
    ).lean();

    res.json({ options: normalizeOptionsShape(savedOptions.options), summary, updatedAt: savedOptions.updatedAt });
  } catch (err) {
    console.error("Failed to update item coding options", err);
    res.status(err.statusCode || 500).json({
      message: err.statusCode ? err.message : "Failed to update item coding options",
    });
  }
});

module.exports = router;
