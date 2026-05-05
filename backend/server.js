require("dotenv").config();

const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");

const User = require("./models/User");
const Sheet = require("./models/Sheet");

const app = express();
app.use(cors());
app.use(express.json());

mongoose.connect(process.env.MONGO_URI)
  .then(() => console.log("DB Connected"));

// ==========================
// AUTH (simple version)
// ==========================
app.post("/signup", async (req, res) => {

  const user = await User.create(req.body);

  res.json(user);
});

app.post("/login", async (req, res) => {

  const user = await User.findOne({
    email: req.body.email,
    password: req.body.password,
  });

  res.json(user);
});

// ==========================
// CREATE SHEET
// ==========================
app.post("/sheet", async (req, res) => {

  const sheet = await Sheet.create({
    userId: req.body.userId,
    name: "New Sheet",
    data: Array.from({ length: 20 }, () =>
      Array(10).fill("")
    ),
  });

  res.json(sheet);
});

// ==========================
// GET USER SHEETS
// ==========================
app.get("/sheets/:userId", async (req, res) => {

  const sheets = await Sheet.find({
    userId: req.params.userId,
  });

  res.json(sheets);
});

// ==========================
app.get("/sheet/:id", async (req, res) => {

  const sheet = await Sheet.findById(req.params.id);

  res.json(sheet);
});

// ==========================
app.put("/sheet/:id", async (req, res) => {

  await Sheet.findByIdAndUpdate(req.params.id, {
    data: req.body.data,
  });

  res.json({ ok: true });
});

// ==========================
app.listen(5000, () => {
  console.log("🚀 SaaS API running");
});