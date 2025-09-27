const ExcelJS = require("exceljs");
const express = require("express");
const bodyParser = require("body-parser");
const path = require("path");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Serve HTML/CSS/JS from public folder
app.use(express.static(path.join(__dirname, "public")));

const filePath = "./users.xlsx";

// ----- Signup route -----
app.post("/signup", async (req, res) => {
  const { role, username, email, password, extra } = req.body;

  if (!role || !username || !email || !password || !extra) {
    return res.send({ success: false, message: "All fields required" });
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath).catch(() => {}); // create if doesn't exist
  const sheet =
    workbook.getWorksheet("Users") || workbook.addWorksheet("Users");

  if (sheet.rowCount === 0)
    sheet.addRow(["Role", "Username", "Email", "Password", "Extra"]);

  sheet.addRow([role, username, email, password, extra]);
  await workbook.xlsx.writeFile(filePath);

  res.send({ success: true, message: "Signup successful" });
});

// ----- Login route -----
app.post("/login", async (req, res) => {
  const { role, username, password, extra } = req.body;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath).catch(() => {
    res.send({ success: false, message: "No users found" });
    return;
  });
  const sheet = workbook.getWorksheet("Users");
  if (!sheet) return res.send({ success: false, message: "No users found" });

  let userFound = false;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // skip header
    if (
      row.getCell(1).value === role &&
      row.getCell(2).value === username &&
      row.getCell(4).value === password &&
      row.getCell(5).value === extra
    ) {
      userFound = true;
    }
  });

  if (userFound) res.send({ success: true, message: "Login successful" });
  else res.send({ success: false, message: "Invalid credentials" });
});

// ----- Start server -----
app.listen(3000, () => console.log("Server running on port 3000"));
