const express = require('express');
const app = express();
const cors = require('cors');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');

const excelFilePath = './student-database.xlsx'; // Ensure this path is correct and the file exists

app.use(cors());
app.use(bodyParser.json());

// Route to handle adding student data
app.post('/api/addStudent', (req, res) => {
  const { barcode, password } = req.body;

  console.log("Received data to save:", { barcode, password });

  // Check if both fields are provided
  if (!barcode || !password) {
    console.error("Error: Barcode or password is missing.");
    return res.status(400).json({ message: "Barcode and password are required." });
  }

  try {
    // Ensure the Excel file exists
    if (!fs.existsSync(excelFilePath)) {
      console.error("Error: Excel file does not exist at", excelFilePath);
      return res.status(500).json({ message: "Excel file not found." });
    }

    // Read the existing Excel file
    const workbook = XLSX.readFile(excelFilePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(worksheet);

    console.log("Data read from Excel:", data);

    // Check if the barcode already exists in Excel
    const existingStudent = data.find(student => student.barcode === barcode);
    if (existingStudent) {
      console.error("Error: Barcode already exists in Excel.");
      return res.status(400).json({ message: "Barcode already exists." });
    }

    // Create a new student entry
    const newStudent = { barcode, password };
    data.push(newStudent);

    // Convert back to worksheet and save to Excel
    const newWorksheet = XLSX.utils.json_to_sheet(data);
    workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;

    console.log("Saving data to Excel...");

    XLSX.writeFile(workbook, excelFilePath);
    console.log("Data saved successfully:", newStudent);

    res.status(200).json({ message: "Data saved successfully." });
  } catch (error) {
    console.error("Error saving data to Excel:", error);
    res.status(500).json({ message: "Failed to save data." });
  }
});

// Route to validate password
app.post('/api/validatePassword', (req, res) => {
  const { barcode, password } = req.body;

  try {
    if (!fs.existsSync(excelFilePath)) {
      console.error("Error: Excel file does not exist at", excelFilePath);
      return res.status(500).json({ message: "Excel file not found." });
    }

    const workbook = XLSX.readFile(excelFilePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(worksheet);

    console.log("Data read for validation:", data);

    const student = data.find(student => student.barcode === barcode);
    if (student && student.password === password) {
      return res.status(200).json({ message: "Password validated." });
    } else {
      return res.status(400).json({ message: "Invalid password." });
    }
  } catch (error) {
    console.error("Error validating password:", error);
    res.status(500).json({ message: "Error validating password." });
  }
});

app.listen(5000, () => {
  console.log("Server is running on http://localhost:5000");
});
