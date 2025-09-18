const express = require('express');
const fs = require('fs');
const XLSX = require('xlsx');
const path = require('path');
const cors = require('cors');
const app = express();

app.use(cors());
app.use(express.json());

const FILE_PATH = path.join(__dirname, 'visitor_data.xlsx');

app.post('/js/save-visitor-data', (req, res) => {
  const newData = req.body;

  let workbook;
  let worksheet;
  if (fs.existsSync(FILE_PATH)) {
    workbook = XLSX.readFile(FILE_PATH);
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
  } else {
    workbook = XLSX.utils.book_new();
    worksheet = XLSX.utils.json_to_sheet([]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Visitors');
  }

  // Convert worksheet to JSON array
  const data = XLSX.utils.sheet_to_json(worksheet);

  // Append new data
  data.push(newData);

  // Convert back to worksheet
  const newWorksheet = XLSX.utils.json_to_sheet(data);
  workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;

  // Write updated workbook
  XLSX.writeFile(workbook, FILE_PATH);

  res.sendStatus(200);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
