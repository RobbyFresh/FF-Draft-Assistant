const path = require('path');
const fs = require('fs');
const express = require('express');
const helmet = require('helmet');
const XLSX = require('xlsx');

// Configuration
const PORT = process.env.PORT || 3000;
const EXCEL_FILENAME = 'Fantasy Football Cheat Sheet with Boom Outlier 2025.xlsx';
const EXCEL_PATH = path.join(__dirname, EXCEL_FILENAME);

const app = express();

// Security headers; relax CSP to avoid eval-block warnings from extensions/devtools
app.use(helmet({
  contentSecurityPolicy: false,
}));

// Serve static assets
app.use(express.static(path.join(__dirname, 'public')));

// Avoid noisy 404s for favicon requests
app.get('/favicon.ico', (req, res) => res.status(204).end());

// Utility to load workbook into JSON representation
function loadWorkbookToJson() {
  if (!fs.existsSync(EXCEL_PATH)) {
    return { error: `Excel file not found at ${EXCEL_PATH}` };
  }

  const workbook = XLSX.readFile(EXCEL_PATH, { cellDates: true });

  const sheets = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    // header:1 gives us an array-of-arrays, closest to a grid mirror
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      raw: true,
      blankrows: false,
    });
    return { name: sheetName, rows };
  });

  return { fileName: EXCEL_FILENAME, sheets };
}

// API endpoint to fetch workbook JSON
app.get('/api/workbook', (req, res) => {
  try {
    const data = loadWorkbookToJson();
    if (data.error) {
      return res.status(404).json(data);
    }
    res.json(data);
  } catch (err) {
    console.error('Failed to load workbook:', err);
    res.status(500).json({ error: 'Failed to load workbook' });
  }
});

// Fallback to index.html for root
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  const exists = fs.existsSync(EXCEL_PATH);
  console.log(`Server listening on http://localhost:${PORT}`);
  console.log(`Excel file ${exists ? 'found' : 'NOT found'} at: ${EXCEL_PATH}`);
});


