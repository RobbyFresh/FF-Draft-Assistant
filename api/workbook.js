const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

// Serverless function for Vercel: /api/workbook
// Looks for the Excel file in several common locations within the deployment
module.exports = (req, res) => {
  try {
    const fileName = process.env.EXCEL_FILENAME || 'Fantasy Football Cheat Sheet with Boom Outlier 2025.xlsx';

    // Candidate paths inside the deployed bundle
    const candidates = [
      path.join(process.cwd(), fileName),
      path.join(process.cwd(), 'public', fileName),
      path.join(__dirname, '..', fileName),
      path.join(__dirname, '..', 'public', fileName),
    ];

    const excelPath = candidates.find((p) => {
      try { return fs.existsSync(p); } catch { return false; }
    });

    let workbook;
    if (excelPath) {
      workbook = XLSX.readFile(excelPath, { cellDates: true });
    } else {
      // Try fetching from static hosting (public/) over HTTP on Vercel
      const proto = req.headers['x-forwarded-proto'] || 'https';
      const host = req.headers.host;
      const url = `${proto}://${host}/${encodeURIComponent(fileName)}`;
      const resp = await fetch(url);
      if (!resp.ok) {
        return res.status(404).json({ error: `Excel file not found locally or via HTTP at ${url}. Tried: ${candidates.join(' | ')}` });
      }
      const ab = await resp.arrayBuffer();
      const buf = Buffer.from(ab);
      workbook = XLSX.read(buf, { type: 'buffer', cellDates: true });
    }
    const sheets = workbook.SheetNames.map((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',
        raw: true,
        blankrows: false,
      });
      return { name: sheetName, rows };
    });

    res.setHeader('Cache-Control', 'no-store');
    return res.status(200).json({ fileName: path.basename(excelPath), sheets });
  } catch (err) {
    console.error('Failed to load workbook (vercel api):', err);
    return res.status(500).json({ error: 'Failed to load workbook' });
  }
};


