## Fantasy Football Workbook - Single Server Website

This project serves a local website that mirrors the contents of your Excel workbook `Fantasy Football Cheat Sheet with Boom Outlier 2025.xlsx`. Each worksheet is shown as a tab with a sortable table.

### Prerequisites

- **Node.js** (LTS recommended). If you don't have it, download from the official site.

### Installation

1. Place this folder (including the Excel file) on your machine. Ensure the Excel file is in the project root and named exactly:

   `Fantasy Football Cheat Sheet with Boom Outlier 2025.xlsx`

   If you need to use a different file name, update `EXCEL_FILENAME` in `server.js` accordingly.

2. Open a terminal in the project folder:

   - Windows PowerShell:

     ```powershell
     cd "C:\Users\Rober\Draft Wizard"
     ```

3. Install dependencies:

   ```bash
   npm install
   ```

### Run the server

```bash
npm start
```

Then open `http://localhost:3000` in your browser.

### How it works

- The server uses `xlsx` to read the workbook on demand when you open the page.
- The frontend fetches `/api/workbook` and renders each sheet as a table with clickable headers for sorting.
- To refresh data after changing the Excel file, just reload the page.

### Configuration

- Port: set `PORT` environment variable to change the default (3000). Example:

  ```bash
  $env:PORT=4000; npm start  # PowerShell
  ```

- Excel file name: edit `EXCEL_FILENAME` in `server.js` if you rename the file.

### Troubleshooting

- If you see "Excel file not found", confirm the file exists in the project root and the name matches `EXCEL_FILENAME` in `server.js`.
- If tables look misaligned, ensure there are no merged cells; merged cells are flattened when rendering as a grid.

### Notes

- This is a single-server, single-page site for local use. No database required.


