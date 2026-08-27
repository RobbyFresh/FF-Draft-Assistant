const assert = require('assert');
const handler = require('../api/workbook');

(async () => {
  let statusCode = 0;
  let payload;
  const headers = {};
  const response = {
    setHeader(name, value) { headers[name] = value; },
    status(code) { statusCode = code; return this; },
    json(value) { payload = value; return this; },
  };
  await handler({ headers: { host: 'localhost' } }, response);
  assert.equal(statusCode, 200);
  assert.equal(payload.season, 2026);
  assert.equal(payload.dataVersion, '2026.5');
  assert.equal(payload.fileName, 'Fantasy Football Cheat Sheet with Boom Outlier 2026.xlsx');
  assert.deepEqual(payload.sheets.map((sheet) => sheet.name), ['STRD', '.5 PPR NEW', 'FULL PPR NEW', 'SUPERFLEX NEW']);
  assert.equal(headers['Cache-Control'], 'no-store');
  console.log('Validated Vercel workbook handler metadata and all four sheets.');
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
