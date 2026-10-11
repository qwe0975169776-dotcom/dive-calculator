// 用假的 Google 服務執行真正的 Code.gs，開一個本機 /exec 給網頁測試
const fs = require('fs'), http = require('http'), path = require('path');
function makeEnv() {
  const props = {};
  const sheets = {};
  function makeSheet(name) {
    const grid = [];
    const sh = {
      name, grid,
      getRange(r, c, nr = 1, nc = 1) {
        const rg = {
          setValues(v) { for (let i = 0; i < nr; i++) { grid[r - 1 + i] = grid[r - 1 + i] || []; for (let j = 0; j < nc; j++) grid[r - 1 + i][c - 1 + j] = v[i][j]; } return rg; },
          getValues() { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) { const v = (grid[r - 1 + i] || [])[c - 1 + j]; row.push(v === undefined ? '' : v); } out.push(row); } return out; },
          clear() { for (let i = 0; i < nr; i++) if (grid[r - 1 + i]) for (let j = 0; j < nc; j++) grid[r - 1 + i][c - 1 + j] = ''; return rg; },
        };
        ['setFontWeight', 'setBackground', 'setFontColor', 'setNumberFormat', 'setHorizontalAlignment'].forEach((f) => (rg[f] = () => rg));
        return rg;
      },
      getLastRow() { for (let i = grid.length - 1; i >= 0; i--) if ((grid[i] || []).some((v) => v !== '' && v != null)) return i + 1; return 0; },
      getMaxRows() { return Math.max(1000, grid.length); },
      getMaxColumns() { return 26; },
      setFrozenRows() {},
    };
    return sh;
  }
  const ss = { getSheetByName: (n) => sheets[n] || null, insertSheet: (n) => (sheets[n] = makeSheet(n)), getName: () => '測試試算表', getId: () => 'TEST' };
  const g = {
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); }, setProperties: (o) => Object.assign(props, o) }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    SpreadsheetApp: { openById: () => ss, getActiveSpreadsheet: () => ss },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: (t) => ({ t, setMimeType() { return this; } }) },
  };
  props.SHEET_ID = 'TEST';
  const src = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  const app = new Function(...Object.keys(g), src + '\n;return {doGet, doPost, setup, CODE_VERSION};')(...Object.values(g));
  app.setup();
  return { app, sheets, props };
}
module.exports = { makeEnv };
if (require.main === module) {
  const env = makeEnv();
  const port = Number(process.argv[2] || 8799);
  http.createServer((req, res) => {
    const u = new URL(req.url, 'http://x');
    const send = (out) => { res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }); res.end(out.t); };
    if (u.pathname === '/dump') { res.writeHead(200, { 'Access-Control-Allow-Origin': '*' }); return res.end(JSON.stringify(env.sheets['名單'].grid)); }
    if (req.method === 'GET') return send(env.app.doGet({ parameter: Object.fromEntries(u.searchParams) }));
    let body = ''; req.on('data', (d) => (body += d)); req.on('end', () => send(env.app.doPost({ postData: { contents: body } })));
  }).listen(port, () => console.log('mock on ' + port));
}
