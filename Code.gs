// ===== 潛水旅費動態結算器：伺服器程式 =====
// 放在 GitHub，由 Apps Script 的「啟動程式」(apps-script/loader.gs) 自動下載執行。
// 這個檔案是公開的，不能寫試算表 ID；一律從指令碼屬性讀取。
// 快取（CacheService）的鍵 'code' 是啟動程式在用的，這裡不能用。
const CODE_VERSION = 2;

const SHEET_NAME = '名單';

// 次數欄位（試算表從 C 欄開始依序排列）
const COUNT_COLS = [
  ['shore25',   '10/25岸潛支數'],
  ['shore26',   '10/26岸潛支數'],
  ['shore27',   '10/27岸潛支數'],
  ['lodging',   '住宿晚數'],
  ['scooter',   '機車天數'],
  ['insurance', '保險天數(自動)'],
  ['lunch',     '午餐次數'],
  ['breakfast', '早餐次數'],
  ['ferry',     '船票'],
  ['boatDive',  '船潛'],
  ['owCourse',  'OW課程']
];
const CHECK_KEYS = ['ferry', 'boatDive', 'owCourse'];

// 費用欄位（接在次數後面，順序和網頁對帳表相同）
const FEE_COLS = [
  ['owCourse',  'OW課程費'],
  ['shore25',   '10/25岸潛費'],
  ['shore26',   '10/26岸潛費'],
  ['shore27',   '10/27岸潛費'],
  ['boatDive',  '船潛費'],
  ['insurance', '保險費'],
  ['lodging',   '住宿費'],
  ['ferry',     '船票費'],
  ['scooter',   '機車費'],
  ['lunch',     '午餐費'],
  ['breakfast', '早餐費']
];

const HEADERS = ['ID', '姓名']
  .concat(COUNT_COLS.map(function (c) { return c[1]; }))
  .concat(FEE_COLS.map(function (c) { return c[1]; }))
  .concat(['個人總應繳']);

// 讀取時用表頭名稱找欄位（包含舊版表頭，升級前存的資料也讀得回來）
const READ_NAMES = {
  '10/25岸潛支數': 'shore25', '10/26岸潛支數': 'shore26', '10/27岸潛支數': 'shore27',
  '岸潛支數': 'shoreDive',   // 第 1 版：只有總支數，網頁會自己拆成三天
  '住宿晚數': 'lodging', '機車天數': 'scooter',
  '保險天數(自動)': 'insurance', '保險天數': 'insurance',
  '午餐次數': 'lunch', '午餐天數': 'lunch',
  '早餐次數': 'breakfast', '早餐天數': 'breakfast',
  '船票': 'ferry', '船潛': 'boatDive', 'OW課程': 'owCourse'
};

// ---------- 入口 ----------
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'load';
  try {
    if (action === 'load') return json_(load_());
    if (action === 'ping') return json_({ ok: true, version: CODE_VERSION });
    return json_({ ok: false, error: '不明的動作：' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '資料格式錯誤' });
  }
  try {
    if (body.action === 'save') return json_(save_(body));
    return json_({ ok: false, error: '不明的動作：' + body.action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

// 啟動程式的「一次設定」會呼叫這裡
function setup() {
  const sh = sheet_();
  if (sh.getLastRow() <= 1) writeHeader_(sh);
}

// ---------- 試算表 ----------
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('尚未設定試算表，請先在 Apps Script 執行「一次設定」');
  return SpreadsheetApp.openById(id);
}

// 取得「名單」分頁；表頭不在這裡覆寫（讀取時要先看舊表頭）
function sheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME, 0);
    writeHeader_(sh);
  }
  return sh;
}

function writeHeader_(sh) {
  const maxCols = sh.getMaxColumns();
  sh.getRange(1, 1, 1, maxCols).clear();
  sh.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function rev_() {
  return Number(PropertiesService.getScriptProperties().getProperty('REV') || 0);
}

function toCount_(v, isCheck) {
  if (isCheck) return (v === true || v === 1 || v === '1' || v === '✓' || v === 'V' || v === 'v') ? 1 : 0;
  const n = Math.floor(Number(v));
  return isFinite(n) && n > 0 ? n : 0;
}

function toMoney_(v) {
  const n = Math.round(Number(v));
  return isFinite(n) && n > 0 ? n : 0;
}

// ---------- 讀取 ----------
// 只讀 A 欄是數字 ID 的列（總計列 A 欄是空的，會自動略過）
// 老師直接在試算表改次數也會被讀回來，網頁會用單價重新算費用（保險會自動重算）。
function load_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn ? sh.getLastColumn() : HEADERS.length, 2);
  const members = [];
  if (last >= 2) {
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const colOf = {};
    header.forEach(function (h, i) {
      const key = READ_NAMES[String(h).trim()];
      if (key && !(key in colOf)) colOf[key] = i;
    });
    const values = sh.getRange(2, 1, last - 1, lastCol).getValues();
    values.forEach(function (r) {
      const id = Number(r[0]);
      const name = String(r[1] == null ? '' : r[1]).trim();
      if (!id || !name) return;
      const counts = {};
      Object.keys(colOf).forEach(function (k) {
        counts[k] = toCount_(r[colOf[k]], CHECK_KEYS.indexOf(k) >= 0);
      });
      members.push({ id: id, name: name, counts: counts });
    });
  }
  return { ok: true, rev: rev_(), version: CODE_VERSION, members: members };
}

// ---------- 儲存（整份名單覆蓋） ----------
// body.rev：網頁手上的版本號；和伺服器不同代表別台裝置剛改過 → 回傳 conflict 與最新名單
// body.force：只在第一次把舊的本機資料搬上雲端時使用
function save_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cur = rev_();
    if (!body.force && (body.rev == null || Number(body.rev) !== cur)) {
      const data = load_();
      data.ok = false;
      data.conflict = true;
      return data;
    }

    const list = Array.isArray(body.members) ? body.members : [];
    const rows = [];
    list.forEach(function (m) {
      const id = Number(m && m.id);
      const name = String(m && m.name || '').trim();
      if (!id || !name) return;
      const counts = (m && m.counts) || {};
      const fees = (m && m.fees) || {};
      const row = [id, name];
      COUNT_COLS.forEach(function (c) {
        const isCheck = CHECK_KEYS.indexOf(c[0]) >= 0;
        const n = toCount_(counts[c[0]], isCheck);
        row.push(isCheck ? (n ? '✓' : '') : n);
      });
      let total = 0;
      FEE_COLS.forEach(function (c) {
        const f = toMoney_(fees[c[0]]);
        total += f;
        row.push(f);
      });
      row.push(total);
      rows.push(row);
    });

    const sh = sheet_();
    const maxRows = sh.getMaxRows();
    if (maxRows >= 2) sh.getRange(2, 1, maxRows - 1, sh.getMaxColumns()).clear();
    writeHeader_(sh);

    const feeStart = 3 + COUNT_COLS.length;   // 第一個費用欄
    const checkCols = CHECK_KEYS.map(function (k) {
      return 3 + COUNT_COLS.map(function (c) { return c[0]; }).indexOf(k);
    });

    if (rows.length) {
      const n = rows.length;
      sh.getRange(2, 1, n, HEADERS.length).setValues(rows);
      sh.getRange(2, feeStart, n, FEE_COLS.length + 1).setNumberFormat('#,##0');
      checkCols.forEach(function (col) { sh.getRange(2, col, n, 1).setHorizontalAlignment('center'); });

      // 空一列，再放總計列（公式自動加總）
      const totalRow = n + 3;
      const totals = ['', '總計（' + n + ' 人）'];
      for (let col = 3; col <= HEADERS.length; col++) {
        const L = colLetter_(col);
        totals.push(checkCols.indexOf(col) >= 0
          ? '=COUNTIF(' + L + '2:' + L + (n + 1) + ',"✓")'
          : '=SUM(' + L + '2:' + L + (n + 1) + ')');
      }
      sh.getRange(totalRow, 1, 1, HEADERS.length)
        .setValues([totals])
        .setFontWeight('bold')
        .setBackground('#fff6c7');
      sh.getRange(totalRow, feeStart, 1, FEE_COLS.length + 1).setNumberFormat('#,##0');
      checkCols.forEach(function (col) { sh.getRange(totalRow, col, 1, 1).setHorizontalAlignment('center'); });
    }

    const next = cur + 1;
    PropertiesService.getScriptProperties().setProperty('REV', String(next));
    return { ok: true, rev: next, version: CODE_VERSION, count: rows.length };
  } finally {
    lock.releaseLock();
  }
}

// ---------- 工具 ----------
function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
