// ===== 潛水旅費動態結算器：伺服器程式 =====
// 放在 GitHub，由 Apps Script 的「啟動程式」(apps-script/loader.gs) 自動下載執行。
// 這個檔案是公開的，不能寫試算表 ID；一律從指令碼屬性讀取。
// 快取（CacheService）的鍵 'code' 是啟動程式在用的，這裡不能用。
const CODE_VERSION = 1;

const SHEET_NAME = '名單';

// 次數欄位（順序 = 試算表 C~K 欄）
const COUNT_KEYS = ['shoreDive', 'lodging', 'scooter', 'insurance', 'lunch', 'breakfast', 'ferry', 'boatDive', 'owCourse'];
const CHECK_KEYS = ['ferry', 'boatDive', 'owCourse'];
// 費用欄位（順序 = 試算表 L~T 欄，和網頁對帳表相同）
const FEE_KEYS = ['owCourse', 'shoreDive', 'boatDive', 'insurance', 'lodging', 'ferry', 'scooter', 'lunch', 'breakfast'];

const HEADERS = [
  'ID', '姓名',
  '岸潛支數', '住宿晚數', '機車天數', '保險天數', '午餐天數', '早餐天數', '船票', '船潛', 'OW課程',
  'OW課程費', '岸潛費', '船潛費', '保險費', '住宿費', '船票費', '機車費', '午餐費', '早餐費',
  '個人總應繳'
];

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
  sheet_();
}

// ---------- 試算表 ----------
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('尚未設定試算表，請先在 Apps Script 執行「一次設定」');
  return SpreadsheetApp.openById(id);
}

function sheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
  sh.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
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
// 老師直接在試算表改次數也會被讀回來，網頁會用單價重新算費用。
function load_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  const members = [];
  if (last >= 2) {
    const values = sh.getRange(2, 1, last - 1, 2 + COUNT_KEYS.length).getValues();
    values.forEach(function (r) {
      const id = Number(r[0]);
      const name = String(r[1] == null ? '' : r[1]).trim();
      if (!id || !name) return;
      const counts = {};
      COUNT_KEYS.forEach(function (k, i) {
        counts[k] = toCount_(r[2 + i], CHECK_KEYS.indexOf(k) >= 0);
      });
      members.push({ id: id, name: name, counts: counts });
    });
  }
  return { ok: true, rev: rev_(), members: members };
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
      COUNT_KEYS.forEach(function (k) {
        const isCheck = CHECK_KEYS.indexOf(k) >= 0;
        const c = toCount_(counts[k], isCheck);
        row.push(isCheck ? (c ? '✓' : '') : c);
      });
      let total = 0;
      FEE_KEYS.forEach(function (k) {
        const f = toMoney_(fees[k]);
        total += f;
        row.push(f);
      });
      row.push(total);
      rows.push(row);
    });

    const sh = sheet_();
    const maxRows = sh.getMaxRows();
    if (maxRows >= 2) sh.getRange(2, 1, maxRows - 1, sh.getMaxColumns()).clear();

    if (rows.length) {
      const n = rows.length;
      sh.getRange(2, 1, n, HEADERS.length).setValues(rows);
      sh.getRange(2, 12, n, FEE_KEYS.length + 1).setNumberFormat('#,##0');
      sh.getRange(2, 9, n, 3).setHorizontalAlignment('center');

      // 空一列，再放總計列（公式自動加總）
      const totalRow = n + 3;
      const totals = ['', '總計（' + n + ' 人）'];
      for (let col = 3; col <= HEADERS.length; col++) {
        const L = colLetter_(col);
        totals.push(col >= 9 && col <= 11
          ? '=COUNTIF(' + L + '2:' + L + (n + 1) + ',"✓")'
          : '=SUM(' + L + '2:' + L + (n + 1) + ')');
      }
      sh.getRange(totalRow, 1, 1, HEADERS.length)
        .setValues([totals])
        .setFontWeight('bold')
        .setBackground('#fff6c7');
      sh.getRange(totalRow, 12, 1, FEE_KEYS.length + 1).setNumberFormat('#,##0');
      sh.getRange(totalRow, 9, 1, 3).setHorizontalAlignment('center');
    }

    const next = cur + 1;
    PropertiesService.getScriptProperties().setProperty('REV', String(next));
    return { ok: true, rev: next, count: rows.length };
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
