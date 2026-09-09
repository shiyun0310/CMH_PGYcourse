/*  =====================================================================
 *  PGY 訓練課程表 — Google Apps Script 後端 (Web App API)
 *  ---------------------------------------------------------------------
 *  用途：把 Google Sheets 的課程表轉成 JSON 給前端網頁讀取，
 *        並提供單一儲存格／整列的寫回功能。
 *
 *  ── 部署步驟 ─────────────────────────────────────────────────────────
 *  1. 打開你的課程表 Google Sheets → 擴充功能 → Apps Script
 *  2. 把本檔內容整份貼上，存檔
 *  3. 若腳本不是綁在該試算表上，請把下方 SPREADSHEET_ID 填成試算表 ID
 *  4. 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *       執行身分：我
 *       誰可以存取：任何人           ← 給網頁讀取用；只給內部看可選「機構內任何人」
 *  5. 複製「網頁應用程式」網址，貼到 assets/config.js 的 GAS_WEB_APP_URL
 *
 *  ── API ──────────────────────────────────────────────────────────────
 *  GET  ?action=schedule&sheet=工作表名稱      取得課表 JSON
 *  GET  ?action=sheets                        取得所有工作表名稱
 *  GET  ?action=notices[&sheet=公告]           取得公告欄留言 JSON
 *  GET  ?action=ping                          健康檢查
 *       任何 GET 加上 &callback=fn 會回傳 JSONP（避開瀏覽器 CORS）
 *
 *  POST body(JSON, Content-Type: text/plain) :
 *    { action:'updateCell', sheet:'11509(詩)', 簡碼:'2204',
 *      month:'2026-08', value:'外', token:'...' }
 *    { action:'updateRow',  sheet:'11509(詩)', 簡碼:'2204',
 *      months:{ '2026-08':'外', '2026-09':'外' }, token:'...' }
 *  ===================================================================== */

/* ======================= 設定區 ======================= */

/** 留空 = 使用綁定本腳本的試算表；否則填試算表 ID（網址 /d/ 與 /edit 之間那段） */
var SPREADSHEET_ID = '';

/** 表頭固定欄位名稱（順序不拘，以名稱比對；請與試算表第 1 列一致） */
var FIXED_COLS = ['長期導師', '學年度', '組別', '期程', '人事號', '簡碼', '受訓醫師'];

/** 表頭所在列（1 起算） */
var HEADER_ROW = 1;

/** 判斷「這一列是有效資料」用的欄位；此欄為空即視為空白列 */
var KEY_COL = '受訓醫師';

/** 寫入用的識別欄位（前端 POST 以此欄定位列） */
var ID_COL = '簡碼';

/** 寫入保護權杖：留空 = 不檢查。建議設定一組亂數字串，並同步寫在前端。 */
var WRITE_TOKEN = '';

/** 讀取快取秒數（0 = 不快取）。試算表常改動就設小一點。 */
var CACHE_SECONDS = 30;


/* ---- 公告欄（留言板）---------------------------------------------------
 *  在同一份試算表新增一個工作表（預設叫「公告」），第 1 列放表頭：
 *
 *    日期        2026/9/9            留空會排在最後面
 *    分類        排程異動            自由填寫，會變成一個小標籤
 *    標題        9 月急診梯次調整
 *    內容        改成…（可換行）      只填內容不填標題也可以
 *    置頂        是                   是／Y／TRUE 會固定排在最上面
 *    顯示        否                   填「否／N／FALSE」就暫時不顯示，不必刪除整列
 *    對象        管理者               管理者=只有 index.html 看得到；
 *                                     公開=只有 office.html；留空=兩頁都顯示
 *    張貼者      教學部
 *
 *  欄位順序不拘，以名稱比對；不需要的欄整欄不放也可以。
 *  找不到這個工作表時回傳空陣列（不是錯誤），前端就整塊不顯示。
 * --------------------------------------------------------------------- */

/** 公告欄工作表名稱（前端可用 ?sheet= 覆寫） */
var NOTICE_SHEET = '公告';

/** 公告欄表頭別名：第一個名稱是標準欄名，後面都是可接受的寫法 */
var NOTICE_ALIASES = {
  date:     ['日期', '時間', '公告日期', 'date'],
  tag:      ['分類', '類別', '標籤', 'tag', 'type'],
  title:    ['標題', '主旨', 'title'],
  body:     ['內容', '訊息', '留言', '說明', 'body', 'message'],
  pin:      ['置頂', '重要', 'pin', 'top'],
  show:     ['顯示', '啟用', '狀態', 'show'],
  audience: ['對象', '顯示頁面', '範圍', 'audience'],
  author:   ['張貼者', '發布者', '公告者', '作者', 'author']
};

/** 公告欄快取秒數；比課表短，改完留言比較快看得到 */
var NOTICE_CACHE_SECONDS = 15;


/* ======================= 進入點 ======================= */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    switch (p.action || 'schedule') {
      case 'ping':
        out = { ok: true, pong: true, time: nowStr_() };
        break;
      case 'sheets':
        out = { ok: true, sheets: listSheets_() };
        break;
      case 'notices':
        out = getNotices_(p.sheet, p.nocache === '1');
        break;
      case 'schedule':
      default:
        out = getSchedule_(p.sheet, p.nocache === '1');
        break;
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return reply_(out, p.callback);
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return reply_({ ok: false, error: '請求內容不是有效的 JSON' });
  }

  var out;
  try {
    if (WRITE_TOKEN && body.token !== WRITE_TOKEN) throw new Error('權杖錯誤，拒絕寫入');
    switch (body.action) {
      case 'updateCell':
        out = updateCell_(body.sheet, body[ID_COL] || body.id, body.month, body.value);
        break;
      case 'updateRow':
        out = updateRow_(body.sheet, body[ID_COL] || body.id, body.months || {});
        break;
      default:
        throw new Error('未知的 action：' + body.action);
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return reply_(out, body.callback);
}

/** 統一輸出：有 callback 就回 JSONP，否則回 JSON */
function reply_(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback && /^[\w.$]+$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}


/* ======================= 讀取 ======================= */

function book_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function listSheets_() {
  return book_().getSheets()
    .filter(function (s) { return !s.isSheetHidden(); })
    .map(function (s) { return s.getName(); });
}

function sheetByName_(name) {
  var ss = book_();
  var sh = name ? ss.getSheetByName(name) : null;
  if (!sh) sh = ss.getSheets()[0];
  if (!sh) throw new Error('找不到任何工作表');
  return sh;
}

/**
 * 讀出整份課表。
 * 回傳：{ ok, sheet, sheets, fixedCols, monthCols[{key,label,year,month}],
 *         rows[{固定欄..., months:{key:{value,color}}}], updatedAt }
 */
function getSchedule_(sheetName, noCache) {
  var sh = sheetByName_(sheetName);
  var name = sh.getName();
  var cache = CacheService.getScriptCache();
  var ck = 'pgy_v2_' + name;

  if (CACHE_SECONDS > 0 && !noCache) {
    var hit = cache.get(ck);
    if (hit) return JSON.parse(hit);
  }

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < HEADER_ROW + 1 || lastCol < 1) {
    return { ok: true, sheet: name, sheets: listSheets_(), fixedCols: FIXED_COLS, monthCols: [], rows: [], updatedAt: nowStr_() };
  }


  var headerVals = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  /* --- 解析表頭：固定欄 vs 月份欄 --- */
  var fixedIdx = {};     // 欄位名 -> 0-based index
  var monthCols = [];    // {key,label,year,month,idx}
  for (var c = 0; c < lastCol; c++) {
    var raw = headerVals[c];
    var m = parseMonthHeader_(raw);
    if (m) {
      m.idx = c;
      monthCols.push(m);
      continue;
    }
    var title = String(raw == null ? '' : raw).trim();
    if (title && FIXED_COLS.indexOf(title) >= 0 && !(title in fixedIdx)) fixedIdx[title] = c;
  }

  if (!(KEY_COL in fixedIdx)) throw new Error('表頭找不到「' + KEY_COL + '」欄，請確認第 ' + HEADER_ROW + ' 列');

  /* --- 讀取內容與底色 --- */
  var n = lastRow - HEADER_ROW;
  var rng = sh.getRange(HEADER_ROW + 1, 1, n, lastCol);
  var disp = rng.getDisplayValues();
  var bg = rng.getBackgrounds();

  var rows = [];
  for (var r = 0; r < n; r++) {
    var keyVal = String(disp[r][fixedIdx[KEY_COL]] || '').trim();
    if (!keyVal) continue;                       // 跳過空白列

    var rec = { _row: HEADER_ROW + 1 + r };
    FIXED_COLS.forEach(function (f) {
      rec[f] = (f in fixedIdx) ? String(disp[r][fixedIdx[f]] || '').trim() : '';
    });

    var months = {};
    monthCols.forEach(function (mc) {
      var v = String(disp[r][mc.idx] || '').trim();
      if (!v) return;
      months[mc.key] = { value: v, color: normColor_(bg[r][mc.idx]) };
    });
    rec.months = months;
    rows.push(rec);
  }

  var out = {
    ok: true,
    sheet: name,
    sheets: listSheets_(),
    fixedCols: FIXED_COLS,
    monthCols: monthCols.map(function (m) {
      return { key: m.key, label: m.label, year: m.year, month: m.month };
    }),
    rows: rows,
    updatedAt: nowStr_()
  };

  if (CACHE_SECONDS > 0) {
    try { cache.put(ck, JSON.stringify(out), CACHE_SECONDS); } catch (e) { /* 超過 100KB 就不快取 */ }
  }
  return out;
}

/** 表頭是否為月份欄；接受日期值，或 2026/8、2026-08、2026年8月、115/08(民國) 等字樣 */
function parseMonthHeader_(raw) {
  if (raw instanceof Date) {
    return mk_(raw.getFullYear(), raw.getMonth() + 1);
  }
  var s = String(raw == null ? '' : raw).trim();
  if (!s) return null;

  var m = s.match(/^(\d{4})\s*[\/\-年.]\s*(\d{1,2})\s*月?$/);
  if (m) return mk_(Number(m[1]), Number(m[2]));

  m = s.match(/^(\d{2,3})\s*[\/\-年.]\s*(\d{1,2})\s*月?$/);      // 民國年
  if (m) return mk_(Number(m[1]) + 1911, Number(m[2]));

  return null;

  function mk_(y, mo) {
    if (!(mo >= 1 && mo <= 12) || !(y >= 1900 && y <= 2200)) return null;
    var mm = ('0' + mo).slice(-2);
    return { key: y + '-' + mm, label: y + '/' + mm, year: y, month: mo };
  }
}

/** Sheets 白色/透明底視為「沒有設色」，交給前端用預設色票 */
function normColor_(hex) {
  var h = String(hex || '').toLowerCase();
  if (!h || h === '#ffffff' || h === '#fff' || h === 'none') return null;
  return h;
}

function nowStr_() {
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  return Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss');
}


/* ======================= 公告欄（留言板） ======================= */

/**
 * 讀出公告欄。
 * 回傳：{ ok, sheet, missing, notices[{date,ts,tag,title,body,pin,audience,author}], updatedAt }
 *
 * 找不到工作表時回傳 missing:true 與空陣列，而不是丟錯 ——
 * 還沒建「公告」分頁的人，網頁只是不顯示公告欄，其他功能照常。
 */
function getNotices_(sheetName, noCache) {
  var name = String(sheetName || NOTICE_SHEET || '').trim();
  var cache = CacheService.getScriptCache();
  var ck = 'pgy_notice_' + name;

  if (NOTICE_CACHE_SECONDS > 0 && !noCache) {
    var hit = cache.get(ck);
    if (hit) return JSON.parse(hit);
  }

  var sh = name ? book_().getSheetByName(name) : null;
  if (!sh) return { ok: true, sheet: name, missing: true, notices: [], updatedAt: nowStr_() };

  var out = { ok: true, sheet: name, missing: false, notices: [], cols: [], missingCols: [], updatedAt: nowStr_() };
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < HEADER_ROW + 1 || lastCol < 1) return out;

  /* --- 表頭比對：以別名找出各欄位置，找不到就當作沒有這一欄 --- */
  var header = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h == null ? '' : h).trim().toLowerCase(); });
  var idx = {};
  Object.keys(NOTICE_ALIASES).forEach(function (field) {
    NOTICE_ALIASES[field].forEach(function (alias) {
      if (field in idx) return;
      var at = header.indexOf(String(alias).toLowerCase());
      if (at >= 0) idx[field] = at;
    });
  });

  /* 認得的欄與沒找到的欄都回報出去 —— check.html 會列出來。
     公告顯示不出日期，最常見的原因就是根本沒有「日期」這一欄。 */
  Object.keys(NOTICE_ALIASES).forEach(function (field) {
    (field in idx ? out.cols : out.missingCols).push(NOTICE_ALIASES[field][0]);
  });

  var n = lastRow - HEADER_ROW;
  var rng = sh.getRange(HEADER_ROW + 1, 1, n, lastCol);
  var disp = rng.getDisplayValues();
  var vals = rng.getValues();

  function cell(r, field) {
    return (field in idx) ? String(disp[r][idx[field]] || '').trim() : '';
  }

  for (var r = 0; r < n; r++) {
    var title = cell(r, 'title');
    var body = cell(r, 'body');
    if (!title && !body) continue;                       // 標題與內容都空 = 空白列
    if (isNo_(cell(r, 'show'))) continue;                // 顯示欄填「否」= 暫時不公告

    /* 日期：畫面上顯示使用者自己打的樣子，排序另外用時間戳。
       手動輸入的文字日期（2026/9/9、2026-09-09）也試著解析出時間戳。 */
    var rawDate = (idx.date != null) ? vals[r][idx.date] : '';
    var ts = 0;
    if (rawDate instanceof Date) ts = rawDate.getTime();
    else {
      var m = String(rawDate == null ? '' : rawDate).match(/(\d{4})\s*[\/\-.年]\s*(\d{1,2})(?:\s*[\/\-.月]\s*(\d{1,2}))?/);
      if (m) ts = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3] || 1)).getTime();
    }

    out.notices.push({
      date: cell(r, 'date'),
      ts: ts,
      tag: cell(r, 'tag'),
      title: title,
      body: body,
      pin: isYes_(cell(r, 'pin')),
      audience: cell(r, 'audience'),
      author: cell(r, 'author'),
      _row: HEADER_ROW + 1 + r
    });
  }

  if (NOTICE_CACHE_SECONDS > 0) {
    try { cache.put(ck, JSON.stringify(out), NOTICE_CACHE_SECONDS); } catch (e) { /* 太大就不快取 */ }
  }
  return out;
}

/** 「是／Y／TRUE／1／V／✓」都算是；空白算否 */
function isYes_(v) {
  return /^(是|y|yes|true|1|v|✓|o)$/i.test(String(v == null ? '' : v).trim());
}

/** 「否／N／FALSE／0／停用／隱藏」都算否；空白算是（預設顯示） */
function isNo_(v) {
  return /^(否|n|no|false|0|x|停用|隱藏|關)$/i.test(String(v == null ? '' : v).trim());
}


/* ======================= 寫入 ======================= */

/** 依 ID_COL（簡碼）找出該列列號與月份欄對應 */
function locate_(sheetName, id) {
  var sh = sheetByName_(sheetName);
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();

  var header = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  var idIdx = -1, monthIdx = {};
  for (var c = 0; c < lastCol; c++) {
    var mm = parseMonthHeader_(header[c]);
    if (mm) { monthIdx[mm.key] = c + 1; continue; }
    if (String(header[c] || '').trim() === ID_COL) idIdx = c + 1;
  }
  if (idIdx < 0) throw new Error('表頭找不到「' + ID_COL + '」欄');

  var ids = sh.getRange(HEADER_ROW + 1, idIdx, Math.max(0, lastRow - HEADER_ROW), 1).getDisplayValues();
  var row = -1;
  for (var r = 0; r < ids.length; r++) {
    if (String(ids[r][0]).trim() === String(id).trim()) { row = HEADER_ROW + 1 + r; break; }
  }
  if (row < 0) throw new Error('找不到 ' + ID_COL + ' = ' + id + ' 的資料列');

  return { sheet: sh, row: row, monthIdx: monthIdx };
}

function updateCell_(sheetName, id, monthKey, value) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var loc = locate_(sheetName, id);
    var col = loc.monthIdx[monthKey];
    if (!col) throw new Error('找不到月份欄：' + monthKey);
    loc.sheet.getRange(loc.row, col).setValue(value == null ? '' : value);
    CacheService.getScriptCache().remove('pgy_v2_' + loc.sheet.getName());
    return { ok: true, sheet: loc.sheet.getName(), row: loc.row, month: monthKey, value: value, updatedAt: nowStr_() };
  } finally {
    lock.releaseLock();
  }
}

function updateRow_(sheetName, id, months) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var loc = locate_(sheetName, id);
    var done = [];
    Object.keys(months).forEach(function (k) {
      var col = loc.monthIdx[k];
      if (!col) return;
      loc.sheet.getRange(loc.row, col).setValue(months[k] == null ? '' : months[k]);
      done.push(k);
    });
    CacheService.getScriptCache().remove('pgy_v2_' + loc.sheet.getName());
    return { ok: true, sheet: loc.sheet.getName(), row: loc.row, updated: done, updatedAt: nowStr_() };
  } finally {
    lock.releaseLock();
  }
}


/* ======================= 試算表選單（選用） ======================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PGY 課程表')
    .addItem('清除 API 快取', 'clearCache')
    .addItem('檢查表頭格式', 'checkHeader')
    .addToUi();
}

function clearCache() {
  var c = CacheService.getScriptCache();
  listSheets_().forEach(function (n) { c.remove('pgy_v2_' + n); c.remove('pgy_notice_' + n); });
  SpreadsheetApp.getActive().toast('已清除 API 快取', 'PGY 課程表', 4);
}

function checkHeader() {
  var sh = SpreadsheetApp.getActiveSheet();

  var header = sh.getRange(HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
  var fixed = [], months = [], unknown = [];
  header.forEach(function (h) {
    var m = parseMonthHeader_(h);
    if (m) { months.push(m.label); return; }
    var t = String(h == null ? '' : h).trim();
    if (!t) return;
    (FIXED_COLS.indexOf(t) >= 0 ? fixed : unknown).push(t);
  });
  var msg = '工作表：' + sh.getName() +
    '\n\n可辨識的固定欄（' + fixed.length + '）：\n' + (fixed.join('、') || '（無）') +
    '\n\n可辨識的月份欄（' + months.length + '）：\n' + (months.join('、') || '（無）') +
    '\n\n未使用的欄（' + unknown.length + '）：\n' + (unknown.join('、') || '（無）');
  SpreadsheetApp.getUi().alert('表頭檢查結果', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
