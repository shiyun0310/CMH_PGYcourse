/* =====================================================================
 *  PGY 訓練課程表 — 前端主程式
 *  資料來源：Google Sheets（透過 GAS Web App）；連線失敗時退回離線範例。
 * ===================================================================== */
(function () {
  'use strict';

  var CFG = window.PGY_CONFIG || {};
  var FIXED_ORDER = ['長期導師', '學年度', '期程', '組別', '人事號', '簡碼', '受訓醫師'];

  var state = {
    data: null,          // 正規化後的資料
    sheets: [],          // 試算表內所有工作表名稱
    source: 'sample',    // 'sheets' | 'sample'
    error: '',
    view: 'grid',
    monthIndex: 0,
    filters: { q: '', 學年度: '', 期程: '', 組別: '', 長期導師: '', cat: '' },   // cat 為單選，'' = 全部
    personQuery: '',     // 個人時程要先輸入姓名或人事號才顯示
    monthUnits: {}       // 月份檢視的科別複選（由 <body data-unit-filter> 啟用）
  };

  /* ---------------------------------------------------------------- 工具 */
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  /* office.html 是純資料版，沒有頁首、篩選列與頁尾。
   * 以下兩個小工具讓同一支 app.js 能同時服務兩種外殼：
   * 元素不存在就安靜跳過，而不是丟 TypeError 讓整頁掛掉。 */
  function on(sel, evt, fn) {
    var el = $(sel);
    if (el) el.addEventListener(evt, fn);
  }
  function withEl(sel, fn) {
    var el = $(sel);
    if (el) fn(el);
    return el;
  }

  /* 外殼決定有哪些檢視：office.html 沒有個人時程頁籤，
   * 總覽表的姓名就不能做成可點的連結，否則會跳到沒有對應分頁的畫面。 */
  function hasView(v) {
    return !!$('.tab[data-view="' + v + '"]');
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  var INK_DARK = '#1b2432', INK_LIGHT = '#ffffff';

  /* 相對亮度（WCAG） */
  function luminance(hex) {
    var h = String(hex || '').replace('#', '');
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    if (h.length !== 6) return null;
    var f = function (v) { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
    return 0.2126 * f(parseInt(h.slice(0, 2), 16)) +
           0.7152 * f(parseInt(h.slice(2, 4), 16)) +
           0.0722 * f(parseInt(h.slice(4, 6), 16));
  }

  function contrast(l1, l2) {
    return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05);
  }

  /* 色塊描邊顏色：白色色塊要在白底上看得見，黑色色塊要在深色模式下看得見。
   * 只有極深的顏色才改用淺色描邊，其餘維持原本的深色描邊。 */
  function ringOn(hex) {
    var L = luminance(hex);
    return (L != null && L < 0.06) ? 'rgba(255,255,255,.34)' : 'rgba(0,0,0,.18)';
  }

  /* 依背景色挑深色或白色文字 —— 實際比較兩者的對比度，取高的那個。
   * （固定門檻會在中間色調挑錯，例如 #00B0F0 配白字只有 2.5:1） */
  function inkOn(hex) {
    var L = luminance(hex);
    if (L == null) return INK_DARK;
    return contrast(L, luminance(INK_DARK)) >= contrast(L, luminance(INK_LIGHT)) ? INK_DARK : INK_LIGHT;
  }

  /* 值 → 科別分類：只看開頭，都不符合就歸 DEFAULT_CATEGORY */
  function categoryOf(v) {
    var s = String(v || '').trim();
    if (!s) return '';                       // 空白格＝沒有排課，不算任何科別
    var rules = CFG.CATEGORY_RULES || [];
    for (var i = 0; i < rules.length; i++) {
      if (s.indexOf(rules[i].prefix) === 0) return rules[i].name;
    }
    return CFG.DEFAULT_CATEGORY || '其他';
  }

  /* 儲存格顏色
   *   COLOR_SOURCE = 'category'（預設）：一律依科別上色，同科別顏色必定一致
   *   COLOR_SOURCE = 'sheets'         ：Sheets 底色優先，沒有底色才用色票 */
  /* 科別名 → 顏色（例：'內科'） */
  function catColor(cat) {
    var map = CFG.CATEGORY_COLORS || {};
    return map[cat] || map[CFG.DEFAULT_CATEGORY] || '#d9d9d9';
  }

  /* 原始值 → 顏色（例：'內(Y2不分)' → 內科的黃色） */
  function valueColor(v) {
    return catColor(categoryOf(v));
  }

  /* 取出「訓練單位」本體
   *   內(Y2不分)、內(Y2內)、內  → 內      （括號一律視為註記）
   *   婦(完訓)、婦(Y2婦)、婦     → 婦
   *   急、急-內、急-外           → 急      （急診不在 KEEP_SUFFIX_CATEGORIES）
   *   選-眼科                   → 選-眼科  （選修在 KEEP_SUFFIX_CATEGORIES，後綴是實際單位）
   *   社-新樓                   → 社-新樓 */
  function baseUnit(v) {
    var s = String(v == null ? '' : v).trim();
    if (!s) return '';

    // 1) 去掉括號註記（半形與全形）
    var cut = s.search(/[（(]/);
    var base = cut < 0 ? s : (s.slice(0, cut).trim() || s);

    // 2) 除非該科別的後綴代表實際單位，否則連字號後面也視為註記
    if ((CFG.KEEP_SUFFIX_CATEGORIES || []).indexOf(categoryOf(s)) >= 0) return base;
    var dash = base.search(/[-–—－]/);
    return dash > 0 ? base.slice(0, dash).trim() : base;
  }

  /* 月份檢視 / 統計的分組鍵
   *   MONTH_GROUP_BY = 'unit'（預設）依訓練單位，選-眼科 與 選-耳鼻喉 分開
   *   MONTH_GROUP_BY = 'category'    依科別，所有 選- 併成一組「選修」 */
  function groupKeyOf(cell) {
    return CFG.MONTH_GROUP_BY === 'category' ? cell.cat : baseUnit(cell.value);
  }

  function groupColor(key) {
    return CFG.MONTH_GROUP_BY === 'category' ? catColor(key) : valueColor(key);
  }

  /* 月份卡片裡名字旁的小標籤。卡片標題已經寫明科別，標籤只用來補「多出來的資訊」：
   *   在「急」卡片   急           → 不標          與卡片同名，標了也是廢話
   *   在「外」卡片   外(Y2不分)   → (Y2不分)      去掉重複的「外」，只留註記
   *   在「急」卡片   急(完訓)     → (完訓)
   *   在「急」卡片   急-外        → 急-外         連字號後面是實際去的次專科，
   *   在「急」卡片   急-外(完訓)  → 急-外(完訓)   要看得出是 急-外 還是 急-內，整串保留
   * 判斷方式：與組名完全相同就不標；去掉組名後若緊接著括號，那只是註記，可以省略組名；
   * 若緊接著連字號（或其他字），代表是不同的訓練單位，保留原文。 */
  /* 外殼是否要在月份檢視提供科別篩選。
   * 由 <body data-unit-filter> 指定，目前三個外殼都有開。
   * 上方篩選列的科別是「科別」單選，這排按鈕是「訓練單位」複選，
   * 粒度不同（內科 vs 選-眼科、社-郭綜），兩者可以疊加使用。 */
  function wantsUnitFilter() {
    return document.body.hasAttribute('data-unit-filter');
  }

  /* 篩選按鈕的單位清單 = 資料裡出現過的所有訓練單位 ∪ config 的固定清單。
   * 併入固定清單是為了讓當月沒人的科別也選得到，才能明確顯示「查無資料」。 */
  function allUnits() {
    var seen = {}, out = [];
    (state.data ? state.data.rows : []).forEach(function (r) {
      Object.keys(r.months).forEach(function (k) {
        var u = filterKeyOf(groupKeyOf(r.months[k]));
        if (u && !seen[u]) { seen[u] = 1; out.push(u); }
      });
    });
    (CFG.MONTH_UNIT_CHOICES || []).forEach(function (u) {
      var f = filterKeyOf(u);
      if (f && !seen[f]) { seen[f] = 1; out.push(f); }
    });

    // 排序：先依科別（CATEGORY_COLORS 的順序），
    // 同科別內沿用 MONTH_UNIT_CHOICES 給的順序，沒列到的排後面再依筆劃
    var catOrder = Object.keys(CFG.CATEGORY_COLORS || {});
    var choice = {};
    (CFG.MONTH_UNIT_CHOICES || []).forEach(function (u, i) { choice[u] = i; });
    var rank = function (v, map, miss) { var i = map.indexOf ? map.indexOf(v) : map[v]; return (i == null || i < 0) ? miss : i; };

    return out.sort(function (a, b) {
      var d = rank(categoryOf(a), catOrder, 99) - rank(categoryOf(b), catOrder, 99);
      if (d) return d;
      d = rank(a, choice, 9999) - rank(b, choice, 9999);
      if (d) return d;
      return a.localeCompare(b, 'zh-Hant');
    });
  }

  /* 篩選按鈕用的鍵：把社區的次分組前綴收斂成同一家醫院。
   *   社-郭綜／社內-郭綜／社外-郭綜／社婦-郭綜 → 篩選鍵都是「社-郭綜」
   * 合併只發生在按鈕與比對上，卡片仍依實際單位分開顯示。 */
  function filterKeyOf(unit) {
    var u = String(unit == null ? '' : unit).trim();
    var merge = (CFG.FILTER_MERGE_PREFIX || {})[categoryOf(u)];
    if (!merge) return u;
    var dash = u.search(/[-–—－]/);
    return dash > 0 ? merge + u.slice(dash) : u;
  }

  function selectedUnits() {
    return Object.keys(state.monthUnits).filter(function (k) { return state.monthUnits[k]; });
  }

  function noteFor(value, key) {
    var v = String(value == null ? '' : value);
    if (v === key) return '';                         // 與組名完全相同，不用標
    if (v.indexOf(key) === 0) {
      var rest = v.slice(key.length);
      if (/^[（(]/.test(rest)) return rest;           // 外(Y2不分) → (Y2不分)
    }
    return v;                                          // 急-外、急-外(完訓) 保留全名
  }

  function colorOf(cell) {
    if (CFG.COLOR_SOURCE === 'sheets' &&
        cell && cell.color && /^#[0-9a-f]{6}$/i.test(cell.color)) return cell.color;
    return valueColor(cell && cell.value);
  }

  /* 期程 → 排序名次：PGY1 → 1、PGY2 → 2、Y1 → 1；取不到數字的排最後 */
  function stageRank(v) {
    var m = String(v == null ? '' : v).match(/\d+/);
    return m ? Number(m[0]) : 9999;
  }

  function uniq(arr) {
    var seen = {}, out = [];
    arr.forEach(function (v) { if (v !== '' && v != null && !seen[v]) { seen[v] = 1; out.push(v); } });
    return out;
  }

  /* ------------------------------------------------------- 讀取 Sheets */
  function jsonp(url, timeoutMs) {
    return new Promise(function (resolve, reject) {
      var cb = 'pgycb_' + Math.random().toString(36).slice(2);
      var sc = document.createElement('script');
      var timer = setTimeout(function () { cleanup(); reject(new Error('JSONP 逾時')); }, timeoutMs || 20000);
      function cleanup() {
        clearTimeout(timer);
        delete window[cb];
        if (sc.parentNode) sc.parentNode.removeChild(sc);
      }
      window[cb] = function (data) { cleanup(); resolve(data); };
      sc.onerror = function () { cleanup(); reject(new Error('JSONP 載入失敗')); };
      sc.src = url + (url.indexOf('?') < 0 ? '?' : '&') + 'callback=' + cb;
      document.body.appendChild(sc);
    });
  }

  function buildUrl(base, params) {
    var qs = Object.keys(params)
      .filter(function (k) { return params[k] !== '' && params[k] != null; })
      .map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); })
      .join('&');
    return base + (base.indexOf('?') < 0 ? '?' : '&') + qs;
  }

  /* 課表與公告欄都走這裡；只有 action 與參數不同 */
  function callGas(params) {
    var base = (CFG.GAS_WEB_APP_URL || '').trim();
    if (!base) return Promise.reject(new Error('尚未設定 GAS_WEB_APP_URL'));
    var url = buildUrl(base, params);
    var mode = CFG.TRANSPORT || 'auto';

    if (mode === 'jsonp') return jsonp(url);

    var direct = fetch(url, { method: 'GET', redirect: 'follow' })
      .then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.json();
      });

    if (mode === 'fetch') return direct;
    return direct.catch(function () { return jsonp(url); });   // auto：CORS 失敗改走 JSONP
  }

  function fetchFromGas(sheetName) {
    return callGas({ action: 'schedule', sheet: sheetName || CFG.DEFAULT_SHEET || '', t: Date.now() });
  }

  /* --------------------------------------------------------- 正規化資料 */
  function normalize(raw) {
    if (!raw || raw.ok === false) throw new Error((raw && raw.error) || '資料格式不正確');

    var monthCols = (raw.monthCols || []).map(function (m, i) {
      return {
        key: m.key || ('c' + i),
        label: m.label || m.key,
        year: Number(m.year) || 0,
        month: Number(m.month) || 0
      };
    });

    var fixedCols = raw.fixedCols && raw.fixedCols.length
      ? raw.fixedCols.slice()
      : ['長期導師', '學年度', '組別', '期程', '人事號', '簡碼', '受訓醫師'];

    var rows = (raw.rows || []).map(function (r, idx) {
      var months = {};
      Object.keys(r.months || {}).forEach(function (k) {
        var c = r.months[k];
        var val = (c && typeof c === 'object') ? c.value : c;
        val = String(val == null ? '' : val).trim();
        if (!val) return;
        months[k] = {
          value: val,
          color: (c && typeof c === 'object' && c.color) ? c.color : null,
          cat: categoryOf(val)
        };
      });
      var rec = { _i: idx, months: months };
      fixedCols.forEach(function (f) { rec[f] = String(r[f] == null ? '' : r[f]).trim(); });
      return rec;
    });

    return {
      sheet: raw.sheet || '',
      fixedCols: fixedCols,
      monthCols: monthCols,
      rows: rows,
      updatedAt: raw.updatedAt || ''
    };
  }

  /* ------------------------------------------------------------ 篩選 */
  function catsInUse() {
    var order = Object.keys(CFG.CATEGORY_COLORS || {});
    var found = {};
    state.data.rows.forEach(function (r) {
      Object.keys(r.months).forEach(function (k) { found[r.months[k].cat] = (found[r.months[k].cat] || 0) + 1; });
    });
    var keys = Object.keys(found).sort(function (a, b) {
      var ia = order.indexOf(a), ib = order.indexOf(b);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
    });
    return keys.map(function (k) { return { name: k, n: found[k] }; });
  }

  function anyCatSelected() {
    return !!state.filters.cat;
  }

  function rowMatches(r) {
    var f = state.filters;
    if (f.學年度 && r['學年度'] !== f.學年度) return false;
    if (f.期程 && r['期程'] !== f.期程) return false;
    if (f.組別 && r['組別'] !== f.組別) return false;
    if (f.長期導師 && r['長期導師'] !== f.長期導師) return false;

    if (f.q) {
      var q = f.q.toLowerCase();
      var hay = state.data.fixedCols.map(function (c) { return r[c]; }).join(' ');
      hay += ' ' + Object.keys(r.months).map(function (k) { return r.months[k].value; }).join(' ');
      if (hay.toLowerCase().indexOf(q) < 0) return false;
    }

    if (anyCatSelected()) {
      var hit = Object.keys(r.months).some(function (k) { return r.months[k].cat === f.cat; });
      if (!hit) return false;
    }
    return true;
  }

  function filteredRows() {
    return state.data.rows.filter(rowMatches);
  }

  /* 儲存格是否因科別篩選而變淡 */
  function cellDim(cell) {
    return anyCatSelected() && cell.cat !== state.filters.cat;
  }

  /* ============================== 畫面 ============================== */

  function renderStatus() {
    var el = $('#status');
    if (!el) return;
    var cls = 'pill', txt = '';
    if (state.source === 'loading') { cls += ' is-loading'; txt = '讀取中…'; }
    else if (state.source === 'sheets') { txt = 'Google Sheets 已連線'; }
    else if (state.source === 'error') { cls += ' is-error'; txt = '連線失敗（顯示範例資料）'; }
    else { cls += ' is-sample'; txt = '離線範例資料'; }
    el.className = cls;
    el.innerHTML = '<span class="dot"></span>' + esc(txt);
  }

  function renderBanner() {
    var box = $('#banner');
    if (!box) return;
    if (state.error) {
      box.innerHTML = '<div class="banner err"><span>⚠</span><div>' + esc(state.error) +
        '　目前顯示的是 <b>assets/sample-data.js</b> 的離線範例資料。</div></div>';
    } else if (state.source === 'sample') {
      box.innerHTML = '<div class="banner warn"><span>ℹ</span><div>目前為<b>離線範例模式</b>：' +
        '請部署 <code>gas/Code.gs</code>，並把網頁應用程式網址填入 <code>assets/config.js</code> 的 ' +
        '<code>GAS_WEB_APP_URL</code>，即可改以 Google Sheets 內容為主。</div></div>';
    } else {
      box.innerHTML = '';
    }
  }

  /* 篩選列的每個元素都是可選的：外殼放了才畫，沒放就安靜跳過。
   * office.html 整列都沒有；115PGY.html 只留四個下拉與清除條件，
   * 沒有搜尋框、工作表下拉與科別色塊。 */
  function renderFilters() {
    var d = state.data, f = state.filters;

    function fillSel(id, values, cur) {
      var sel = $(id);
      if (!sel) return;
      var opts = ['<option value="">全部</option>'].concat(values.map(function (v) {
        return '<option value="' + esc(v) + '"' + (v === cur ? ' selected' : '') + '>' + esc(v) + '</option>';
      }));
      sel.innerHTML = opts.join('');
    }

    fillSel('#f-year', uniq(d.rows.map(function (r) { return r['學年度']; })).sort(), f['學年度']);
    fillSel('#f-stage', uniq(d.rows.map(function (r) { return r['期程']; })).sort(), f['期程']);
    fillSel('#f-group', uniq(d.rows.map(function (r) { return r['組別']; })).sort(), f['組別']);
    fillSel('#f-mentor', uniq(d.rows.map(function (r) { return r['長期導師']; })).sort(), f['長期導師']);

    var sheetSel = $('#f-sheet');
    if (!sheetSel) { /* 純資料版沒有工作表下拉 */ }
    else if (state.sheets.length) {
      sheetSel.parentNode.style.display = '';
      sheetSel.innerHTML = state.sheets.map(function (s) {
        return '<option value="' + esc(s) + '"' + (s === d.sheet ? ' selected' : '') + '>' + esc(s) + '</option>';
      }).join('');
    } else {
      sheetSel.parentNode.style.display = 'none';
    }

    withEl('#chips', function (box) {
      box.innerHTML = catsInUse().map(function (c) {
        var col = catColor(c.name);
        // title 保留人月數，滑鼠停留才顯示，chip 上不放數字
        return '<button class="chip' + (f.cat === c.name ? ' on' : '') + '" data-cat="' + esc(c.name) + '"' +
          ' title="' + esc(c.name + '　' + c.n + ' 人月') + '">' +
          '<i class="sw" style="background:' + col + ';border-color:' + ringOn(col) + '"></i>' +
          esc(c.name) + '</button>';
      }).join('') || '<span style="color:var(--ink-3);font-size:12.5px">尚無課程資料</span>';
    });
  }

  /* ------------------------------------------------------- 檢視：總覽表 */
  function viewGrid() {
    var d = state.data, months = d.monthCols, rows = filteredRows();
    if (!rows.length) return emptyState('沒有符合條件的受訓醫師');

    // 年份分組表頭
    var groups = [], last = null;
    months.forEach(function (m) {
      if (!last || last.year !== m.year) { last = { year: m.year, n: 0 }; groups.push(last); }
      last.n++;
    });

    var fixed = FIXED_ORDER.filter(function (c) { return d.fixedCols.indexOf(c) >= 0; });

    var h = '<div class="tscroll"><table class="grid"><thead>';
    h += '<tr class="yrow">';
    fixed.forEach(function (c, i) {
      h += '<th class="fx' + (i === fixed.length - 1 ? ' fx-last' : '') +
        (c === '受訓醫師' ? ' fx-id' : '') + '"' +
        ' data-fx="' + i + '" rowspan="2">' + esc(c) + '</th>';
    });
    groups.forEach(function (g) {
      h += '<th colspan="' + g.n + '">' + esc(g.year) + ' 年</th>';
    });
    h += '</tr><tr class="mrow">';
    months.forEach(function (m) {
      h += '<th>' + esc(String(m.month).padStart(2, '0')) + '月</th>';
    });
    h += '</tr></thead><tbody>';

    rows.forEach(function (r) {
      h += '<tr>';
      fixed.forEach(function (c, i) {
        var cls = 'fx' + (i === fixed.length - 1 ? ' fx-last' : '') +
          (c === '受訓醫師' ? ' fx-id name' : ' dim');
        var inner = (c === '受訓醫師' && hasView('person'))
          ? '<button class="rowbtn" data-person="' + esc(r[c]) + '">' + esc(r[c] || '—') + '</button>'
          : esc(r[c] || '—');
        h += '<td class="' + cls + '" data-fx="' + i + '">' + inner + '</td>';
      });
      months.forEach(function (m) {
        var cell = r.months[m.key];
        if (!cell) { h += '<td class="empty"></td>'; return; }
        var bg = colorOf(cell), fg = inkOn(bg), dim = cellDim(cell);
        h += '<td' + (dim ? ' style="opacity:.22"' : '') + '><span class="cell" style="background:' + bg +
          ';color:' + fg + ';box-shadow:inset 0 0 0 1px ' + ringOn(bg) + '" title="' +
          esc(m.label + '　' + cell.value) + '">' + esc(cell.value) + '</span></td>';
      });
      h += '</tr>';
    });

    h += '</tbody></table></div>';
    return h;
  }

  /* ------------------------------------------------------- 檢視：月份 */
  function viewMonth() {
    var months = state.data.monthCols, rows = filteredRows();
    if (!months.length) return emptyState('沒有可顯示的月份');
    if (state.monthIndex >= months.length) state.monthIndex = 0;
    var m = months[state.monthIndex];

    var buckets = {};
    rows.forEach(function (r) {
      var c = r.months[m.key];
      if (!c) return;
      if (anyCatSelected() && c.cat !== state.filters.cat) return;
      var k = groupKeyOf(c);
      (buckets[k] = buckets[k] || []).push({ row: r, value: c.value });
    });

    var picked = selectedUnits();
    var keys = Object.keys(buckets).filter(function (k) {
      return !picked.length || state.monthUnits[filterKeyOf(k)];
    }).sort(function (a, b) {
      return buckets[b].length - buckets[a].length || a.localeCompare(b, 'zh-Hant');
    });

    var nav = '<div style="display:flex;gap:8px;align-items:center">' +
      '<button class="btn" id="m-prev"' + (state.monthIndex === 0 ? ' disabled' : '') + '>‹ 上個月</button>' +
      '<select id="m-pick" style="padding:7px 10px;border-radius:9px;border:1px solid var(--line);background:var(--panel-2)">' +
      months.map(function (x, i) {
        return '<option value="' + i + '"' + (i === state.monthIndex ? ' selected' : '') + '>' + esc(x.label) + '</option>';
      }).join('') +
      '</select>' +
      '<button class="btn" id="m-next"' + (state.monthIndex === months.length - 1 ? ' disabled' : '') + '>下個月 ›</button>' +
      '</div>';

    var h = '<div class="view-head"><div><h2>' + esc(m.label) + ' 輪訓分布</h2></div>' +
      nav + '</div>';

    if (wantsUnitFilter()) h += unitFilterBar(picked);

    if (!keys.length) {
      if (picked.length) {
        return h + '<div class="empty-state"><div class="big">🔎</div>' +
          '<div><b>查無資料</b></div>' +
          '<div style="margin-top:6px;font-size:12.5px">' +
          esc(m.label) + ' 沒有受訓醫師被排到：' + esc(picked.join('、')) + '</div></div>';
      }
      return h + emptyState('本月沒有符合條件的排課');
    }

    h += '<div class="mgrid">';
    keys.forEach(function (k) {
      var col = groupColor(k), fg = inkOn(col);
      h += '<div class="mcard"><h3 style="background:' + col + ';color:' + fg +
        ';box-shadow:inset 0 0 0 1px ' + ringOn(col) + '">' +
        '<span>' + esc(k) + '</span><span class="count-badge">' + buckets[k].length + ' 人</span></h3><ul>';
      buckets[k].sort(function (a, b) {
        // 先依期程 PGY1 → PGY2 → …，同期程再依人事號、姓名
        var d = stageRank(a.row['期程']) - stageRank(b.row['期程']);
        if (d) return d;
        // 人事號是固定長度代碼（B508F1、B50810），用字典序；不可加 numeric，
        // 否則會被拆成 B + 數字比較，B5089 會排到 B50810 前面
        d = String(a.row['人事號']).localeCompare(String(b.row['人事號']), 'zh-Hant');
        if (d) return d;
        return String(a.row['受訓醫師']).localeCompare(String(b.row['受訓醫師']), 'zh-Hant');
      }).forEach(function (it) {
        var tag = noteFor(it.value, k);
        var note = tag ? '<i class="vtag">' + esc(tag) + '</i>' : '';
        h += '<li><i class="pid">' + esc(it.row['人事號'] || '—') + '</i>' +
          '<b>' + esc(it.row['受訓醫師']) + '</b>' + note +
          '<span>' + esc([it.row['期程'], it.row['簡碼']].filter(Boolean).join('・')) + '</span></li>';
      });
      h += '</ul></div>';
    });
    h += '</div>';
    return h;
  }

  /* 月份檢視的科別複選按鈕列 */
  function unitFilterBar(picked) {
    var units = allUnits();
    if (!units.length) return '';

    var chips = units.map(function (u) {
      var col = valueColor(u);
      return '<button class="chip' + (state.monthUnits[u] ? ' on' : '') + '" data-unit="' + esc(u) + '">' +
        '<i class="sw" style="background:' + col + ';border-color:' + ringOn(col) + '"></i>' + esc(u) + '</button>';
    }).join('');

    var sel = picked.length
      ? '<div class="ufilter-sel"><span class="lbl">已選擇</span>' +
        picked.map(function (u) { return '<b>' + esc(u) + '</b>'; }).join('') +
        '<button class="btn" id="u-clear">清除篩選</button></div>'
      : '<div class="ufilter-sel none">未選擇任何科別，顯示本月全部單位。可複選。</div>';

    return '<div class="ufilter">' +
      '<div class="ufilter-title">科別篩選</div>' +
      '<div class="chips">' + chips + '</div>' + sel + '</div>';
  }

  /* ------------------------------------------------------- 檢視：個人 */
  function viewPerson() {
    var q = String(state.personQuery || '').trim();

    /* 查詢框與查詢結果分成兩塊：打字時只重繪 #p-results，
     * 輸入框本身完全不動 —— 中文輸入法在組字時若把輸入框換掉，
     * 組字會被打斷，只剩沒組完的注音符號留在欄位裡。 */
    var box = '<div class="view-head"><div><h2>個人輪訓時程</h2>' +
      '<div class="sub">輸入姓名或人事號查詢單一受訓醫師</div></div>' +
      '<div class="pfind">' +
      '<input id="p-q" type="search" placeholder="姓名或人事號…" autocomplete="off" value="' + esc(q) + '">' +
      '<button class="btn" id="p-clear"' + (q ? '' : ' hidden') + '>清除</button>' +
      '</div></div>';

    return box + '<div id="p-results">' + personResults() + '</div>';
  }

  /* 個人時程的查詢結果（不含查詢框） */
  function personResults() {
    var months = state.data.monthCols;
    var q = String(state.personQuery || '').trim();

    if (!q) {
      return '<div class="empty-state"><div class="big">🔎</div>' +
        '<div>請先輸入<b>姓名</b>或<b>人事號</b></div>' +
        '<div style="margin-top:6px;font-size:12.5px">也可以在「總覽表」直接點受訓醫師的姓名</div></div>';
    }

    var needle = q.toLowerCase();
    var rows = filteredRows().filter(function (r) {
      return String(r['受訓醫師']).toLowerCase().indexOf(needle) >= 0 ||
             String(r['人事號']).toLowerCase().indexOf(needle) >= 0;
    });

    if (!rows.length) {
      return emptyState('找不到「' + q + '」，請確認姓名或人事號', '🔎');
    }

    var h = '<div class="plist">';
    rows.forEach(function (r) {
      var tally = {};
      months.forEach(function (m) {
        var c = r.months[m.key];
        if (c) tally[c.cat] = (tally[c.cat] || 0) + 1;
      });
      var filled = Object.keys(tally).reduce(function (s, k) { return s + tally[k]; }, 0);

      h += '<div class="pcard"><header>' +
        '<span class="pn">' + esc(r['受訓醫師'] || '—') + '</span>' +
        '<span class="kv">期程 <b>' + esc(r['期程'] || '—') + '</b></span>' +
        '<span class="kv">組別 <b>' + esc(r['組別'] || '—') + '</b></span>' +
        '<span class="kv">長期導師 <b>' + esc(r['長期導師'] || '—') + '</b></span>' +
        '<span class="kv">人事號 <b>' + esc(r['人事號'] || '—') + '</b></span>' +
        '<span class="kv">簡碼 <b>' + esc(r['簡碼'] || '—') + '</b></span>' +
        '<span class="kv spacer">已排 <b>' + filled + '</b> 個月</span>' +
        '</header><div class="tl">';

      months.forEach(function (m) {
        var c = r.months[m.key];
        if (!c) {
          h += '<div class="seg blank"><span class="m">' + esc(m.label) + '</span><span class="v">—</span></div>';
          return;
        }
        var bg = colorOf(c), fg = inkOn(bg), dim = cellDim(c);
        h += '<div class="seg" style="background:' + bg + ';color:' + fg +
          ';border-color:' + ringOn(bg) + (dim ? ';opacity:.22' : '') + '">' +
          '<span class="m">' + esc(m.label) + '</span><span class="v">' + esc(c.value) + '</span></div>';
      });

      h += '</div><div class="psum">';
      Object.keys(tally).sort(function (a, b) { return tally[b] - tally[a]; }).forEach(function (k) {
        var col = catColor(k);
        h += '<span class="s" style="background:' + col + ';color:' + inkOn(col) +
          ';border-color:' + ringOn(col) + '">' + esc(k) + ' ' + tally[k] + ' 月</span>';
      });
      h += '</div></div>';
    });

    return h + '</div>';
  }

  function emptyState(msg, icon) {
    return '<div class="empty-state"><div class="big">' + (icon || '🗓') + '</div>' +
      '<div>' + esc(msg) + '</div></div>';
  }

  /* --------------------------------------------------------- 主要繪製 */
  function render() {
    renderStatus();
    renderBanner();
    if (!state.data) { $('#view').innerHTML = emptyState('資料載入中…'); return; }

    renderFilters();
    if (!hasView(state.view)) state.view = 'grid';   // 外殼沒有這個頁籤就退回總覽表
    $$('.tab').forEach(function (t) { t.classList.toggle('on', t.dataset.view === state.view); });

    var body;
    if (state.view === 'grid') {
      var months = state.data.monthCols, rows = filteredRows();
      body = '<div class="view-head"><div><h2>' + esc(state.data.sheet || '輪訓總覽') + '</h2>' +
        '<div class="sub">' + rows.length + ' 位受訓醫師 ／ ' + months.length + ' 個月' +
        (state.data.updatedAt ? '　·　更新於 ' + esc(state.data.updatedAt) : '') + '</div></div>' +
        '<div style="display:flex;gap:8px"><button class="btn" id="btn-xlsx" title="下載的 Excel 會保留畫面上的色塊">匯出 Excel</button>' +
        '<button class="btn" id="btn-csv" title="純文字表格，不含顏色">匯出 CSV</button>' +
        '<button class="btn" id="btn-print">列印 / PDF</button></div></div>' + viewGrid();
    } else if (state.view === 'month') {
      body = viewMonth();
    } else {
      body = viewPerson();
    }
    $('#view').innerHTML = body;
    syncStickyOffset();
  }

  /* 表頭與固定欄的黏著位置。
   * 上方：固定欄的 th 有 rowspan="2"，量它會得到兩列的高度，
   *       必須量沒有 rowspan 的「年份」th 才是年份列的真實高度。
   * 左側：七個固定欄要一個接一個排好，所以逐欄累加寬度算出各自的 left。
   *       視窗太窄時七欄會吃掉整個畫面，改成只固定「受訓醫師」一欄 ——
   *       否則橫向捲動後完全看不出這是誰的列。 */
  var FX_MIN_WIDTH = 900;

  function syncStickyOffset() {
    var mrow = $$('.grid thead tr.mrow th');
    if (!mrow.length) return;

    var y = $('.grid thead tr.yrow th:not([rowspan])');
    var h = y ? y.getBoundingClientRect().height : 0;
    mrow.forEach(function (th) { th.style.top = h + 'px'; });

    var wide = window.innerWidth >= FX_MIN_WIDTH;
    var offsets = [], acc = 0;
    $$('.grid thead th.fx').forEach(function (th) {
      offsets.push(acc);
      acc += th.getBoundingClientRect().width;
    });
    $$('.grid .fx').forEach(function (cell) {
      var i = Number(cell.dataset.fx);
      if (wide && offsets[i] != null) cell.style.left = offsets[i] + 'px';
      else cell.style.left = cell.classList.contains('fx-id') ? '0px' : 'auto';
    });
  }

  /* ------------------------------------------------------------ CSV */
  function exportCsv() {
    var d = state.data, months = d.monthCols, rows = filteredRows();
    var head = d.fixedCols.concat(months.map(function (m) { return m.label; }));
    var lines = [head];
    rows.forEach(function (r) {
      lines.push(d.fixedCols.map(function (c) { return r[c] || ''; })
        .concat(months.map(function (m) { return r.months[m.key] ? r.months[m.key].value : ''; })));
    });
    var csv = lines.map(function (row) {
      return row.map(function (v) {
        v = String(v == null ? '' : v);
        return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
      }).join(',');
    }).join('\r\n');

    saveBlob(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' }), fileBase() + '.csv');
  }

  /* 匯出檔名：PGY課程表_工作表名稱（去掉檔名不能用的字元） */
  function fileBase() {
    return 'PGY課程表_' + ((state.data && state.data.sheet) || 'export').replace(/[\\/:*?"<>|]/g, '');
  }

  function saveBlob(blob, filename) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }

  /* ------------------------------------------- 匯出 Excel（保留色塊）
   * CSV 是純文字，存不了顏色 —— 要讓下載的檔案跟試算表一樣看得到色塊，
   * 就得輸出真正的 .xlsx。.xlsx 其實是一個 zip 包著幾個 XML，
   * 所以這裡自己寫最小可用的 zip 與 XML，不載入任何外部程式庫：
   * 院內網路連不連得到 CDN 都不影響，也少一個要跟著更新的相依套件。
   *
   * 顏色與文字色沿用網頁上同一套判斷（colorOf / inkOn），
   * 所以 Excel 裡看到的配色與畫面上一致；欄序也與總覽表相同。
   * ------------------------------------------------------------------ */

  var CRC32_TABLE = (function () {
    var t = new Uint32Array(256), c, n, k;
    for (n = 0; n < 256; n++) {
      c = n;
      for (k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC32_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function utf8Bytes(str) { return new TextEncoder().encode(str); }

  /* 以「不壓縮（store）」方式打包 zip —— xlsx 允許不壓縮，
     省下一整個壓縮程式庫，課表這種大小的檔案也不需要壓。 */
  function zipStore(files) {
    var parts = [], central = [], offset = 0;

    // zip 的時間是 DOS 格式：日期 0 是無效值，有些工具會抱怨，所以填上現在時間
    var now = new Date();
    var dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();

    files.forEach(function (f) {
      var name = utf8Bytes(f.name), data = f.data, crc = crc32(data);

      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true);
      h.setUint16(4, 20, true);            // 需要的版本
      h.setUint16(6, 0x0800, true);        // 檔名以 UTF-8 編碼
      h.setUint16(8, 0, true);             // 0 = 不壓縮
      h.setUint16(10, dosTime, true);
      h.setUint16(12, dosDate, true);
      h.setUint32(14, crc, true);
      h.setUint32(18, data.length, true);
      h.setUint32(22, data.length, true);
      h.setUint16(26, name.length, true);
      parts.push(new Uint8Array(h.buffer), name, data);

      var c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true);
      c.setUint16(4, 20, true);
      c.setUint16(6, 20, true);
      c.setUint16(8, 0x0800, true);
      c.setUint16(10, 0, true);            // 0 = 不壓縮
      c.setUint16(12, dosTime, true);
      c.setUint16(14, dosDate, true);
      c.setUint32(16, crc, true);
      c.setUint32(20, data.length, true);
      c.setUint32(24, data.length, true);
      c.setUint16(28, name.length, true);
      c.setUint32(42, offset, true);       // 這個檔案的起始位置
      central.push(new Uint8Array(c.buffer), name);

      offset += 30 + name.length + data.length;
    });

    var cdSize = central.reduce(function (n, b) { return n + b.length; }, 0);
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);

    return new Blob(parts.concat(central, [new Uint8Array(end.buffer)]),
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function xmlEsc(s) {
    return String(s == null ? '' : s)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')     // 控制字元會讓 Excel 拒開
      .replace(/[&<>"']/g, function (c) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c];
      });
  }

  /* #FFC000 → FFFFC000（Excel 的顏色是 AARRGGBB） */
  function argb(hex) {
    var h = String(hex || '').replace('#', '');
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    return 'FF' + (h.length === 6 ? h.toUpperCase() : 'FFFFFF');
  }

  /* 0 → A、25 → Z、26 → AA */
  function colName(i) {
    var s = '';
    for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + (i - 1) % 26) + s;
    return s;
  }

  function exportXlsx() {
    var d = state.data, months = d.monthCols, rows = filteredRows();
    var fixed = FIXED_ORDER.filter(function (c) { return d.fixedCols.indexOf(c) >= 0; });
    var headers = fixed.concat(months.map(function (m) { return m.label; }));

    /* --- 樣式：一種底色一個 fill；fill 0、1 是 Excel 規定的固定兩格 --- */
    var fills = [
      '<fill><patternFill patternType="none"/></fill>',
      '<fill><patternFill patternType="gray125"/></fill>',
      '<fill><patternFill patternType="solid"><fgColor rgb="FFF1F3F5"/><bgColor indexed="64"/></patternFill></fill>'
    ];
    var xfs = [
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>',
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>'
    ];
    var HEAD_XF = 1, TEXT_XF = 2;

    var colorXf = {};
    function xfForColor(hex) {
      var key = argb(hex);
      if (key in colorXf) return colorXf[key];
      fills.push('<fill><patternFill patternType="solid"><fgColor rgb="' + key +
        '"/><bgColor indexed="64"/></patternFill></fill>');
      // 底色深就用白字，與網頁上同一個判斷，不會出現黑底黑字
      xfs.push('<xf numFmtId="0" fontId="' + (inkOn(hex) === INK_LIGHT ? 3 : 1) +
        '" fillId="' + (fills.length - 1) + '" borderId="1" xfId="0"' +
        ' applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
        '<alignment horizontal="center" vertical="center"/></xf>');
      colorXf[key] = xfs.length - 1;
      return colorXf[key];
    }

    function cell(ref, xf, v) {
      v = String(v == null ? '' : v);
      if (!v) return '<c r="' + ref + '" s="' + xf + '"/>';
      return '<c r="' + ref + '" s="' + xf + '" t="inlineStr"><is><t xml:space="preserve">' +
        xmlEsc(v) + '</t></is></c>';
    }

    var body = ['<row r="1" ht="22" customHeight="1">' + headers.map(function (h, i) {
      return cell(colName(i) + '1', HEAD_XF, h);
    }).join('') + '</row>'];

    rows.forEach(function (r, ri) {
      var rn = ri + 2;
      var cells = fixed.map(function (c, i) { return cell(colName(i) + rn, TEXT_XF, r[c]); });
      months.forEach(function (m, i) {
        var mc = r.months[m.key], ref = colName(fixed.length + i) + rn;
        cells.push(mc ? cell(ref, xfForColor(colorOf(mc)), mc.value) : cell(ref, TEXT_XF, ''));
      });
      body.push('<row r="' + rn + '">' + cells.join('') + '</row>');
    });

    var head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

    /* 凍結表頭那一列與左邊的固定欄，捲到右邊的月份時還看得出是誰 */
    var sheetXml = head +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<dimension ref="A1:' + colName(headers.length - 1) + (rows.length + 1) + '"/>' +
      '<sheetViews><sheetView workbookViewId="0"><pane xSplit="' + fixed.length + '" ySplit="1" topLeftCell="' +
      colName(fixed.length) + '2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="18"/>' +
      '<cols><col min="1" max="' + fixed.length + '" width="11" customWidth="1"/>' +
      '<col min="' + (fixed.length + 1) + '" max="' + headers.length + '" width="12" customWidth="1"/></cols>' +
      '<sheetData>' + body.join('') + '</sheetData>' +
      '<pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>' +
      '</worksheet>';

    var stylesXml = head +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="4">' +
        '<font><sz val="11"/><color rgb="FF1B2432"/><name val="Calibri"/></font>' +
        '<font><b/><sz val="11"/><color rgb="FF1B2432"/><name val="Calibri"/></font>' +
        '<font><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
      '</fonts>' +
      '<fills count="' + fills.length + '">' + fills.join('') + '</fills>' +
      '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right>' +
      '<top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom>' +
      '<diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/>' +
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>' +
      '</styleSheet>';

    // Excel 的工作表名稱上限 31 字，且不能有 : \ / ? * [ ]
    var tabName = xmlEsc(String(d.sheet || '課程表').replace(/[:\\\/?*\[\]]/g, '-').slice(0, 31));

    var pkg = [
      { name: '[Content_Types].xml', data: utf8Bytes(head +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>') },
      { name: '_rels/.rels', data: utf8Bytes(head +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>') },
      { name: 'xl/workbook.xml', data: utf8Bytes(head +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="' + tabName + '" sheetId="1" r:id="rId1"/></sheets></workbook>') },
      { name: 'xl/_rels/workbook.xml.rels', data: utf8Bytes(head +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>') },
      { name: 'xl/styles.xml', data: utf8Bytes(stylesXml) },
      { name: 'xl/worksheets/sheet1.xml', data: utf8Bytes(sheetXml) }
    ];

    saveBlob(zipStore(pkg), fileBase() + '.xlsx');
  }

  /* ------------------------------------------------------------ 載入 */
  function loadSample() {
    state.data = normalize(JSON.parse(JSON.stringify(window.PGY_SAMPLE_DATA || { rows: [] })));
    state.sheets = [];
  }

  function load(sheetName) {
    state.source = 'loading';
    state.error = '';
    renderStatus();

    if (!(CFG.GAS_WEB_APP_URL || '').trim()) {
      loadSample();
      state.source = 'sample';
      render();
      return Promise.resolve();
    }

    return fetchFromGas(sheetName)
      .then(function (raw) {
        state.data = normalize(raw);
        state.sheets = raw.sheets || [];
        state.source = 'sheets';
      })
      .catch(function (e) {
        loadSample();
        state.source = 'error';
        state.error = '無法讀取 Google Sheets：' + (e && e.message ? e.message : e);
      })
      .then(render);
  }

  /* ========================= 公告欄（留言板） =========================
   *  內容來自 Google Sheets 的另一個工作表（config.js 的 NOTICE_SHEET，
   *  預設叫「公告」），教學部在試算表上填一列，兩頁就同步看得到。
   *
   *  這一塊刻意與課表完全獨立：沒建那個工作表、一則都沒有、
   *  或連線失敗時，#board 整塊隱藏，頁面跟沒有這個功能時一模一樣，
   *  也不會讓課表跟著顯示錯誤。
   * ================================================================= */

  var board = { notices: [], expanded: false };

  /* 還沒設定 GAS 網址（離線範例模式）時拿來示範版面用的兩則。
   * 一旦填了 GAS_WEB_APP_URL，內容一律以試算表的「公告」工作表為準，
   * 這裡的字不會出現在正式頁面上。 */
  var DEMO_NOTICES = [
    {
      date: '2026/09/01', tag: '排程異動', pin: true, author: '教學部',
      title: '（範例）9 月急診梯次調整',
      body: '原排在 9 月急診的第 3 組，改到 10 月；9 月改為內科。\n細節請看試算表，如有疑問請與教學部聯絡。'
    },
    {
      date: '2026/08/20', tag: '提醒', author: '教學部',
      title: '（範例）選修志願調查將於月底截止',
      body: '這一塊是「公告欄」的示範內容 —— 正式使用時請在 Google Sheets 新增一個叫「公告」的工作表，一則公告填一列。'
    }
  ];

  /* 這個外殼是哪一頁：試算表「對象」欄要靠它決定某一則要不要出現。
   * office.html 的 <body> 有 class="office"，index.html 沒有。 */
  function pageRole() {
    return document.body.classList.contains('office') ? 'office' : 'admin';
  }

  /* 對象欄：留空或「全部」= 兩頁都顯示；
   * 只寫到管理者類的字 = 只有 index.html；只寫到公開類的字 = 只有 office.html。
   * 兩類都寫到（或看不懂）就當作全部，寧可多顯示也不要讓公告憑空消失。 */
  function noticeForThisPage(n) {
    var a = String(n.audience || '').trim();
    if (!a || /^(全部|兩頁|all|both)$/i.test(a)) return true;
    var admin = /管理|權限|內部|index|admin/i.test(a);
    var office = /公開|科部|助理|外部|office/i.test(a);
    if (admin && !office) return pageRole() === 'admin';
    if (office && !admin) return pageRole() === 'office';
    return true;
  }

  function normalizeNotices(list) {
    var out = (list || []).map(function (n, i) {
      return {
        date: String(n.date == null ? '' : n.date).trim(),
        ts: Number(n.ts) || 0,
        tag: String(n.tag == null ? '' : n.tag).trim(),
        title: String(n.title == null ? '' : n.title).trim(),
        body: String(n.body == null ? '' : n.body).trim(),
        author: String(n.author == null ? '' : n.author).trim(),
        pin: !!n.pin,
        audience: String(n.audience == null ? '' : n.audience).trim(),
        seq: i
      };
    }).filter(function (n) { return (n.title || n.body) && noticeForThisPage(n); });

    // 後端讀不出時間戳時（例如日期是手打的文字），前端再試著解析一次，
    // 不然所有公告都會當成「沒填日期」而照試算表的列順序排。
    out.forEach(function (n) {
      if (n.ts || !n.date) return;
      var m = n.date.match(/(\d{4})\s*[\/\-.年]\s*(\d{1,2})(?:\s*[\/\-.月]\s*(\d{1,2}))?/);
      if (m) n.ts = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3] || 1)).getTime();
    });

    // 置頂優先，其次日期新的在前；沒填日期的排最後（ts = 0），順序照試算表
    out.sort(function (a, b) {
      if (a.pin !== b.pin) return a.pin ? -1 : 1;
      if (a.ts !== b.ts) return b.ts - a.ts;
      return a.seq - b.seq;
    });
    return out;
  }

  /* 已讀：只記在這台電腦的瀏覽器裡，不寫回試算表。
   * 用「日期＋標題＋內容開頭」當識別碼，改過內容就會重新變成未讀。 */
  var SEEN_KEY = 'pgy-board-seen';

  function noticeId(n) {
    return (n.date || '') + '|' + (n.title || '') + '|' + (n.body || '').slice(0, 40);
  }
  function seenIds() {
    try {
      var a = JSON.parse(localStorage.getItem(SEEN_KEY) || '[]');
      return Object.prototype.toString.call(a) === '[object Array]' ? a : [];
    } catch (e) { return []; }
  }
  function markAllSeen() {
    var ids = seenIds().concat(board.notices.map(noticeId));
    try { localStorage.setItem(SEEN_KEY, JSON.stringify(uniq(ids).slice(-300))); } catch (e) {}
  }

  /* 內容維持使用者在試算表裡打的斷行；網址自動變成可點的連結。
   * 先 esc 再處理，貼進來的內容不會變成 HTML。 */
  function noticeText(s) {
    return esc(s)
      .replace(/(https?:\/\/[^\s<]+)/g, function (u) {
        return '<a href="' + u + '" target="_blank" rel="noopener noreferrer">' + u + '</a>';
      })
      .replace(/\r?\n/g, '<br>');
  }

  function renderBoard() {
    var box = $('#board');
    if (!box) return;

    var list = board.notices;
    if (!list.length) { box.hidden = true; box.innerHTML = ''; return; }

    var seen = {};
    seenIds().forEach(function (id) { seen[id] = 1; });
    var unread = list.filter(function (n) { return !seen[noticeId(n)]; }).length;

    // 收合時至少顯示所有置頂的那幾則，置頂了卻被折起來就沒有意義
    var pinned = list.filter(function (n) { return n.pin; }).length;
    var preview = Math.max(Number(CFG.NOTICE_PREVIEW) || 0, pinned, 1);
    var shown = board.expanded ? list : list.slice(0, preview);
    var rest = list.length - shown.length;

    var tools = '';
    if (unread) tools += '<button class="btn ghost sm" id="board-seen">標記已讀</button>';
    if (rest > 0) tools += '<button class="btn ghost sm" id="board-more">顯示全部 ' + list.length + ' 則</button>';
    else if (board.expanded && list.length > preview) tools += '<button class="btn ghost sm" id="board-less">收合</button>';

    box.innerHTML =
      '<div class="board-head">' +
        '<h2>📌 公告欄' +
          (unread ? '<span class="board-n">' + unread + ' 則未讀</span>' : '') +
        '</h2>' +
        '<div class="board-tools">' + tools + '</div>' +
      '</div>' +
      '<ul class="board-list">' + shown.map(function (n) {
        var id = noticeId(n);
        // 公告日期放在整則的最前面 —— 排程改動是「哪一天發的」最重要，
        // 放在內容下方的小字會被略過。張貼者留在下方那一行。
        return '<li class="board-item' + (n.pin ? ' is-pin' : '') + (seen[id] ? '' : ' is-new') + '">' +
          '<div class="board-line">' +
            (n.date ? '<time class="board-date" title="公告日期">' + esc(n.date) + '</time>' : '') +
            (n.pin ? '<span class="board-pin">置頂</span>' : '') +
            (n.tag ? '<span class="board-tag">' + esc(n.tag) + '</span>' : '') +
            (n.title ? '<b class="board-title">' + esc(n.title) + '</b>' : '') +
            (seen[id] ? '' : '<span class="board-new">NEW</span>') +
          '</div>' +
          (n.body ? '<div class="board-text">' + noticeText(n.body) + '</div>' : '') +
          (n.author ? '<div class="board-meta">' + esc(n.author) + '</div>' : '') +
        '</li>';
      }).join('') + '</ul>';
    box.hidden = false;
  }

  function loadNotices() {
    var box = $('#board');
    if (!box) return Promise.resolve();

    var sheet = String(CFG.NOTICE_SHEET == null ? '公告' : CFG.NOTICE_SHEET).trim();
    if (!sheet) { box.hidden = true; return Promise.resolve(); }   // 設成空字串 = 關掉公告欄

    if (!(CFG.GAS_WEB_APP_URL || '').trim()) {                     // 離線範例模式
      board.notices = normalizeNotices(window.PGY_SAMPLE_NOTICES || DEMO_NOTICES);
      renderBoard();
      return Promise.resolve();
    }

    return callGas({ action: 'notices', sheet: sheet, t: Date.now() })
      .then(function (raw) { board.notices = normalizeNotices(raw && raw.notices); })
      .catch(function () { board.notices = []; })   // 公告讀不到就不顯示，不干擾課表
      .then(renderBoard);
  }

  /* ------------------------------------------------------------ 事件 */
  function bind() {
    on('#btn-refresh', 'click', function () {
      load(state.data ? state.data.sheet : '');
      loadNotices();
    });

    on('#btn-theme', 'click', function () {
      var cur = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', cur);
      try { localStorage.setItem('pgy-theme', cur); } catch (e) {}
    });

    on('#f-q', 'input', function (e) { state.filters.q = e.target.value.trim(); render(); });
    ['year:學年度', 'stage:期程', 'group:組別', 'mentor:長期導師'].forEach(function (pair) {
      var p = pair.split(':');
      on('#f-' + p[0], 'change', function (e) { state.filters[p[1]] = e.target.value; render(); });
    });
    on('#f-sheet', 'change', function (e) { load(e.target.value); });

    on('#btn-reset', 'click', function () {
      state.filters = { q: '', 學年度: '', 期程: '', 組別: '', 長期導師: '', cat: '' };
      state.personQuery = '';
      state.monthUnits = {};
      withEl('#f-q', function (el) { el.value = ''; });
      render();
    });

    on('#chips', 'click', function (e) {
      var b = e.target.closest('[data-cat]');
      if (!b) return;
      var k = b.dataset.cat;
      state.filters.cat = (state.filters.cat === k) ? '' : k;   // 單選；點已選的那個等於取消
      render();
    });

    on('#tabs', 'click', function (e) {
      var t = e.target.closest('.tab');
      if (!t) return;
      state.view = t.dataset.view;
      render();
    });

    on('#view', 'click', function (e) {
      var p = e.target.closest('[data-person]');
      if (p) {
        if (!hasView('person')) return;
        state.personQuery = p.dataset.person; state.view = 'person'; render(); return;
      }
      if (e.target.id === 'p-clear') {
        withEl('#p-q', function (el) { el.value = ''; el.focus(); });
        personQueryChanged('');
        return;
      }
      if (e.target.id === 'btn-xlsx') { exportXlsx(); return; }
      if (e.target.id === 'btn-csv') { exportCsv(); return; }
      if (e.target.id === 'btn-print') { window.print(); return; }
      var u = e.target.closest('[data-unit]');
      if (u) {
        var key = u.dataset.unit;
        state.monthUnits[key] = !state.monthUnits[key];
        render(); return;
      }
      if (e.target.id === 'u-clear') { state.monthUnits = {}; render(); return; }
      if (e.target.id === 'm-prev') { state.monthIndex = Math.max(0, state.monthIndex - 1); render(); return; }
      if (e.target.id === 'm-next') { state.monthIndex = state.monthIndex + 1; render(); return; }
    });

    on('#view', 'change', function (e) {
      if (e.target.id === 'm-pick') { state.monthIndex = Number(e.target.value); render(); }
    });

    /* 個人時程的查詢框。
     * 中文輸入法（注音、拼音、手寫）在選字前會先在欄位裡放一段「組字中」的文字，
     * 這段期間若把畫面重繪、輸入框被換掉，組字就會被打斷 ——
     * 欄位裡只剩沒組完的注音符號，看起來就像打字打不完整。
     * 所以：組字期間完全不動畫面，等 compositionend（選完字）再更新；
     * 平時打英數字也只重繪結果那一塊，輸入框自始至終不換掉。 */
    var composing = false;

    on('#view', 'compositionstart', function (e) {
      if (e.target.id === 'p-q') composing = true;
    });

    on('#view', 'compositionend', function (e) {
      if (e.target.id !== 'p-q') return;
      composing = false;
      personQueryChanged(e.target.value);
    });

    on('#view', 'input', function (e) {
      if (e.target.id !== 'p-q') return;
      if (composing || e.isComposing) return;      // 組字中，等選完字再說
      personQueryChanged(e.target.value);
    });

    on('#board', 'click', function (e) {
      var t = e.target;
      if (t.id === 'board-more') { board.expanded = true; renderBoard(); return; }
      if (t.id === 'board-less') { board.expanded = false; renderBoard(); return; }
      if (t.id === 'board-seen') { markAllSeen(); renderBoard(); return; }
    });

    window.addEventListener('resize', syncStickyOffset);
  }

  /* 查詢字串變了：只換結果那一塊，輸入框與游標位置都不受影響 */
  function personQueryChanged(value) {
    if (value === state.personQuery) return;
    state.personQuery = value;
    var res = $('#p-results');
    if (!res) { render(); return; }             // 不在個人時程頁就照舊整頁重繪
    res.innerHTML = personResults();
    withEl('#p-clear', function (b) { b.hidden = !String(value).trim(); });
  }

  /* ------------------------------------------------------------ 啟動 */
  function init() {
    try {
      var t = localStorage.getItem('pgy-theme');
      if (t) document.documentElement.setAttribute('data-theme', t);
    } catch (e) {}

    // 從自己的 script src 取出 ?v=，顯示在頁尾；版本只維護 index.html 一處
    var me = document.querySelector('script[src*="app.js"]');
    var ver = me && (me.getAttribute('src').split('v=')[1] || '');
    if (ver) withEl('#build', function (el) { el.textContent = '　·　版本 ' + ver; });

    // 標題以各頁 HTML 為準；只有 config 有填值時才覆寫，
    // 兩頁標題不同，不能被共用的設定檔蓋掉
    var title = String(CFG.TITLE || '').trim();
    var subtitle = String(CFG.SUBTITLE || '').trim();
    if (title) withEl('#title', function (el) { el.textContent = title; });
    if (subtitle) withEl('#subtitle', function (el) { el.textContent = subtitle; });
    document.title = title || (withEl('#title', function () {}) || {}).textContent || 'PGY 訓練課程表';

    bind();
    load();
    loadNotices();

    if (CFG.AUTO_REFRESH_MS > 0) {
      setInterval(function () {
        load(state.data ? state.data.sheet : '');
        loadNotices();
      }, CFG.AUTO_REFRESH_MS);
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
