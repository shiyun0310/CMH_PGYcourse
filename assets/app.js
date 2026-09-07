/* =====================================================================
 *  PGY 訓練課程表 — 前端主程式
 *  資料來源：Google Sheets（透過 GAS Web App）；連線失敗時退回離線範例。
 * ===================================================================== */
(function () {
  'use strict';

  var CFG = window.PGY_CONFIG || {};
  var FIXED_ORDER = ['受訓醫師', '期程', '組別', '長期導師', '人事號', '簡碼', '學年度'];

  var state = {
    data: null,          // 正規化後的資料
    sheets: [],          // 試算表內所有工作表名稱
    source: 'sample',    // 'sheets' | 'sample'
    error: '',
    view: 'grid',
    monthIndex: 0,
    filters: { q: '', 學年度: '', 期程: '', 組別: '', 長期導師: '', cats: {} },
    hideEmptyMonths: true,
    focusPerson: ''
  };

  /* ---------------------------------------------------------------- 工具 */
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  /* 依背景色決定文字用黑或白，確保對比 */
  function inkOn(hex) {
    var h = String(hex || '').replace('#', '');
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    if (h.length !== 6) return '#1b2432';
    var f = function (v) { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
    var L = 0.2126 * f(parseInt(h.slice(0, 2), 16)) +
            0.7152 * f(parseInt(h.slice(2, 4), 16)) +
            0.0722 * f(parseInt(h.slice(4, 6), 16));
    return L > 0.42 ? '#1b2432' : '#ffffff';
  }

  /* 值 → 科別分類 */
  function categoryOf(v) {
    var s = String(v || '').trim();
    if (!s) return '';
    var rules = CFG.CATEGORY_RULES || [];
    for (var i = 0; i < rules.length; i++) {
      if (s.indexOf(rules[i].prefix) === 0) return rules[i].name;
    }
    return '其他';
  }

  /* 儲存格顏色
   *   COLOR_SOURCE = 'category'（預設）：一律依科別上色，同科別顏色必定一致
   *   COLOR_SOURCE = 'sheets'         ：Sheets 底色優先，沒有底色才用色票 */
  /* 科別名 → 顏色（例：'內科'） */
  function catColor(cat) {
    var map = CFG.CATEGORY_COLORS || {};
    return map[cat] || map['其他'] || '#d9d9d9';
  }

  /* 原始值 → 顏色（例：'內(Y2不分)' → 內科的黃色） */
  function valueColor(v) {
    return catColor(categoryOf(v));
  }

  function colorOf(cell) {
    if (CFG.COLOR_SOURCE === 'sheets' &&
        cell && cell.color && /^#[0-9a-f]{6}$/i.test(cell.color)) return cell.color;
    return valueColor(cell && cell.value);
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

  function fetchFromGas(sheetName) {
    var base = (CFG.GAS_WEB_APP_URL || '').trim();
    if (!base) return Promise.reject(new Error('尚未設定 GAS_WEB_APP_URL'));
    var params = { action: 'schedule', sheet: sheetName || CFG.DEFAULT_SHEET || '', t: Date.now() };
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
  function activeMonths() {
    var d = state.data;
    if (!state.hideEmptyMonths) return d.monthCols;
    var rows = d.rows;
    return d.monthCols.filter(function (m) {
      return rows.some(function (r) { return r.months[m.key]; });
    });
  }

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
    return Object.keys(state.filters.cats).some(function (k) { return state.filters.cats[k]; });
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
      var hit = Object.keys(r.months).some(function (k) { return f.cats[r.months[k].cat]; });
      if (!hit) return false;
    }
    return true;
  }

  function filteredRows() {
    return state.data.rows.filter(rowMatches);
  }

  /* 儲存格是否因科別篩選而變淡 */
  function cellDim(cell) {
    return anyCatSelected() && !state.filters.cats[cell.cat];
  }

  /* ============================== 畫面 ============================== */

  function renderStatus() {
    var el = $('#status');
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

  function renderFilters() {
    var d = state.data, f = state.filters;

    function fillSel(id, values, cur) {
      var sel = $(id);
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
    if (state.sheets.length) {
      sheetSel.parentNode.style.display = '';
      sheetSel.innerHTML = state.sheets.map(function (s) {
        return '<option value="' + esc(s) + '"' + (s === d.sheet ? ' selected' : '') + '>' + esc(s) + '</option>';
      }).join('');
    } else {
      sheetSel.parentNode.style.display = 'none';
    }

    $('#chips').innerHTML = catsInUse().map(function (c) {
      var col = catColor(c.name);
      return '<button class="chip' + (f.cats[c.name] ? ' on' : '') + '" data-cat="' + esc(c.name) + '">' +
        '<i class="sw" style="background:' + col + '"></i>' + esc(c.name) +
        '<i class="n">' + c.n + '</i></button>';
    }).join('') || '<span style="color:var(--ink-3);font-size:12.5px">尚無課程資料</span>';

    $('#f-empty').checked = state.hideEmptyMonths;
  }

  /* ------------------------------------------------------- 檢視：總覽表 */
  function viewGrid() {
    var d = state.data, months = activeMonths(), rows = filteredRows();
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
      h += '<th class="' + (i === 0 ? 'fx' : '') + '" rowspan="2">' + esc(c) + '</th>';
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
        var cls = i === 0 ? 'fx name' : 'dim';
        if (i === 0) {
          h += '<td class="' + cls + '"><button class="rowbtn" data-person="' + esc(r['受訓醫師']) + '">' +
            esc(r[c] || '—') + '</button></td>';
        } else {
          h += '<td class="' + cls + '">' + esc(r[c] || '—') + '</td>';
        }
      });
      months.forEach(function (m) {
        var cell = r.months[m.key];
        if (!cell) { h += '<td class="empty"></td>'; return; }
        var bg = colorOf(cell), fg = inkOn(bg), dim = cellDim(cell);
        h += '<td' + (dim ? ' style="opacity:.22"' : '') + '><span class="cell" style="background:' + bg +
          ';color:' + fg + '" title="' + esc(m.label + '　' + cell.value) + '">' + esc(cell.value) + '</span></td>';
      });
      h += '</tr>';
    });

    h += '</tbody></table></div>';
    return h;
  }

  /* ------------------------------------------------------- 檢視：月份 */
  function viewMonth() {
    var months = activeMonths(), rows = filteredRows();
    if (!months.length) return emptyState('沒有可顯示的月份');
    if (state.monthIndex >= months.length) state.monthIndex = 0;
    var m = months[state.monthIndex];

    var buckets = {};
    rows.forEach(function (r) {
      var c = r.months[m.key];
      if (!c) return;
      if (anyCatSelected() && !state.filters.cats[c.cat]) return;
      (buckets[c.value] = buckets[c.value] || []).push(r);
    });

    var keys = Object.keys(buckets).sort(function (a, b) {
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

    var assigned = keys.reduce(function (s, k) { return s + buckets[k].length; }, 0);

    var h = '<div class="view-head"><div><h2>' + esc(m.label) + ' 輪訓分布</h2>' +
      '<div class="sub">' + keys.length + ' 個訓練單位 ／ ' + assigned + ' 位受訓醫師已排定（篩選後共 ' + rows.length + ' 位）</div></div>' +
      nav + '</div>';

    if (!keys.length) return h + emptyState('本月沒有符合條件的排課');

    h += '<div class="mgrid">';
    keys.forEach(function (k) {
      var col = colorOf(buckets[k][0].months[m.key]), fg = inkOn(col);
      h += '<div class="mcard"><h3 style="background:' + col + ';color:' + fg + '">' +
        '<span>' + esc(k) + '</span><span class="count-badge">' + buckets[k].length + ' 人</span></h3><ul>';
      buckets[k].sort(function (a, b) { return String(a['受訓醫師']).localeCompare(String(b['受訓醫師']), 'zh-Hant'); })
        .forEach(function (r) {
          h += '<li><b>' + esc(r['受訓醫師']) + '</b>' +
            '<span>' + esc([r['期程'], r['簡碼']].filter(Boolean).join('・')) + '</span></li>';
        });
      h += '</ul></div>';
    });
    h += '</div>';
    return h;
  }

  /* ------------------------------------------------------- 檢視：個人 */
  function viewPerson() {
    var d = state.data, months = activeMonths();
    var rows = filteredRows();
    if (state.focusPerson) {
      rows = rows.filter(function (r) { return r['受訓醫師'] === state.focusPerson; });
    }
    if (!rows.length) return emptyState('沒有符合條件的受訓醫師');

    var h = '<div class="view-head"><div><h2>個人輪訓時程</h2>' +
      '<div class="sub">共 ' + rows.length + ' 位' + (state.focusPerson ? '（已鎖定 ' + esc(state.focusPerson) + '）' : '') + '</div></div>' +
      (state.focusPerson ? '<button class="btn" id="p-clear">顯示全部</button>' : '') + '</div><div class="plist">';

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
        h += '<div class="seg" style="background:' + bg + ';color:' + fg + (dim ? ';opacity:.22' : '') + '">' +
          '<span class="m">' + esc(m.label) + '</span><span class="v">' + esc(c.value) + '</span></div>';
      });

      h += '</div><div class="psum">';
      Object.keys(tally).sort(function (a, b) { return tally[b] - tally[a]; }).forEach(function (k) {
        var col = catColor(k);
        h += '<span class="s" style="background:' + col + ';color:' + inkOn(col) + '">' +
          esc(k) + ' ' + tally[k] + ' 月</span>';
      });
      h += '</div></div>';
    });

    return h + '</div>';
  }

  /* ------------------------------------------------------- 檢視：統計 */
  function viewStats() {
    var months = activeMonths(), rows = filteredRows();
    if (!rows.length) return emptyState('沒有符合條件的資料');

    var catTotal = {}, byMonth = {}, unitTotal = {}, slots = 0;
    rows.forEach(function (r) {
      months.forEach(function (m) {
        var c = r.months[m.key];
        if (!c) return;
        slots++;
        catTotal[c.cat] = (catTotal[c.cat] || 0) + 1;
        unitTotal[c.value] = (unitTotal[c.value] || 0) + 1;
        (byMonth[m.key] = byMonth[m.key] || {})[c.cat] = (byMonth[m.key][c.cat] || 0) + 1;
      });
    });

    var capacity = rows.length * months.length;
    var catKeys = Object.keys(catTotal).sort(function (a, b) { return catTotal[b] - catTotal[a]; });
    var maxCat = catKeys.length ? catTotal[catKeys[0]] : 1;

    var h = '<div class="view-head"><div><h2>統計總覽</h2><div class="sub">依目前篩選條件計算</div></div></div><div class="stats">';

    h += '<div class="stat-cards">' +
      card('受訓醫師', rows.length, '人') +
      card('涵蓋月份', months.length, '個月') +
      card('已排課月數', slots, '人月') +
      card('排課完成率', capacity ? Math.round(slots / capacity * 100) + '%' : '—', slots + ' / ' + capacity) +
      card('訓練單位', Object.keys(unitTotal).length, '種') +
      '</div>';

    h += '<h3 class="sec">各科別總人月數</h3><div class="bars">';
    catKeys.forEach(function (k) {
      var col = catColor(k);
      h += '<div class="bar"><span class="bt">' + esc(k) + '</span>' +
        '<span class="bw"><i class="bf" style="width:' + (catTotal[k] / maxCat * 100) + '%;background:' + col + '"></i></span>' +
        '<span class="bn">' + catTotal[k] + ' 人月</span></div>';
    });
    h += '</div>';

    h += '<h3 class="sec">逐月科別分布</h3>';
    months.forEach(function (m) {
      var b = byMonth[m.key] || {};
      var tot = Object.keys(b).reduce(function (s, k) { return s + b[k]; }, 0);
      h += '<div class="mrow-stat"><span class="ml">' + esc(m.label) + '</span><span class="stack">';
      if (!tot) {
        h += '<i style="width:100%;background:var(--line-2)"></i>';
      } else {
        catKeys.forEach(function (k) {
          if (!b[k]) return;
          var col = catColor(k);
          h += '<i style="width:' + (b[k] / tot * 100) + '%;background:' + col + '" title="' +
            esc(m.label + ' ' + k + ' ' + b[k] + ' 人') + '"></i>';
        });
      }
      h += '</span></div>';
    });

    var unitKeys = Object.keys(unitTotal).sort(function (a, b) { return unitTotal[b] - unitTotal[a]; });
    var maxUnit = unitKeys.length ? unitTotal[unitKeys[0]] : 1;
    h += '<h3 class="sec">各訓練單位人月數</h3><div class="bars">';
    unitKeys.forEach(function (k) {
      var col = valueColor(k);
      h += '<div class="bar"><span class="bt">' + esc(k) + '</span>' +
        '<span class="bw"><i class="bf" style="width:' + (unitTotal[k] / maxUnit * 100) + '%;background:' + col + '"></i></span>' +
        '<span class="bn">' + unitTotal[k] + ' 人月</span></div>';
    });
    h += '</div></div>';
    return h;

    function card(lbl, num, foot) {
      return '<div class="scard"><div class="lbl">' + esc(lbl) + '</div><div class="num">' +
        esc(num) + '</div><div class="foot">' + esc(foot) + '</div></div>';
    }
  }

  function emptyState(msg) {
    return '<div class="empty-state"><div class="big">🗓</div><div>' + esc(msg) + '</div></div>';
  }

  /* --------------------------------------------------------- 主要繪製 */
  function render() {
    renderStatus();
    renderBanner();
    if (!state.data) { $('#view').innerHTML = emptyState('資料載入中…'); return; }

    renderFilters();
    $$('.tab').forEach(function (t) { t.classList.toggle('on', t.dataset.view === state.view); });

    var body;
    if (state.view === 'grid') {
      var months = activeMonths(), rows = filteredRows();
      body = '<div class="view-head"><div><h2>' + esc(state.data.sheet || '輪訓總覽') + '</h2>' +
        '<div class="sub">' + rows.length + ' 位受訓醫師 ／ ' + months.length + ' 個月' +
        (state.data.updatedAt ? '　·　更新於 ' + esc(state.data.updatedAt) : '') + '</div></div>' +
        '<div style="display:flex;gap:8px"><button class="btn" id="btn-csv">匯出 CSV</button>' +
        '<button class="btn" id="btn-print">列印 / PDF</button></div></div>' + viewGrid();
    } else if (state.view === 'month') {
      body = viewMonth();
    } else if (state.view === 'person') {
      body = viewPerson();
    } else {
      body = viewStats();
    }
    $('#view').innerHTML = body;
    syncStickyOffset();
  }

  /* 讓月份表頭固定在年份表頭下方（高度隨字型變化）。
   * 注意：固定欄的 th 有 rowspan="2"，量它會得到兩列的高度，
   * 必須量沒有 rowspan 的「年份」th 才是年份列的真實高度。 */
  function syncStickyOffset() {
    var mrow = $$('.grid thead tr.mrow th');
    if (!mrow.length) return;
    var y = $('.grid thead tr.yrow th:not([rowspan])');
    var h = y ? y.getBoundingClientRect().height : 0;
    mrow.forEach(function (th) { th.style.top = h + 'px'; });
  }

  /* ------------------------------------------------------------ CSV */
  function exportCsv() {
    var d = state.data, months = activeMonths(), rows = filteredRows();
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

    var blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'PGY課程表_' + (d.sheet || 'export').replace(/[\\/:*?"<>|]/g, '') + '.csv';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
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

  /* ------------------------------------------------------------ 事件 */
  function bind() {
    $('#btn-refresh').addEventListener('click', function () {
      load(state.data ? state.data.sheet : '');
    });

    $('#btn-theme').addEventListener('click', function () {
      var cur = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', cur);
      try { localStorage.setItem('pgy-theme', cur); } catch (e) {}
    });

    $('#f-q').addEventListener('input', function (e) { state.filters.q = e.target.value.trim(); render(); });
    ['year:學年度', 'stage:期程', 'group:組別', 'mentor:長期導師'].forEach(function (pair) {
      var p = pair.split(':');
      $('#f-' + p[0]).addEventListener('change', function (e) { state.filters[p[1]] = e.target.value; render(); });
    });
    $('#f-sheet').addEventListener('change', function (e) { load(e.target.value); });
    $('#f-empty').addEventListener('change', function (e) { state.hideEmptyMonths = e.target.checked; render(); });

    $('#btn-reset').addEventListener('click', function () {
      state.filters = { q: '', 學年度: '', 期程: '', 組別: '', 長期導師: '', cats: {} };
      state.focusPerson = '';
      $('#f-q').value = '';
      render();
    });

    $('#chips').addEventListener('click', function (e) {
      var b = e.target.closest('[data-cat]');
      if (!b) return;
      var k = b.dataset.cat;
      state.filters.cats[k] = !state.filters.cats[k];
      render();
    });

    $('#tabs').addEventListener('click', function (e) {
      var t = e.target.closest('.tab');
      if (!t) return;
      state.view = t.dataset.view;
      render();
    });

    $('#view').addEventListener('click', function (e) {
      var p = e.target.closest('[data-person]');
      if (p) { state.focusPerson = p.dataset.person; state.view = 'person'; render(); return; }
      if (e.target.id === 'p-clear') { state.focusPerson = ''; render(); return; }
      if (e.target.id === 'btn-csv') { exportCsv(); return; }
      if (e.target.id === 'btn-print') { window.print(); return; }
      if (e.target.id === 'm-prev') { state.monthIndex = Math.max(0, state.monthIndex - 1); render(); return; }
      if (e.target.id === 'm-next') { state.monthIndex = state.monthIndex + 1; render(); return; }
    });

    $('#view').addEventListener('change', function (e) {
      if (e.target.id === 'm-pick') { state.monthIndex = Number(e.target.value); render(); }
    });

    window.addEventListener('resize', syncStickyOffset);
  }

  /* ------------------------------------------------------------ 啟動 */
  function init() {
    try {
      var t = localStorage.getItem('pgy-theme');
      if (t) document.documentElement.setAttribute('data-theme', t);
    } catch (e) {}

    $('#title').textContent = CFG.TITLE || 'PGY 訓練課程表';
    $('#subtitle').textContent = CFG.SUBTITLE || '';
    document.title = CFG.TITLE || 'PGY 訓練課程表';

    bind();
    load();

    if (CFG.AUTO_REFRESH_MS > 0) {
      setInterval(function () { load(state.data ? state.data.sheet : ''); }, CFG.AUTO_REFRESH_MS);
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
