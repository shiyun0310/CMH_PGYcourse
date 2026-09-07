# PGY 訓練課程表（CMH_PGYcourse）

畢業後一般醫學訓練（PGY）年度輪訓課程表的前端網頁。
**資料一律以 Google Sheets 為準**，前端透過 Google Apps Script（GAS）Web App 讀取，
連線失敗或尚未設定時，會自動退回內建的離線範例資料。

---

## 一、檔案結構

```
index.html                 主頁面
check.html                 GAS 連線檢測頁（串接後先開這頁確認）
assets/config.js           ★ 唯一需要修改的設定檔（GAS 網址、色票）
assets/style.css           樣式（含深色模式、列印樣式）
assets/app.js              前端主程式
assets/sample-data.js      離線範例資料（由範例 course 表轉出）
gas/Code.gs                ★ Google Apps Script 後端（貼到試算表的 Apps Script）
```

---

## 二、試算表格式

第 1 列為表頭，欄位名稱與範例表完全一致：

| 長期導師 | 學年度 | 組別 | 期程 | 人事號 | 簡碼 | 受訓醫師 | 2026/08 | 2026/09 | … | 2028/07 |
|---|---|---|---|---|---|---|---|---|---|---|
| 胡婷勻 | 115 | 未分 | PGY1 | B508F1 | 2204 | 陳怡蓁 | 外 | 外 | … | |

規則：

- **前 7 欄為固定欄**，順序不拘（程式以「欄位名稱」比對，不是以位置）。
- **其後為月份欄**，格式 `yyyy/mm`（日期格式或文字皆可）。
  也接受 `2026-08`、`2026年8月`、民國 `115/08`。欄數不限，可以是 12 個月或 24 個月。
- **`受訓醫師` 為空的列視為空白列**，會自動略過。
- **儲存格底色會一起帶到網頁上**；沒有設底色時，才使用 `config.js` 的預設色票。
- 新增/刪除列、新增月份欄，網頁都會自動跟著變，不必改程式。

---

## 三、部署 GAS 後端

1. 開啟課程表 Google Sheets → **擴充功能 → Apps Script**
2. 把 `gas/Code.gs` 整份內容貼上，存檔
3. （選用）若腳本不是綁在該試算表上，把最上方的 `SPREADSHEET_ID` 填成試算表 ID
4. **部署 → 新增部署作業 → 類型選「網頁應用程式」**
   - 執行身分：**我**
   - 誰可以存取：**任何人**（若只給院內看，選「機構內的任何人」）
5. 複製「網頁應用程式」網址（`https://script.google.com/macros/s/AKfycb.../exec`）

### API

| 方法 | 參數 | 說明 |
|---|---|---|
| GET | `?action=schedule&sheet=工作表名稱` | 取得課表 JSON（`sheet` 可省略＝第一個工作表） |
| GET | `?action=sheets` | 取得所有工作表名稱 |
| GET | `?action=ping` | 健康檢查 |
| GET | 任一請求加 `&callback=fn` | 回傳 JSONP（繞過瀏覽器 CORS 限制） |
| POST | `{action:'updateCell', sheet, 簡碼, month:'2026-08', value:'外'}` | 寫回單一儲存格 |
| POST | `{action:'updateRow', sheet, 簡碼, months:{'2026-08':'外', …}}` | 寫回整列月份 |

回應格式：

```json
{
  "ok": true,
  "sheet": "11509(詩)",
  "sheets": ["11509(詩)", "11409(甲)"],
  "fixedCols": ["長期導師","學年度","組別","期程","人事號","簡碼","受訓醫師"],
  "monthCols": [{ "key":"2026-08", "label":"2026/08", "year":2026, "month":8 }],
  "rows": [{
    "長期導師":"胡婷勻", "學年度":"115", "組別":"未分", "期程":"PGY1",
    "人事號":"B508F1", "簡碼":"2204", "受訓醫師":"陳怡蓁",
    "months": { "2026-08": { "value":"外", "color":"#92d050" } }
  }],
  "updatedAt": "2026/09/07 10:00:00"
}
```

`Code.gs` 上方可調整的參數：`FIXED_COLS`、`HEADER_ROW`、`KEY_COL`、`ID_COL`、
`WRITE_TOKEN`（寫入保護權杖，留空＝不檢查）、`CACHE_SECONDS`（讀取快取秒數）。

試算表上也會多出「**PGY 課程表**」選單，可清除 API 快取與檢查表頭格式。

---

## 四、設定前端

編輯 `assets/config.js`：

```js
GAS_WEB_APP_URL: 'https://script.google.com/macros/s/AKfycb.../exec',
DEFAULT_SHEET:   '',        // 留空＝第一個工作表
TRANSPORT:       'auto',    // auto：先 fetch，被 CORS 擋住自動改用 JSONP
AUTO_REFRESH_MS: 0,         // 想自動同步就設 300000（5 分鐘）
```

同一檔案也可調整 `CATEGORY_COLORS`（預設科別色票）與 `CATEGORY_RULES`（值→科別的判斷前綴）。

> 若 GAS 網址填錯或未授權，網頁不會壞掉，會顯示紅色提示並改用離線範例資料。

### 連線檢測

串接後打開 **`check.html`**，它會依序測：fetch（CORS）→ JSONP 備援 → `action=sheets` → `action=schedule`，
並把實際回應內容印出來。四項都通過就表示可以正常使用；有失敗時頁面下方會列出對應的排查步驟。

常見狀況：

| 症狀 | 原因 |
|---|---|
| 第 1 項失敗、第 2 項通過 | fetch 被 CORS 擋住，程式自動改走 JSONP，**屬正常，不必處理** |
| 全部失敗，訊息含「登入頁」 | 部署時「誰可以存取」不是「任何人」 |
| 全部失敗，訊息為載入失敗 | 網址錯誤，或結尾是 `/dev` 而非 `/exec` |
| 資料是舊的 | 改過 `Code.gs` 後沒有重新部署新版本；或試算表選單「PGY 課程表 → 清除 API 快取」 |

---

## 五、發布網頁

純靜態頁面，任何靜態空間都可以：

- **GitHub Pages**：Settings → Pages → Branch 選 `main`、資料夾選 `/ (root)`
- **本機預覽**：`python3 -m http.server 8000`，開 http://localhost:8000

---

## 六、網頁功能

| 分頁 | 內容 |
|---|---|
| **總覽表** | 完整輪訓矩陣。表頭（年／月）與受訓醫師欄捲動時固定；點姓名可跳到個人時程 |
| **月份檢視** | 選定某個月，依訓練單位分組列出該月所有受訓醫師 |
| **個人時程** | 每位醫師一張卡片，逐月色塊時間軸 ＋ 各科月數統計 |
| **統計** | 受訓人數、排課完成率、各科別／各訓練單位人月數、逐月科別分布 |

其他：關鍵字搜尋、學年度／期程／組別／長期導師篩選、科別色塊篩選、
隱藏無排課月份、匯出 CSV（含 BOM，Excel 可直接開）、列印／另存 PDF、深色模式。
