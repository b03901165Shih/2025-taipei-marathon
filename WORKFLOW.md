# 專案架構與賽事資料更新流程

## 結論

這個網站的資料流程是先從賽事官方成績頁（目前為 BraveLog）抓下成績、存成 Excel，再將 Excel 預先處理成 JavaScript 資料檔，最後把該 JS 檔以 `<script>` 加入 `index.html`。瀏覽器只讀取靜態 JS，不會在使用者開啟頁面時重新爬蟲。

```text
BraveLog 成績頁
    ↓ Selenium
scrap_result.py
    ↓ .xlsx
excels/<賽事>_完整成績.xlsx
    ↓ pandas
extract_excel_result.py
    ↓ .js
<event_id>_data.js
    ↓ <script src="...">
index.html（Chart.js 圖表、排名與 PR 查詢）
```

## 檔案職責

| 檔案／資料夾 | 用途 |
| --- | --- |
| `scrap_result.py` | 以 Selenium 操作 BraveLog 的賽別、分組與分頁，擷取每位跑者的姓名、背號、賽別、分組及完賽時間，輸出完整成績 Excel。 |
| `excels/` | 已下載的原始整理成果。三個現有賽事 Excel 都在此資料夾。 |
| `events_config.json` | 前處理階段的賽事清單：事件 ID、顯示名稱、日期、Excel 名稱、賽別、成績來源 URL。 |
| `extract_excel_result.py` | 將 Excel 轉成網站可直接讀取的資料 JS；同時計算每組 5 分鐘分布與依秒數排序的成績陣列。 |
| `2025_tpe_data.js`、`2026_chartered_tpe_data.js`、`2026_freeway_tpe_data.js` | 目前三場比賽的前端資料檔；都會把資料寫到全域 `window.marathonData[event_id]`。 |
| `index.html` | 純靜態前端。依已載入的 `window.marathonData` 動態建立賽事／賽別／分組選單，顯示圖表、時間查排名與 PR 反查。 |
| `marathon_bins_and_pr.js` | 舊版或整合用的大型資料檔；目前 `index.html` 沒有載入它，現行網站以各賽事獨立的 `*_data.js` 為準。 |

## 爬蟲輸出的 Excel 結構

`scrap_result.py` 最後輸出的主工作表是「完整成績」。主要欄位為：

- `姓名`、`背號`
- `賽別`：成績卡本身列出的距離／賽別，前處理程式用這個欄位分組。
- `賽事類型`：爬蟲目前選到的賽事類型名稱。
- `分組`、`來源分組標籤`
- `完賽時間`、`完賽時間_td`
- `總排名`

爬蟲會針對頁面動態找到的每個賽別，再逐一切換分組並翻完所有分頁；資料排序後也會額外寫出「分組統計」、「賽事類型統計」及「賽事類型_分組統計」工作表。

## 前端資料格式

每個 `*_data.js` 的核心形式如下：

```js
window.marathonData = window.marathonData || {};
window.marathonData['2026_freeway_tpe'] = {
  metadata: { event_id: '2026_freeway_tpe', /* ... */ },
  binsAndPr: {
    '2026_freeway_tpe__MA__ALL': {
      histogram_5min: [/* 每 5 分鐘的人數 */],
      sorted_seconds: [/* 由快到慢的完賽秒數 */]
    }
  }
};
```

同一筆成績會被放入兩種群組：`賽別__ALL`（該賽別全場）與 `賽別__原始分組`。`index.html` 使用 `histogram_5min` 畫圖；使用已排序的 `sorted_seconds` 以二分搜尋計算排名與 PR，或由 PR 回推目標時間。

## 新增或更新一場賽事的操作順序

### 1. 設定並執行爬蟲

修改 `scrap_result.py` 頂端的兩個值：

```python
BASE_URL = 'https://www.bravelog.tw/contest/rank/<賽事代碼>'
output_file = '2026_<賽事名稱>_完整成績.xlsx'
```

執行：

```powershell
python scrap_result.py
```

所需套件：`selenium`、`beautifulsoup4`、`pandas`、`openpyxl`，以及可由 Selenium 啟動的 Chrome／ChromeDriver。完成後，將產生的 Excel 放到 `excels/`。

### 2. 登記賽事設定

在 `events_config.json` 的 `events` 陣列新增一筆，至少填入：

```json
{
  "id": "2026_example",
  "name": "2026 範例馬拉松",
  "excel": "2026_範例馬拉松_完整成績.xlsx",
  "date": "2026-12-31",
  "race_types": ["MA", "HM"],
  "source_url": "https://www.bravelog.tw/contest/rank/..."
}
```

`id` 必須是唯一且穩定的英文／數字底線 ID；它同時決定輸出的檔案名 `2026_example_data.js` 和前端資料鍵值。

### 3. Excel 轉成前端 JavaScript

`extract_excel_result.py` 會依設定逐場讀 Excel，產出 `<id>_data.js`。

現況注意：設定檔中的 `excel` 只有檔名，但現有 Excel 實際位於 `excels/`，而程式直接 `pd.read_excel(excel_path)`，不會自動加上 `excels/`。因此執行前請採其中一種方式：

1. 將 `events_config.json` 的 `excel` 值改為 `excels/檔名.xlsx`；或
2. 在 `extract_excel_result.py` 中讓 `excel_path` 指向 `excels/`；或
3. 在 `excels/` 目錄內執行腳本，並明確指定根目錄的設定檔。

建議採第 1 種，讓設定檔與實際檔案位置一致。之後在專案根目錄執行：

```powershell
python extract_excel_result.py
```

它會檢查必要欄位、排除無效時間，把時間轉成秒數，建立每 5 分鐘直方圖與排序陣列，再輸出 `<id>_data.js` 到執行所在位置。

### 4. 讓 `index.html` 載入新資料檔

在 `index.html` 的既有資料載入區加入一行：

```html
<script src="2026_example_data.js"></script>
```

必須放在頁面主程式 `<script>` 之前。前端初始化會自動偵測 `window.marathonData` 的所有賽事，所以不需要另外改選單、圖表或 PR 查詢邏輯。

### 5. 驗證

以本機伺服器或部署後頁面開啟 `index.html`，確認：

1. 新賽事出現在三個賽事選單中。
2. 每個賽別與分組都能繪製圖表。
3. 輸入一個 Excel 已知成績時，總場及分組排名合理。
4. PR 反查的時間存在於該分組的成績範圍內。

## 現行資料與載入關係

`index.html` 目前載入下列三場資料：

- `2025_tpe_data.js`：2025 台北馬拉松
- `2026_chartered_tpe_data.js`：2026 渣打台北公益馬拉松
- `2026_freeway_tpe_data.js`：2026 南山人壽臺北國道馬拉松

這與 `events_config.json` 及 `excels/` 中的三份成績檔一致。

## 維護提醒

- `scrap_result.py` 的 URL 與輸出檔名目前是手動寫死的；每換一場比賽都要改，且要確認 BraveLog 畫面結構、賽別／分組選單與分頁 selector 仍相容。
- 前處理程式認定 Excel 必備欄位為 `姓名`、`背號`、`賽別`、`賽事類型`、`分組`、`完賽時間`；爬蟲或人工整理時不可任意更名。
- `race_types` 主要是 metadata，前端實際可選賽別由輸出的 `binsAndPr` keys 自動推導。
- 網站使用大會時間（Official Time），與選手個人淨時間或官方個人名次可能有差異。
