// ================================================
// 2026南山人壽臺北國道馬拉松 前處理資料（待匯入完整成績）
// 生成時間：2026-03-23 00:00:00
// 目前僅建立 metadata 與擴充入口，尚未包含 histogram_5min / sorted_seconds
// ================================================

window.marathonData = window.marathonData || {};

window.marathonData['2026_freeway_tpe'] = {
  "metadata": {
    "event_id": "2026_freeway_tpe",
    "event_name": "2026南山人壽臺北國道馬拉松",
    "event_date": "2026-03-08",
    "generated_at": "2026-03-23 00:00:00",
    "total_participants": 0,
    "race_types": [
      "MA",
      "HM"
    ],
    "group_categories": [
      "ALL",
      "一般"
    ],
    "source_url": "https://www.bravelog.tw/contest/rank/2026030801",
    "data_status": "pending_results_ingest",
    "notes": "已將賽事納入 repo 與前端事件清單；待取得可程式化抓取的完整成績後再補 binsAndPr。",
    "data_structure": {
      "histogram_bin_size": "5min",
      "percentile_precision": "0.1%",
      "time_format": "HH:MM:SS"
    }
  },
  "binsAndPr": {}
};
