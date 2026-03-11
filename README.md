# TSMC 台積電每日股市分析報告

自動化每日股市分析日報系統，涵蓋 TSMC（台積電，2330.TW / TSM NYSE）完整投資分析。

## GitHub Pages 線上閱覽

**https://tomwu229622.github.io/tsmc-daily-report/**

- 上方下拉選單選擇日期，下方即時載入該日報告
- 前 / 後一日導覽按鈕 + 鍵盤 `←` `→` 快速切換
- 最新 10 日快速選取 Pill
- 「全螢幕開啟」可在新分頁瀏覽完整報告

## 檔案說明

| 檔案 | 說明 |
|------|------|
| `index.html` | GitHub Pages 入口頁（日期選擇器） |
| `manifest.json` | 報告清單（自動維護） |
| `generate_tsmc_html_report.py` | HTML 日報產生器（每日執行） |
| `TSMC_日報_YYYY-MM-DD.html` | 每日 HTML 報告（含日期） |
| `TSMC_股市分析報告.xlsx` | Excel 分析報告（含每日更新記錄） |

## HTML 日報內容（13 個分析區塊）

1. **股價快照** — 台股 2330 + NYSE TSM 雙市場即時數據
2. **財務重點 KPI** — Q4 2025 / FY 2025 12 項關鍵指標
3. **三大法人買賣超** — 外資/投信/自營商 + 5日走勢
4. **技術分析** — RSI、MACD、KD、布林通道、MA5/20/60/120
5. **期貨選擇權** — 未平倉量、Put/Call Ratio、最大痛苦點
6. **國際重要新聞** — 6則深度分析（含影響評估）
7. **主要客戶動態** — 7大客戶佔比（NVIDIA 21%、Apple 17%...）
8. **產業生態系股價** — 14檔供應鏈/客戶/競爭對手股票
9. **法人評級彙整** — Goldman、JPM、Bernstein 等分析師評級
10. **財報倒計時** — 下次財報日期/共識預估/歷史超預期率
11. **晶圓廠稼動率** — 7個製程節點稼動率估算
12. **風險評估矩陣** — 地緣政治/法規/競爭等8項風險
13. **IR 行事曆** — 除息日/財報/技術論壇等重要日期

## 每日更新方式

每日只需修改 `generate_tsmc_html_report.py` 頂端的**每日更新區塊**：

```python
STOCK = { ... }      # 股價數據
INSTITUTIONAL = { ...}  # 三大法人
TECHNICAL = { ... }  # 技術指標
NEWS = [ ... ]       # 當日新聞
```

然後執行：

```bash
python generate_tsmc_html_report.py
```

即可產生 `TSMC_日報_YYYY-MM-DD.html`。

## 排程自動更新

透過 Claude Code 排程技能（`tsmc-daily-update`），每個工作日早上 9:07 自動：
1. 搜尋最新股價與新聞
2. 更新 `generate_tsmc_html_report.py`
3. 產生當日 HTML 日報
4. 更新 Excel 每日記錄

## 數據來源

- Yahoo Finance / CNBC / Investing.com（股價）
- 玩股網 / Goodinfo / TWSE（法人買賣超）
- TSMC Investor Relations（財報/法說會）
- Morgan Stanley / Goldman Sachs / JPMorgan（法人評級）
- TrendForce / Digitimes（產業分析）

## 免責聲明

> 本報告僅供參考，**不構成任何投資建議**。投資有風險，請自行評估並諮詢專業顧問。
