"""
TSMC 台積電每日股市分析日報產生器 v2.0
執行：python generate_tsmc_html_report.py
輸出：TSMC_日報_YYYY-MM-DD.html + manifest.json (GitHub Pages)
"""
from datetime import date, timedelta, datetime
import json, os

TODAY = date.today().isoformat()

# ═══════════════════════════════════════════════════════════════════
#  每日更新區 — 每天只需修改此區塊
# ═══════════════════════════════════════════════════════════════════
STOCK = {
    # 台股 2330
    "tw_price":       "2,210",        # 2026-05-13 Wed 收盤（-45，-2.00% vs 5/12 NT$2,255；Apple 傳轉單英特爾 + 跨市場套利賣壓）
    "tw_change":      "-45",          # 漲跌點數（vs 5/12 NT$2,255）
    "tw_change_pct":  "-2.00%",       # 漲跌幅
    "tw_volume":      "50,200",       # 成交量（張，5/13 放量回檔）
    "tw_open":        "2,205",        # 5/13 跳空低開
    "tw_high":        "2,220",        # 5/13 盤中高
    "tw_low":         "2,205",        # 5/13 盤中低
    "market_cap":     "NT$57.3 兆",   # 5/13 收盤市值
    "pe_ratio":       "~32.6x",       # 台股 P/E（TTM，NT$2,210 基礎）
    "pb_ratio":       "~11.4x",       # 台股 P/B（股淨比，NT$2,210 基礎）
    "div_yield":      "~0.68%",       # 殖利率（NT$2,210 基礎）
    "52w_high_tw":    "2,345",        # 52 週高（2026-05-07 盤中新高）
    "52w_low_tw":     "780",          # 52 週低（2025-04-07）
    # NYSE TSM ADR
    "nyse_price":     "404.54",       # 5/12 Tue 最新收盤（5/13 美股盤中、Asia 9am 尚未收盤；-7.14，-1.73% vs 5/11 $411.68）
    "nyse_change":    "-7.14",        # 5/12 漲跌
    "nyse_change_pct":"-1.73%",       # 5/12 NYSE 收 $404.54、仍站穩 $400
    "52w_high_nyse":  "420.00",       # 52 週高（2026-05-07 盤中觸及 $420.00 創史高）
    "52w_low_nyse":   "134.25",       # 52 週低（2025-04-07）
    "pe_nyse":        "~33.4x",       # NYSE USD 基礎 P/E（TTM，估算，$404.54 基礎）
    "pb_nyse":        "~11.7x",
    "beta":           "1.18",
    "div_next":       "下次配息預計 2026-06",
    "ex_date":        "2026-04-09（Q2 除息完成，US$0.968/ADR）",
    "next_earnings":  "2026-07（Q2 2026 財報，7 月第三週預估）",
    # 資料日期備註
    "data_note":      "今日（2026-05-13 Wed）報告基準資料：台股 2330 5/13 收 NT$2,210（-45，-2.00% vs 5/12 NT$2,255）跌破 5MA、回測 20MA 區間；盤中區間 NT$2,205–2,220；量放至約 50,200 張（vs 5/12 量縮 39,800）；市值回落 NT$57.3 兆；NYSE TSM 最新收盤（5/12 Tue）$404.54（-7.14，-1.73% vs 5/11 $411.68）仍站穩 $400 心理整數、距 5/7 盤中史高 $420.00 累計 -3.7%；台美 ADR 溢價自 +9.6% 擴大至 +13.9%（台股回檔速度大於 ADR）  |  ⚠️ 今日下跌主因：Apple 傳考慮將部分晶片訂單轉至 Intel Foundry 18A 製程、客戶集中度風險升溫、爆量飆漲後獲利了結賣壓共振  |  ⭐ 利多支撐續存：Applied Materials (AMAT) 與 TSMC 矽谷 EPIC Center $5B AI 晶片研發合作（美國史上最大半導體 R&D 投資）；TSMC Arizona 5/12 董事會批准追加 $20B 資本注入（強化 Phase 2/3 擴張）；Sony 5/9 合資 CIS 熊本擴張；2nm 製程 2026-2028 產能 CAGR +70%、5 座 2nm 廠今年量產啟動  |  NVIDIA 對 TSMC 採購承諾揭露已逾 $95B（兩年前僅 $16B）、#1 客戶 22% 結構性強化  |  5/13 估三大法人合計賣超 -15,500 張（外資轉賣 -12,800、投信 -2,000、自營 -700），籌碼短線轉空  |  Barclays $470 對應 NT$2,450（剩 +10.9%）、共識目標 NT$2,320（剩 +5.0%）|  催化：NVDA 5/28 Q1 FY27 財報距今 15 天為下一最重要 AI 算力指標；TSMC 5 月月營收預計 6/10 前公布；明日 5/14 Thu 台股關注：能否守 NT$2,200 心理整數 + NT$2,190 20MA 雙重支撐、若失守恐測試 NT$2,150 第一支撐",
}

# 今日重點摘要（每日更新 — 白話文，供非專業讀者快速掌握）
SUMMARY = {
    # 整體信號：強烈看多 / 看多 / 中性偏多 / 中性 / 中性偏空 / 看空 / 強烈看空
    "sentiment":  "中性",
    # 一句話標題
    "headline":   "📉 台股 2330 5/13 跌破 5MA 收 NT$2,210（-2.00%）量放回檔；Apple 傳轉單英特爾衝擊；NYSE TSM 5/12 收 $404.54 仍站穩 $400",
    # 3–5 條白話重點（不用術語，每條約 30 字內）
    "bullets": [
        "📉 台股 2330 5/13 跌破 5MA 收 NT$2,210（-45，-2.00% vs 5/12 NT$2,255），量放至 50,200 張",
        "⚠️ 主因：Apple 傳考慮將部分晶片訂單轉至 Intel Foundry 18A、客戶集中度風險升溫",
        "🇺🇸 NYSE TSM 5/12 最新收 $404.54（-7.14，-1.73% vs 5/11 $411.68）仍站穩 $400 強勢區",
        "⭐ Arizona 追加 $20B 資本注入 + AMAT EPIC Center $5B 合作 + Sony CIS 合資三大利多續存",
        "🏦 5/13 估三大法人合計賣超 -15,500 張（外資 -12,800、投信 -2,000、自營 -700），籌碼轉空",
        "📅 NVDA 5/28 Q1 FY27 財報距今 15 天為下一最重要 AI 算力指標；TSMC 5 月營收 6/10 前公布",
        "🛡️ 短線關鍵支撐 NT$2,200 心理整數 / NT$2,190 20MA，反彈目標 NT$2,255 5MA 與 NT$2,290 前壓",
    ],
    # 一句結論
    "bottom_line": "今日（2026-05-13 Wed）報告摘要：台股 2330 5/13 在 5/12 短暫反彈後再度回測，收 NT$2,210（-45，-2.00% vs 5/12 NT$2,255），盤中區間 NT$2,205-2,220 緊縮無反彈、跌破 5MA NT$2,250 與心理整數 NT$2,250，量放至約 50,200 張（vs 5/12 量縮 39,800），市值回落至 NT$57.3 兆。NYSE TSM 最新收盤（5/12 Tue）$404.54（-7.14，-1.73% vs 5/11 $411.68）仍站穩 $400 強勢區，距 5/7 盤中 52 週高 $420.00 累計 -3.7%。台美 ADR 溢價自 +9.6% 擴大至 +13.9%（理論值 = 5 × 2,210 ÷ 31.10 = $355.30，溢價 = ($404.54-$355.30)/$355.30 = +13.9%），代表台股回檔速度大於 ADR、跨市場套利機會擴大。⚠️ 今日下跌主因：（1）市場傳出 Apple 考慮將部分晶片訂單轉至 Intel Foundry 18A 製程、客戶集中度風險升溫（Apple 為 TSMC 第 2 大客戶、佔比 17%）；（2）爆量飆漲後獲利了結賣壓共振；（3）4 月營收 MoM -1.1% 略低市場估利空尚未完全消化。⭐ 利多續存：（1）TSMC Arizona 子公司 5/12 董事會批准追加 $20B 資本注入、強化美國 Phase 2/3 擴張；（2）Applied Materials EPIC Center $5B AI 晶片研發合作（美國史上最大半導體 R&D 投資）；（3）Sony 5/9 合資 CIS 熊本擴張；（4）2nm 製程 2026-2028 產能 CAGR +70%、5 座 2nm 廠今年量產啟動；（5）NVIDIA 對 TSMC 採購承諾揭露已逾 $95B（兩年前僅 $16B）、確認 #1 客戶 22% 結構性強化。籌碼面：5/13 估三大法人合計賣超約 -15,500 張——外資轉賣 -12,800 張、投信賣 -2,000 張、自營 -700 張，短線籌碼轉空。基本面：Q1 2026 淨利 $18.2B（+58.3% YoY）、毛利率 66.2% 雙創史高；2026 全年 >30% 成長指引維持；CapEx $52-56B 高端不變；2nm 至 2028 年產能 CAGR +70%；A14/A16 製程路線完整。即將事件：（1）NVDA Q1 FY27 財報 5/28（距今 15 天，最重要 AI 算力指標）；（2）TSMC 5 月月營收預計 6/10 前公布；（3）Q2 法說會 7 月。技術面：5/13 收 NT$2,210 跌破 5MA NT$2,250 與前支撐 NT$2,235，回測 20MA 區間 NT$2,190-2,200；距共識目標 NT$2,320 剩 +5.0%、距 Barclays $470 對應 NT$2,450 剩 +10.9% 上行空間；RSI 估自 56 降至 48 跌至中軸下、MACD +1.5 紅柱續收斂逼近翻黑、KDJ K(58)/D(60)/J(54) 死叉訊號浮現；下方支撐 NT$2,200 心理整數 / NT$2,190 20MA / NT$2,150 第一支撐。31 位分析師全數看多，0 賣出。",
    # 重大警示（選填，空白則不顯示）
    "risk_alert":  "今日（5/13 Wed）關鍵觀察：（1）⚠️ Apple 傳考慮將部分晶片訂單轉至 Intel Foundry 18A 製程為今日下跌主因、需觀察客戶集中度風險是否擴散（Apple 佔比 17%、TSMC 第 2 大客戶）；（2）台股 5/13 跌破 5MA NT$2,250 + 前支撐 NT$2,235、明日（5/14 Thu）能否守 NT$2,200 心理整數 + NT$2,190 20MA 雙重支撐為關鍵；（3）若失守 NT$2,190 將測試 NT$2,150 第一支撐 / NT$2,100 60MA；（4）反彈目標 NT$2,255 5MA / NT$2,290 前壓 / NT$2,320 共識目標（剩 +5.0%）；（5）外資 5/13 轉賣 -12,800 張、投信賣壓重現為短線最大警訊；（6）ADR 溢價擴大至 +13.9%、台股下跌速度大於 ADR、跨市場套利空間擴大；（7）短線估值 P/E 32.6x 仍在 5 年 85% 高位、需財報持續驗證；（8）台幣 31.10 升值壓力結構性續存；（9）CoWoS/ABF 基板瓶頸 + 美中 AI 晶片出口管制持續；（10）NVDA 5/28 Q1 FY27 財報距今 15 天為最重要 AI 算力指標；（11）市值 NT$57.3 兆續守 NT$57 兆心理大關（對應股價 NT$2,200）",
}

# 三大法人買賣超（張）— 最新一日
INSTITUTIONAL = {
    "date":     "2026-05-13",
    "foreign":  -12800,   # 外資（5/13 估算，轉賣超）
    "trust":    -2000,    # 投信（5/13 估算，賣壓重現）
    "dealer":   -700,     # 自營商（5/13 估算，技術性賣超）
    # 近5日（由舊到新：5/7, 5/8, 5/11, 5/12, 5/13）
    "foreign_5d": [+28500, +8500, -18500, +5200, -12800],    # 5/13：轉賣 -12,800
    "trust_5d":   [+6800, +3200, -2500, +1200, -2000],       # 5/13：賣壓 -2,000
    "dealer_5d":  [+1500, +600, -800, +300, -700],           # 5/13：技術性賣超
    "foreign_ownership_pct": "74.8%",  # 外資持股比例（5/13 微降）
    "foreign_ownership_shares": "約 194.0 億股",
}

# 技術分析指標（需每日從看盤軟體更新）
TECHNICAL = {
    "rsi_14":       48.0,   "rsi_signal": "RSI 估 48.0（5/13 -2.00% 回檔後自 56 降至 48、跌破中軸 50），中性偏空、多頭氣勢轉弱、需觀察是否反彈守 45",
    "macd":         +1.5,   "macd_signal": "MACD 估 +1.5 紅柱續收斂逼近翻黑（自 +2.0 收窄），死叉訊號浮現、動能轉弱",   "macd_color": "#FF9800",
    "kdj_k":        58.0,   "kdj_d": 60.0, "kdj_j": 54.0, "kdj_signal": "KDJ K(58)/D(60)/J(54)，K 跌破 D 死叉訊號浮現、多頭結構鬆動",
    "ma5":       "2,250",   # 5MA 微降（含 5/13 數據）
    "ma20":      "2,195",   # 20MA 持續上升
    "ma60":      "2,110",
    "ma120":     "1,990",
    "bb_upper":  "2,375",   # 布林上軌
    "bb_mid":    "2,195",   # 布林中軌（20MA）
    "bb_lower":  "2,015",   # 布林下軌
    "vol_ma5":  "48,500",   # 5日均量（張，含 5/13 量放）
    "trend":    "台股 2330 5/13 Wed 在 5/12 +0.89% 短暫反彈後再度回測，收 NT$2,210（-45，-2.00% vs 5/12 NT$2,255），盤中區間 NT$2,205-2,220 緊縮無反彈、跌破 5MA NT$2,250 與前支撐 NT$2,235，量放至約 50,200 張（vs 5/12 量縮 39,800），市值回落至 NT$57.3 兆。NYSE TSM 最新收盤（5/12 Tue）$404.54（-7.14，-1.73% vs 5/11 $411.68）仍站穩 $400 強勢區、台美 ADR 溢價自 +9.6% 擴大至 +13.9%。技術面：RSI 估 48 跌至中軸下、MACD +1.5 紅柱續收斂逼近翻黑、KDJ K(58)/D(60)/J(54) 死叉訊號浮現；5/13 收 NT$2,210 跌破 5MA NT$2,250 + 跌破 NT$2,235 前支撐、回測 20MA(2,195) 區間、惟站穩 60MA(2,110) / 120MA(1,990) 中期多頭結構未破。下一觀察：（1）明日（5/14 Thu）關注能否守 NT$2,200 心理整數 + NT$2,190 20MA 雙重支撐；（2）若失守 NT$2,190 將測試 NT$2,150 第一支撐 / NT$2,100 60MA；（3）反彈目標 NT$2,255 5MA / NT$2,290 前壓 / NT$2,320 共識目標；（4）Barclays $470 對應 NT$2,450 剩 +10.9% 上行空間；（5）⚠️ Apple 轉單英特爾風險為短期最大利空、需觀察是否擴散；（6）NVDA 5/28 財報為下波最重要催化、Arizona $20B 資本注入 + AMAT $5B EPIC Center + Sony CIS 合資為結構性正面。",
    "support":  "2,200 / 2,190 / 2,150 (心理整數、20MA、第一支撐)",
    "resist":   "2,255 / 2,290 / 2,320",
    "note": "技術指標數值為估算，請以看盤軟體即時數據為準；5/13 回檔後 RSI 48 跌破中軸、KDJ K 跌破 D 死叉訊號浮現、Apple 轉單英特爾風險需密切觀察",
}

# 期貨與選擇權（來源：台灣期交所 TAIFEX）
DERIVATIVES = {
    "date":          "2026-05-13",
    "futures_price": "2,205",      # 台積電期貨近月（估，5/13 逆價差擴大）
    "futures_oi":    "123,500",    # 未平倉量（口）（估，5/13 微增）
    "futures_basis": "-5",         # 逆價差 -5 點（避險需求升溫）
    "call_oi":       "76,500",
    "put_oi":        "58,200",
    "pcr":           "0.76",       # Put/Call 比率自 0.64 升至 0.76（避險需求升溫）
    "pcr_signal":    "中性偏空（Put/Call 0.76 > 0.7，Put 部位回升、避險需求升溫）",
    "max_pain":      "2,230",      # 最大痛苦價格（近期到期，5/13 下調）
    "key_strikes": [
        ("2,100", "Call 26,500 / Put 11,200", "60MA 支撐"),
        ("2,150", "Call 28,500 / Put 10,500", "第一支撐"),
        ("2,180", "Call 30,800 / Put 9,500", "前 20MA 支撐"),
        ("2,200", "Call 32,500 / Put 8,800", "心理整數支撐"),
        ("2,210", "Call 33,200 / Put 8,200", "5/13 收盤"),
        ("2,230", "Call 34,500 / Put 7,500", "最大痛苦價"),
        ("2,255", "Call 36,200 / Put 6,800", "5/12 高反彈目標"),
        ("2,290", "Call 38,500 / Put 5,500", "前壓力"),
        ("2,320", "Call 39,200 / Put 4,800", "共識目標"),
        ("2,345", "Call 37,800 / Put 4,200", "5/7 盤中史高、強壓力"),
    ],
    "note": "衍生品數據需每日至台灣期交所(taifex.com.tw)更新；5/13 回檔後逆價差擴大至 -5、Put/Call 升至 0.76 避險需求升溫",
}

# 供應鏈與客戶股價（需每日更新）
ECOSYSTEM_STOCKS = [
    # 類別、股票代碼、名稱、收盤價、漲跌幅、備註（2026-05-12 美股收盤 + 台股 5/13 收盤）
    ("設備", "ASML",    "艾司摩爾",    "702.30", "-0.41%", "EUV 設備龍頭（5/12 微跌；TSMC $52-56B CapEx 高端確認）"),
    ("設備", "AMAT",    "應用材料",    "207.50", "+0.83%", "⭐ TSMC EPIC Center $5B AI 合作續發酵（美國史上最大半導體 R&D 投資）"),
    ("設備", "LRCX",    "科林研發",    "850.20", "-0.60%", "蝕刻設備（5/12 微跌；2nm 訂單長線結構強化）"),
    ("設備", "KLAC",    "科磊",        "635.80", "-0.42%", "量測設備（5/12 微跌；AI CapEx 利多基底續強）"),
    ("記憶", "000660.KS","SK 海力士",  "324,000 KRW", "-0.61%", "HBM4 供應（5/12 微跌；GB300 訂單續強）"),
    ("記憶", "MU",      "美光",        "111.50", "-1.50%", "HBM / DDR5（5/12 跟進整理；CSP CapEx 利多續存）"),
    ("封裝", "3711.TW", "日月光",      "268",    "-1.47%", "CoWoS 封裝（5/13 跟進回檔；2nm/CoWoS 訂單續強）"),
    ("客戶", "NVDA",    "輝達",        "215.20", "-1.19%", "#1 客戶 22%（5/12 高檔整理；5/28 Q1 FY27 財報為下一催化）"),
    ("客戶", "AAPL",    "蘋果",        "281.50", "+1.19%", "⚠️ #2 客戶 17%（5/12 反彈；傳考慮將部分晶片訂單轉至 Intel Foundry 18A）"),
    ("客戶", "AVGO",    "博通",        "421.80", "-0.89%", "#3 客戶 13%（5/12 高檔整理；Custom XPU 訂單續強）"),
    ("客戶", "AMD",     "超微",        "452.30", "-0.98%", "#4 客戶 8%（5/12 高檔整理；Q1 大超預期 + Q2 指引 $11.2B 利多續存）"),
    ("客戶", "2454.TW", "聯發科",      "2,395",  "-1.44%", "手機 SoC 客戶 9%；5/13 跟進回檔"),
    ("競爭", "INTC",    "英特爾",      "25.20",  "+5.66%", "⚠️ Intel Foundry 18A 爬坡 + Apple 訂單轉單傳聞利多；5/12 大漲"),
    ("競爭", "005930.KS","三星電子",   "57,300 KRW", "-0.35%", "Samsung Foundry 5/12 微跌；HBM4E NVIDIA 送樣中"),
    ("處分後", "ARM",   "Arm Holdings","216.50", "-0.37%", "TSMC 4/29 全數出脫後股價持穩（5/12 微跌）；Arm 自身基本面強勢"),
]

# 法人評級彙整
ANALYST_RATINGS = [
    ("Barclays",       "增持",     "US$470",   "2026-04-22", "⭐ AI 需求結構性 + A14/A16 領先；對應 NT$2,450，距 5/4 $401.61 剩 +17.0% 空間"),
    ("Goldman Sachs",  "強烈買入", "NT$2,330", "2026-04-16", "Conviction Buy，Q1 財報後上調；5/4 NT$2,275 飆漲後距現價剩 +2.4%"),
    ("JPMorgan",       "增持",     "NT$2,200", "2026-04-17", "N3/N2 需求強勁，毛利率改善，財報後上調"),
    ("Susquehanna",    "正面",     "US$425",   "2026-04",    "AI 基礎設施資本支出持續，財報後上調"),
    ("Bernstein",      "優於大盤", "US$360",   "2026-04",    "技術領先優勢穩固，財報後上調"),
    ("Morgan Stanley", "增持",     "NT$2,290", "2026-04",    "市場共識平均目標"),
    ("MarketScreener", "共識買入", "NT$2,320", "2026-05-04", "31 位分析師、0 位賣出（5/4 NT$2,275 飆漲後上行空間收斂至 +2.0%，需財報數據持續驗證）"),
]

CONSENSUS = {
    "rating":       "強烈買入",
    "avg_target_tw":"NT$2,320",  # 共識（Barclays 升評後）— 5/13 NT$2,210 距共識剩 +5.0%
    "high_target":  "NT$2,770",
    "low_target":   "NT$1,740",
    "avg_target_us":"US$525",    # 分析師上調目標（Barclays $470 上拉均值）
    "analyst_count": 31,
    "buy": 31, "hold": 0, "sell": 0,
    "upside":       "+5.0%",    # 從 NT$2,210（5/13 收盤）到 NT$2,320
}

# 財報倒計時
EARNINGS = {
    "next_date":   "2026-07（Q2 2026 財報，7 月第三週預估）",
    "quarter":     "Q2 2026（下季財報，預計 2026-07）",
    "rev_guide_lo":"$39.0B",   # Q2 指引（法說會 4/16 公布，遠超共識）
    "rev_guide_hi":"$40.2B",
    "rev_guide_mid":"$39.6B",
    "gm_guide":    "65–67%（Q2 指引，法說會公布）",
    "op_guide":    "55–57%（Q2 指引，估算）",
    "eps_consensus":"NT$22.50 / US$3.60（Q2 估，法說會後上修）",
    "last_beat":   "+8.0%（Q1 2026 實際 vs 共識，淨利超預期）",
    "beat_rate":   "100%（連續 10 季全部超預期，Q1 2026 再度確認）",
    # ⭐ Q1 2026 實際財報數據（4/16 法說會公布，已確認）
    "q1_rev_actual":"NT$1.134 兆（$35.9B，+35.1% YoY，+8.4% QoQ）",  # 超越共識
    "mar_rev":      "NT$415.2B（+45.2% YoY，+30.7% MoM，季度最強月）",
    "apr_rev":      "NT$410.73B（+17.5% YoY，-1.1% MoM，5/8 公布）；1-4 月累計 NT$1,544.83B（+29.9% YoY）；5/11 股價回檔反映 MoM 略低市場估 NT$415B+",
}

# 晶圓廠稼動率（業界分析師估算）
FAB_UTILIZATION = [
    ("N2 (2nm)",     "Taiwan",   "90–95%",  "HVM 啟動，良率爬升中",            "#4CAF50"),
    ("N3/N3P",       "Taiwan",   "~100%",   "滿載，NVIDIA/Apple 全包",          "#2196F3"),
    ("N5/N4",        "Taiwan",   "~100%",   "主力製程，AI GPU + 手機 SoC",       "#2196F3"),
    ("N7/N6",        "Taiwan",   "90–95%",  "成熟先進製程，AI 推論需求支撐",      "#4CAF50"),
    ("N16–28",       "Taiwan",   "75–85%",  "成熟製程，IoT / MCU 需求穩定",      "#FF9800"),
    ("N3",           "Arizona",  "初期爬坡","Phase 1 全面量產中（2nm Phase 2 建設中）","#FF9800"),
    ("N12",          "Japan",    "~90%",    "熊本廠 Phase 1 量產中",              "#4CAF50"),
]

# IR 行事曆
IR_CALENDAR = [
    ("2026-03-17", "✅ 除息完成",  "2026 Q1 現金股利 NT$1.50/股（US$0.968/ADR）已於 3/17 除息",  "#9E9E9E"),
    ("2026-04-09", "✅ 除息完成", "2026 Q2 現金股利 NT$1.50/股（US$0.968/ADR）已於 4/9 除息完成","#9E9E9E"),
    ("2026-04-10", "✅ 月營收公布","TSMC 4/10 公布 3 月合併月營收 NT$415.2B（+45.2% YoY，+30.7% MoM）；Q1 2026 合計 NT$1.134 兆（+35.1%），超越共識預期 NT$1.12 兆","#9E9E9E"),
    ("2026-04-14", "✅ 週一反彈","台股 2330 收 NT$2,030（+2.56%，vs 4/10 NT$1,950）；NYSE TSM $383.30（+3.15%，強勢反彈）","#9E9E9E"),
    ("2026-04-15", "✅ 財報前夕創高","台股 2330 收 NT$2,055（+1.26%），52 週新高；NYSE TSM $379.89（4/15 收盤）；財報法說會明日（4/16）","#9E9E9E"),
    ("2026-04-16", "✅ Q1 2026 財報法說會","台股 2330 收 NT$2,100（+2.19%），創 52 週新高；NYSE TSM $392.00（+3.19%，財報後）；Q1 淨利 NT$572.5B（$18.1B，+58% YoY）創歷史新高；Q2 指引 $39.0-40.2B（遠超共識 $37.5-39.0B）；全年 >30% 成長確認；CapEx 向 $52-56B 高端靠齊","#9E9E9E"),
    ("2026-04-17", "✅ 財報後深度獲利了結","台股 2330 實際收 NT$2,030（-2.64%，-55 點），深度獲利了結大於預期；NYSE TSM 實際 $370.50（+1.97%，美股反彈）；台美分歧明顯；毛利率 66.2%（+740bps YoY）創歷史新高確認；NVIDIA #1 客戶 22% 確認","#9E9E9E"),
    ("2026-04-20", "✅ 週末休市","週末非交易日（Sunday）；週末消息：Amazon Trainium 省數百億美元資本支出、Meta/Anthropic 擴大自研 AI 芯片均依賴 TSMC；TSMC $165B 美國投資承諾確認，關稅豁免框架公布","#9E9E9E"),
    ("2026-04-21", "✅ 三大法人淨買超 2,829 張","台股 2330 收 NT$2,050（vs 盤中 NT$2,055）；外資實際買超 +2,210 張、投信 +696 張、自營 -77 張，合計淨買超 2,829 張；外資持股比例 74.9%；NYSE TSM 4/21 實際收盤 $366.24（-1.15%，vs 4/17 $370.50）","#9E9E9E"),
    ("2026-04-22", "✅ 北美技術論壇+ADR 大漲","TSMC 2026 North America Technology Symposium 舉辦：A14 製程宣布（2028 量產）、A16 H2 2026 量產確認；台股 2330 收盤 NT$2,050（持平），加權指數收 37,878（+273）；NYSE TSM 爆量大漲 +5.30% 收 $387.72——Barclays 升評增持、目標 $470；TSMC Arizona 先進封裝廠（CoWoS）2029 啟用確認","#9E9E9E"),
    ("2026-04-23", "✅ 台股跟進 ADR、ASML 高 NA EUV 遞延","台股 2330 4/23 收盤 NT$2,080（+30，+1.46% vs 4/22 NT$2,050），跟進 NYSE 前日 +5.3% 大漲；NYSE TSM 4/23 收 $387.53（-0.05%，技術整理消化）；TSMC 4/22 宣布暫不採用 ASML 最新高 NA EUV 設備（€350M/台過貴）→ A13/N2U 製程路線改以設計協同突破，ASML 股價 4/23 承壓 -2.6%；Siemens 與 TSMC 合作 AI 晶片設計；2028 先進封裝將整合 10 大晶片 +20 HBM","#9E9E9E"),
    ("2026-04-24", "✅ 爆量飆漲創歷史新高、台美雙市共振","台股 2330 4/24 爆量飆漲收 NT$2,185（+105，+5.05%）創歷史新高；NYSE TSM 4/24 大漲 +5.17% 收 $402.46 首次站上 $400；SOX 收 10,513.66（+4.32%）寫 30 年新高","#9E9E9E"),
    ("2026-04-27", "🎯 盤中觸 NT$2,330 與股號神奇巧合、單日漲點 +145 史上最大","台股 2330 4/27 收 NT$2,265（+80，+3.66%）續創歷史新高；盤中一度衝至 NT$2,330 與股票代號完美巧合、單日漲點 +145 創台股史上最大紀錄；市值盤中突破 NT$60 兆、收盤 NT$58.7 兆雙創新高；台股加權指數同日 +1,262 點站上 40,194 突破 4 萬大關；NYSE TSM 4/27 收 $404.98 續創史高；4/27 外資反手賣超 TSMC 1.8940 萬張（前三名皆 ETF 調節）","#9E9E9E"),
    ("2026-04-28", "📉 OpenAI 不達標 WSJ 報導引爆 AI 晶片全面拋售","台股 2330 4/28 收 NT$2,215（-50，-2.21%）自史高拉回；NYSE TSM 4/28 收 $392.34（-3.12%）跌破 $400 關卡；WSJ 報導 OpenAI 週活用戶與月營收雙不達標，CFO Sarah Friar 警告恐難資助未來算力協議，引爆 SOX -3.2%、AMD -6%、ARM -8%、Broadcom -5%、Intel/Micron -4%、Oracle -7%、SoftBank -10%；4/28 籌碼：外資連 2 賣超 22,114 張擴大、投信連 4 買 +861、自營連 3 買 +500，三大法人合計賣超 20,752 張；量縮 25% 至 46,932 張屬健康整理","#9E9E9E"),
    ("2026-04-29", "✅ 連 3 日修正、ADR 領先止跌微反彈","台股 2330 4/29 收 NT$2,180（-35，-1.58%）連 3 日修正、跌破 5MA NT$2,250 與心理 NT$2,200 兩道支撐；NYSE TSM 4/29 微反彈收 $393.79（+0.37%）；外資連 3 賣超估 -8,500 張、投信連 5 買 +650、自營連 4 買 +320；TSMC 4/29 揭露全數出脫 Arm 持股 1.11M 股 @ $207.65 共 $231M","#9E9E9E"),
    ("2026-04-30", "📉 連 4 日修正、跌破 NT$2,150 第一支撐","台股 2330 4/30 收 NT$2,135（-45，-2.06%）連 4 日修正、跌破 NT$2,150 第一支撐，距 4/27 史高累計 -5.74%；NYSE TSM 4/30 估收 $396.06（連 2 反彈）；4/30 籌碼：外資 4 連賣 -21,183 張、投信連 6 買 +596、自營 +115，合計賣超 20,472 張；外資持股降至 74.9%；技術面 RSI 跌破 50 中軸至 46.5、KDJ J 接近超賣 25、MACD 瀕臨死叉","#9E9E9E"),
    ("2026-05-01", "🎌 台灣勞動節休市；NYSE TSM 連 3 反彈逼近 $400","5/1 Fri 台灣勞動節 TWSE 休市；NYSE TSM 5/1 收 $397.67（+1.61，+0.41%）連 3 日反彈逼近 $400 關卡，自 4/29 $393.79 累計 +0.99%；ADR 對台股溢價自 +13.1% 進一步擴大至 +16.2%；美股 AI 晶片股延續反彈：NVDA $940.20（+0.84%）、AMD +0.60%、AVGO +0.95%、SOX 10,540（+0.28%）連 3 日漲；OpenAI 雜訊基本消化、ADR 領先確認築底訊號","#9E9E9E"),
    ("2026-05-04", "✅ 爆量飆漲收 NT$2,275 創 52 週新高、加權指數首破 4 萬大關","台股 2330 5/4 Mon 爆量飆漲收 NT$2,275（+140，+6.56%）創 52 週新高、單日漲幅 2024 以來最大；台股加權指數同日 +1,778.51（+4.57%）史上首度收盤站穩 40,705.14 站上 4 萬大關；TSMC 市值大幅回升至約 NT$59 兆；NYSE TSM 5/4 收 $401.61（+0.99%）連 4 反彈收復 $400；催化：4 大美國 CSP（Google/AWS/MSFT/Meta）2026 AI CapEx 合計 $725B（+77% YoY）；5/4 籌碼估外資反手大買 +25,800 張、投信連 7 買 +1,820、自營連 6 買 +1,350、合計買超估 28,970 張","#9E9E9E"),
    ("2026-05-05", "✅ 量縮續強收 NT$2,250、盤中觸 NT$2,300 心理整數新高","台股 2330 5/5 Tue 收 NT$2,250（持平自 5/4）、盤中觸 NT$2,300 心理整數創 52 週新高；成交量量縮至 21,565 張屬爆量後健康整理；台股加權指數同日 +64.15（+0.16%）收 40,769.29 連 2 日站穩 4 萬；NYSE TSM 5/5 拉回收 $396.89（-1.18%）跌破 $400；AMD Q1 盤後大超預期 AH +15%（EPS $1.37、營收 $10.25B、Q2 指引 $11.2B）；TSMC 重啟桃園龍潭 Phase 3 fab 計畫（$16.9B 投資、A14/1.6nm 用地）","#9E9E9E"),
    ("2026-05-06", "🚀 TWii 創歷史新高、NYSE TSM +6.36% 創 52 週新高 $419.50","台股 2330 5/6 Wed 收 NT$2,250（持平 0.00%）盤中觸 NT$2,275 量能放大；台股加權指數同日 +369.56（+0.91%）爆量 1.4491 兆收 41,138.85 連 3 日創歷史新高；NYSE TSM 5/6 大漲 +6.36% 收 $419.50 盤中觸 52 週新高；AMD 5/6 飆漲 +25% 反映 Q1 大超預期、4 大美國 CSP 2026 AI CapEx $725B（+77% YoY）確認；5/6 全市場外資爆量大買 +751.05 億元、投信 +17.8 億、自營 -96.11 億，三大法人合計 +672.85 億","#9E9E9E"),
    ("2026-05-07", "🚀 台股爆量飆漲創史高 NT$2,310、TWii 首破 4 萬 2、市值衝 60.81 兆","台股 2330 5/7 Thu 爆量飆漲收 NT$2,310（+60，+2.67%）創歷史新高，盤中觸 NT$2,345（+95，+4.22%）史高、市值收盤 59.9 兆、盤中達 60.81 兆雙創高；台股加權指數同日 +794.93（+1.93%）收 41,933.78、盤中觸 42,156.06 首破 4 萬 2 大關；最後一盤爆 7,112 張賣單仍守住天價；NYSE TSM 5/7 高位整理收 $416.58（-0.70% vs 5/6 $419.50）、開 $418.09 高 $420.00（盤中 52 週新高）；5/7 全市場三大法人合計買超 +583.31 億——外資 +464.12 億、投信 +160.62 億、自營 -41.43 億，TSMC 為外資+投信買超第一名","#9E9E9E"),
    ("2026-05-08", "✅ 高檔整理收 NT$2,290、TSMC 4 月月營收公布 NT$410.73B","台股 2330 5/8 Fri 受 5/7 爆量飆漲後高檔換手壓力小幅整理，收 NT$2,290（-20，-0.87% vs 5/7 NT$2,310），量縮至約 48,500 張屬健康消化；NYSE TSM 5/8 連 2 日小幅整理收 $414.15（-2.43，-0.58% vs 5/7 $416.58）仍站穩 $410；⭐ 5/8 盤後 TSMC 公布 4 月合併月營收 NT$410.73 億：YoY +17.5%、MoM -1.1%，1–4 月累計 NT$1,544.83 億（+29.9% YoY）略低於市場估 NT$415B+ 但維持高速年增","#9E9E9E"),
    ("2026-05-11", "📉 跌破 5MA 收 NT$2,235、AMAT-TSMC $5B EPIC Center 合作宣布","台股 2330 5/11 Mon 受 5/8 盤後公布 4 月月營收 NT$410.73B（MoM -1.1%）略低於市場估 NT$415B+ 衝擊 + 連續飆漲後獲利了結賣壓共振，收 NT$2,235（-55，-2.40% vs 5/8 NT$2,290）跌破 5MA + 前壓轉支撐 NT$2,250；NYSE TSM 5/11 同步跟進收 $404.54（-1.69%）仍站穩 $400；⭐ 重大利多：Applied Materials (AMAT) 5/11 宣布與 TSMC 在矽谷 EPIC Center 共建 $5B AI 晶片研發合作中心、為美國史上最大半導體設備 R&D 投資、AMAT 同日大漲 +3.92%；5/11 估三大法人合計賣超 -21,800 張（外資轉賣 -18,500、投信連 10 買終止 -2,500、自營 -800）","#9E9E9E"),
    ("2026-05-12", "✅ 反彈收 NT$2,255、Arizona $20B 資本注入","台股 2330 5/12 Tue 在 5/11 重挫 -2.40% 後出現技術性反彈，收 NT$2,255（+20，+0.89% vs 5/11 NT$2,235），盤中區間 NT$2,210-2,280，量能縮減至約 39,800 張屬健康整理；NYSE TSM 5/12 收 $404.54（-1.73% vs 5/11 $411.68）仍站穩 $400；⭐ TSMC Arizona 子公司董事會 5/12 批准追加 $20B 資本注入、強化美國 Phase 2/3 擴張、為迄今最大單一資本承諾","#9E9E9E"),
    ("2026-05-13", "📉 今日：跌破 5MA 收 NT$2,210、Apple 傳轉單英特爾","今日（5/13 Wed）台股 2330 收 NT$2,210（-45，-2.00% vs 5/12 NT$2,255）跌破 5MA NT$2,250 + 前支撐 NT$2,235、量放至 50,200 張、市值回落 NT$57.3 兆；下跌主因：⚠️ 市場傳出 Apple 考慮將部分晶片訂單轉至 Intel Foundry 18A 製程、客戶集中度風險升溫（Apple 為 TSMC 第 2 大客戶、佔比 17%）+ 爆量飆漲後獲利了結賣壓共振；5/13 估三大法人合計賣超 -15,500 張（外資 -12,800、投信 -2,000、自營 -700）；技術面 RSI 降至 48 跌破中軸、KDJ K(58)/D(60)/J(54) 死叉訊號浮現、MACD +1.5 紅柱續收斂逼近翻黑；明日（5/14 Thu）關鍵：能否守 NT$2,200 心理整數 + NT$2,190 20MA 雙重支撐；ADR 溢價自 +9.6% 擴大至 +13.9%","#BBDEFB"),
    ("2026-05-28", "📅 NVIDIA Q1 2026 財報","NVIDIA 預計 5/28 公布 Q1 FY27 財報；GB300 出貨進度與 H2 指引為 TSMC #1 客戶風向標、最重要 AI 算力指標","#FF9800"),
    ("2026-06-10", "📅 TSMC 5 月月營收公布","TSMC 將於 6/10 前公布 5 月合併月營收；4 月 NT$410.73B 後續觀察 5 月能否回升至 NT$420B+","#FF9800"),
    ("2026-06",    "2026 技術論壇","North America Technology Symposium（已於 4/22 舉辦，宣布 A13/A12/N2U）",  "#9E9E9E"),
    ("2026-07",    "Q2 2026 財報", "預計 7 月第三週（估）",                                        "#FF9800"),
    ("2026-09",    "月營收持續",   "每月 10 日前公告上月合併營收",                                  "#2196F3"),
    ("2026-10",    "Q3 2026 財報", "預計 10 月第三週（估）",                                        "#FF9800"),
]

# 總體經濟監控
MACRO = [
    ("SOX 半導體指數",  "10,895.30","+0.69%", False),  # 5/12 SOX 微反彈
    ("台股加權指數",    "40,820.50","-0.80%", False),  # 5/13 加權指數跟進回檔
    ("VIX 恐慌指數",    "18.2",    "+0.4",   False),   # 恐慌指數微升
    ("WTI 原油",        "$64.8",   "+0.47%", False),   # 油價微反彈
    ("10Y 美債殖利率",  "4.09%",   "+0.02pp",False),  # 殖利率微幅上揚
    ("台幣/美元",       "31.10",   "0.00",   False),  # 台幣持平（每升值 1% 毛利率 -40bps）
    ("NVIDIA NVDA",     "$215.20", "-1.19%", True),   # #1 TSMC 客戶（22%），5/12 微跌
    ("ASML",            "$702.30", "-0.41%", False),  # 設備股 5/12 微跌；TSMC $52-56B CapEx 上限確認
]

# ═══════════════════════════════════════════════════════════════════
#  進階分析師模組資料（每日同步更新）v3.0
# ═══════════════════════════════════════════════════════════════════

# 估值歷史位階 (Valuation Context)
VALUATION = {
    "pe_current":   32.6,   # 同 STOCK["pe_ratio"]（5/13 NT$2,210）
    "pe_5y_avg":    18.5,   # 5 年均 P/E（估算）
    "pe_5y_high":   37.8,   # 5 年高點（2021 AI 泡沫）
    "pe_5y_low":     9.8,   # 5 年低點（2022 下行周期）
    "pe_5y_pct":      85,   # 目前在 5 年區間的百分位（%，5/13 回檔後降至 85 百分位）
    "pe_10y_avg":   16.2,   # 10 年均 P/E（估算）
    "pe_10y_pct":     90,   # 目前在 10 年區間的百分位（%）
    "pb_current":  11.40,   # 同 STOCK["pb_ratio"]（5/13 NT$2,210）
    "pb_5y_avg":     5.8,
    "pb_5y_pct":      91,   # P/B 回檔後微降
    "ev_ebitda":    19.5,   # EV/EBITDA（估算）
    "peg":          0.99,   # PEG = P/E ÷ 5 年 EPS CAGR ~33%
    "note": "P/E 歷史百分位為估算（Bloomberg/Refinitiv 基準），請每季更新基準值；5/13 回檔後估值降至 5 年 85% 高位、需財報持續驗證",
}

# 現金流品質 (Quality of Earnings)
CASHFLOW = {
    "revenue_fy25":    "NT$3.8兆 ($122B)",  # 同 KPI 區塊
    "gm_fy25":         "56.1%（估）",        # FY25 毛利率（Q4 為 62.3%，全年估算）
    "op_margin_fy25":  "47.7%（估）",
    "capex_fy25":      "$34.5B USD（估）",   # 2025 實際資本支出
    "capex_pct_rev":   "28.3%（估）",        # CapEx / Revenue
    "fcf_fy25":        "NT$1兆+（FY25 KPI）",# 同 KPI 自由現金流
    "fcf_yield":       "~2.0%（NT$49.27T）", # FCF / 市值
    "doi_current":      82,                  # 存貨天數（Days of Inventory，估算）
    "doi_yoy":          -9,                  # 較去年同期變化（天）
    "ar_days":          43,                  # 應收帳款天數（估算）
    "ar_days_yoy":      -4,
    "inventory_trend": "去化中（連續 3 季下降）",
    "note": "FCF 來自 KPI 區塊；DOI / AR 為業界估算，建議至 TSMC 季報核實",
}

# 相對強弱 + ADR 折溢價 (Relative Strength + ADR Premium/Discount)
RELATIVE = {
    "date":           "2026-05-13",
    "tsmc_tw_1d":     -2.00,   # 2330 5/13 收盤 NT$2,210（-45 vs 5/12 NT$2,255）
    "taiex_1d":       -0.80,   # 加權指數 5/13 收 40,820（-330）
    "sox_1d":         +0.69,   # SOX 5/12 微反彈
    "tsmc_nyse_1d":   -1.73,   # TSM ADR 5/12 $404.54（-7.14 vs 5/11 $411.68）
    "tsmc_tw_5d":     -3.49,   # 5 日（NT$2,210 vs 5/6 NT$2,290 約 -3.5%）
    "taiex_5d":       -0.78,
    "tsmc_tw_1m":     +3.51,   # 月線（NT$2,210 vs ~NT$2,135 一個月前）
    "taiex_1m":       +5.8,
    "tsmc_tw_ytd":    +5.49,   # YTD 累計（NT$2,210 vs 年初 ~NT$2,095）
    "taiex_ytd":      +5.6,
    "sox_ytd":        -4.6,    # SOX 5/12 YTD 累計負值
    "beta_60d":        1.30,   # 60 日滾動 Beta vs TAIEX（5/13 波動率微升）
    # ADR 折溢價（以 5/12 NYSE 最新收盤計算）
    "adr_price":     404.54,   # TSM NYSE 5/12 最新收盤
    "adr_parity":    355.30,   # 理論值 = 5 × NT$2,210 ÷ 31.10
    "adr_ratio":     "5:1",    # 1 ADR = 5 台積電普通股
    "usd_ntd":        31.10,   # 同 MACRO 台幣匯率（5/13 持平）
    "adr_premium_pct": "+13.9",  # (404.54 - 355.30) / 355.30 × 100 = +13.9%（台股回檔速度大於 ADR、溢價擴大）
    "note": "ADR 理論值 = 5 × 台股收盤 ÷ 匯率（不含交易摩擦與流動性溢價）。5/13 台股 -2.00% 回檔、ADR 最新 (5/12) -1.73%，溢價自 +9.6% 擴大至 +13.9%；台美明顯分歧、跨市場套利機會擴大",
}

# CoWoS 先進封裝產能
COWOS = {
    "cowos_s_cap":  "~15,000",   # wspm（晶圓片/月）
    "cowos_l_cap":  "~5,000",
    "cowos_r_cap":  "~2,000",    # 2026 年爬坡
    "total_2025":   "~22,000",
    "target_2026":  "~35,000",
    "growth_yoy":   "+59%",
    "utilization":  ">95%",
    "backlog_qtrs":  4,           # 訂單積壓季數
    "bottleneck":   "ABF 基板與 HBM 供應限制（非 TSMC 封裝本身）",
    "customers": [
        ("NVIDIA",   "~65%", "GB200/GB300 NVL72", "#76B900"),
        ("AMD",      "~18%", "MI350/MI450",        "#ED1C24"),
        ("Broadcom", "~12%", "Custom XPU",         "#CC0000"),
        ("Others",   "~5%",  "各種 AI ASIC",       "#90CAF9"),
    ],
    "asp_note": "CoWoS 封裝 ASP 約為 HPC 裸晶圓的 2–3 倍，毛利率高於公司均值",
    "note": "數據來源：TrendForce 2026-Q1 預估；wspm = 晶圓片/月（均為估算）",
}

# 匯率敏感度 (FX Sensitivity)
FX_SENSITIVITY = {
    "spot":              31.10,  # 今日匯率（同 MACRO，5/11 持平）
    "gm_bps_per_pct":     -40,  # 台幣每升值 1%，毛利率下降 ~40bps（法說會指引）
    "scenarios": [
        (30.0, "NTD +3.5%（強升）", "約 -140 bps", "約 -4.9%"),
        (31.0, "NTD +0.3%（小升）", "約 -12 bps",  "約 -0.4%"),
        (31.10,"基準（5/11 收盤）", "—",           "—"),
        (32.0, "NTD -2.9%（小貶）", "約 +116 bps", "約 +4.0%"),
        (34.0, "NTD -9.3%（大貶）", "約 +372 bps", "約 +13.0%"),
    ],
    "hedge_ratio": "~75%（外幣避險比例估算）",
    "note": "TSMC 官方法說：新台幣每升值 1%，毛利率約下降 40bps（FY2025 法說會指引）",
}

# EPS 修正追蹤 (Earnings Revision Tracker)
EPS_REVISION = {
    "q1_consensus":  20.38,   # Q1 2026 EPS 共識（NT$）
    "q1_high":       22.10,
    "q1_low":        18.90,
    "q1_analysts":    32,
    "q1_vs_3m_ago":  "+6.7%",  # 較 3 個月前共識變化
    "q1_trend":      "上修",
    "fy_consensus":  84.50,    # FY 2026 EPS 共識
    "fy_high":       92.00,
    "fy_low":        78.00,
    "fy_vs_3m_ago":  "+8.1%",
    "fy_trend":      "上修",
    "beat_streak":    10,      # 連續超預期季數（Q1 2026 再度確認，連續 10 季）
    "beat_avg_pct":  "+6.2%",  # 近 10 季平均超預期幅度（Q1 2026 實際超預期 ~+8%）
    "next_report":   "2026-07（Q2 2026 財報）",
    "triggers": [
        ("正面", "2月月營收年增22.2%，催化法人上修EPS預估"),
        ("正面", "N3/N2 ASP調升，毛利率改善趨勢明確"),
        ("負面", "全球AI晶片出口管制擴大，訂單能見度風險"),
        ("中性", "台幣升值壓力，毛利率中性至下修風險"),
    ],
    "analyst_view": "Q1 2026 財報確認：淨利 NT$572.5B（$18.2B，+58.3% YoY）、毛利率 66.2%（+740bps YoY）雙創歷史新高，連續第 10 季超越市場預期（平均超預期幅度 +6.2%）。Q2 指引收入 $39.0-40.2B（遠超共識 $37.5-39.0B），全年 2026 收入成長指引上修至 >30%，CapEx 向 $52-56B 高端靠齊。近期價格：台股 2330 5/4 Mon 爆量飆漲收 NT$2,275（+140，+6.56% vs 4/30 NT$2,135）創 52 週新高、單日漲幅 2024 以來最大；台股加權指數同日 +1,778.51（+4.57%）史上首度收盤站穩 40,705.14 站上 4 萬大關，TSMC 市值大幅回升至約 NT$59 兆；NYSE TSM 5/4 收 $401.61（+0.99%）連 4 反彈、收復 $400 關卡。台美雙市共振、籌碼徹底翻多。本波反彈核心催化：（1）4 大美國 CSP（Alphabet、AWS、Microsoft、Meta）公布 2026 年 AI 基礎建設資本支出合計 $725B（+77% YoY），遠超預期；（2）OpenAI 不達標雜訊已被市場消化；（3）TSMC 北美技術論壇 4/22 宣布 A14/A16 製程路線、5 座 2nm 廠 2026 量產啟動延續催化。AI 晶片股延續反彈基礎：NVDA $952.80（+1.34%）、AMD $186.40（+0.97%）、AVGO $218.50（+1.72%）、SOX 收 10,795（+2.42%）強漲。5/4 籌碼徹底翻多：外資反手大幅買超 TSMC 估 +25,800 張（終結 4 連賣）、投信連 7 買 +1,820、自營連 6 買 +1,350，三大法人合計買超估 28,970 張，外資持股回升至 75.1%。即將事件：（1）AMD Q1 財報今日（5/5 Tue 盤後）為短線最大變數，市場預期 EPS $1.29（+34% YoY）、營收 $9.89B（+33% YoY）；（2）TSMC 4 月合併營收 5/10 前公布、預估 NT$330B+；（3）NVDA Q1 財報 5/28（最重要 AI 算力指標）。今日（2026-05-05 Tue）關注：（1）台股能否守住 NT$2,275 並挑戰盤中高 NT$2,290 / 心理 NT$2,300；（2）共識目標 NT$2,320 距現價剩僅 +2.0% 上行空間明顯收斂、Barclays $470 對應 NT$2,450 剩 +7.7%；（3）短線估值 P/E 33.6x 已回升至 5 年 88% 高位；（4）成交量爆量 ~92,800 張若無法延續可能形成短期高點；（5）RSI 進強勢區 68.5、KDJ J 衝至 100 接近超買，技術短線過熱。技術論壇後續：5 座 2nm 廠（新竹 2 + 高雄 3）2026 量產啟動，首年出貨 +45% vs 3nm 2023；2nm 至 2028 CAGR +70%；Arizona +80% YoY、熊本 +130% YoY；A13/A12 2029、N2U 2028、A16 2027 路線完整確認；CoWoS 5.5-reticle 量產（>98% 良率）、2028 推出 14-reticle 整合 10 大晶片+20 HBM。NVIDIA #1 客戶 22% 確認。31 位分析師全數看多，共識目標 NT$2,320（自 NT$2,275 +2.0%），Barclays 目標 $470 對應台股約 NT$2,450（剩 +7.7%）。",
    "note": "資料來源：Bloomberg / MarketScreener（估算），請每日核實更新",
}

NEWS = [
    {
        "tag": "客戶動態", "tag_color": "#2196F3",
        "title": "⚠️ Apple 傳考慮將部分晶片訂單轉至 Intel Foundry 18A 製程；TSMC 5/13 跌 -2.00% 反映客戶集中度風險",
        "body": "市場 5/12-13 傳出 Apple 正在評估將部分晶片訂單（傳為次世代 A 系列或 M 系列部分產品）轉至 Intel Foundry 18A 製程的可能性，為 Apple 多元化代工供應商策略的一環。Apple 目前為 TSMC 第 2 大客戶、佔總營收 17%（僅次於 NVIDIA 22%）。Intel 同日股價大漲 +5.66%、TSMC 台股 5/13 重挫 -2.00% 收 NT$2,210 反映客戶集中度風險。重點分析：（1）即使部分轉單成真、TSMC 在 3nm 以下先進製程技術領先優勢仍未動搖（Intel 18A 良率與量產時程仍有疑問）；（2）Apple 多元化代工為長期趨勢、無立即衝擊 2026 營收能見度；（3）NVIDIA $95B 採購承諾 + AMD/Broadcom Custom XPU 訂單續強為主要結構支撐；（4）此傳聞短期可能造成情緒面壓抑、但 TSMC 72% 全球代工市占護城河結構穩固。Wedbush 與 Bernstein 分析師均表示：Apple 多元化策略已被市場部分計入、不改變 TSMC 長期成長故事。",
        "source": "Reuters / Bloomberg / Taipei Times / 鉅亨網 2026-05-12~13", "impact": "負面",
    },
    {
        "tag": "客戶動態", "tag_color": "#2196F3",
        "title": "📉 台股 2330 5/13 跌破 5MA 收 NT$2,210（-2.00%）：Apple 轉單疑慮 + 獲利了結；市值 NT$57.3 兆",
        "body": "台股 2330 5/13 Wed 受 Apple 傳考慮將部分晶片訂單轉至 Intel Foundry 18A 利空衝擊 + 連續飆漲後獲利了結賣壓共振，收 NT$2,210（-45，-2.00% vs 5/12 NT$2,255），盤中區間 NT$2,205-2,220 緊縮無反彈，量能放大至約 50,200 張（vs 5/12 量縮 39,800 張），市值降至 NT$57.3 兆。籌碼面：5/13 估三大法人合計賣超 -15,500 張——外資轉賣 -12,800 張、投信賣 -2,000 張、自營 -700 張，短線籌碼轉空、外資持股微降至 74.8%。技術面：5/13 跌破 5MA NT$2,250 + 跌破 NT$2,235 前支撐、回測 20MA NT$2,195 區間；RSI 估自 56 降至 48 跌至中軸下、KDJ K(58)/D(60)/J(54) 死叉訊號浮現、MACD +1.5 紅柱續收斂逼近翻黑、潛在死叉風險。明日 5/14 Thu 關鍵：能否守穩 NT$2,200 心理整數 + NT$2,190 20MA 雙重支撐，若失守將測試 NT$2,150 第一支撐。",
        "source": "TWSE / 玩股網 / 永豐金證券 / 鉅亨網 2026-05-13", "impact": "負面",
    },
    {
        "tag": "法規政策", "tag_color": "#FF9800",
        "title": "⭐ TSMC Arizona 子公司董事會 5/12 批准追加 $20B 資本注入、強化美國 Phase 2/3 擴張",
        "body": "TSMC 5/12 公告：TSMC Arizona Corporation 董事會 5/12 決議授權自台積電總公司追加最高 US$20B 資本注入（待主管機關核准），為迄今 TSMC 在美單一最大資本承諾。本筆資金用途：（1）支持 Arizona Fab 21 Phase 2 / Phase 3 大規模廠房建設與設備支出；（2）強化 TSMC Arizona 資產負債表；（3）支持 N3 量產 + N2 製程導入（Phase 2 2028 量產 N2）。意義：本案為 TSMC 美國 $165B 累計投資承諾的關鍵執行步驟、回應川普政府關稅政策與半導體在岸製造要求；同時為 4/22 北美技術論壇宣布 A14（2028 量產）+ A16（H2 2026 量產）+ Arizona 先進封裝廠（2029）路線的資本配套。Barclays 5/4 分析：TSMC 美國產能擴張將有助於：（i）化解地緣政治風險溢價；（ii）提供北美 AI 客戶（NVIDIA / AMD / Broadcom）就近供應；（iii）支撐 N2 製程 ASP 與毛利率。",
        "source": "TSMC SEC 6-K Filing / Stocktitan / TipRanks 2026-05-12", "impact": "正面",
    },
    {
        "tag": "產品發布", "tag_color": "#9C27B0",
        "title": "🚀 TSMC 2026 技術論壇：2nm 產能 2026-2028 CAGR +70%、5 座 2nm 廠今年量產啟動",
        "body": "TSMC 2026 Technology Symposium 矽谷登場期間，TSMC 揭露 2nm 製程量產進度與長期成長路線：（1）2nm（N2）製程 2026 年首批 5 座晶圓廠（新竹 2 座 + 高雄 3 座）開始大量生產；（2）N2 產能 2026-2028 年複合成長率（CAGR）達 +70%，遠超 3nm 量產第一年的擴張速度；（3）N2 首年出貨量較 N3 量產第一年（2023）增加 +45%；（4）A16（H2 2026 量產）+ A14（2028 量產）+ N2U（2028）+ A13/A12（2029）製程路線完整確認；（5）CoWoS 5.5-reticle 量產（>98% 良率）、2028 推出 14-reticle 整合 10 大晶片 + 20 HBM 先進封裝；（6）Siemens EDA 等設計生態系全面整合 N2/A16/A14 PDK。Foundry 龍頭地位結構性強化，呼應 NVIDIA $95B 採購承諾 + AI 加速器 2029 CAGR 54-56% 內部目標。",
        "source": "Semiwiki / TSMC Symposium 2026 / Siemens Press 2026-05-12", "impact": "正面",
    },
    {
        "tag": "產業趨勢", "tag_color": "#4CAF50",
        "title": "🤝 Sony 與 TSMC 5/9 合資 CIS 公司、熊本廠 Physical AI 佈局擴大",
        "body": "Sony 與 TSMC 5/9（週六）宣布簽署非綁定 MOU，合資成立次世代影像感測器（CIS）公司，將在日本熊本縣（既有 JASM 廠基地）開發與生產下世代 CMOS 影像感測器；本合資聚焦「實體 AI（Physical AI）」應用：包括自動駕駛、機器人、邊緣 AI 相機等高階感測需求，將整合 TSMC 先進 stacked-die 製程技術。重點意義：（1）TSMC 與 Sony 既有 JASM 合作（已於 2024 量產 N12/N16）再進一步延伸至 CIS 高階感測器產品；（2）日本半導體政策強化補貼支持；（3）TSMC 熊本廠 Phase 2 已宣布、本次合資將推升熊本基地策略地位；（4）非車載 AI 應用驗證 TSMC 客戶結構持續多元化（NVIDIA 22% + Apple 17% + Broadcom 13% + AMD 8% + Sony 等新合資）。",
        "source": "Bloomberg / Reuters / 日經 / Stocktitan 2026-05-09", "impact": "正面",
    },
    {
        "tag": "產業趨勢", "tag_color": "#4CAF50",
        "title": "📊 全球半導體 2026 預估 +26% YoY 至 $1.32 兆；AI 晶片驅動 TSMC Q1 +40% YoY",
        "body": "Gartner / Deloitte 等市調機構發布 2026 半導體展望：（1）全球半導體收入 2026 預估成長 +26% YoY、若 AI 基建週期高峰可上看 +64% 至 $1.32 兆 / 全年；記憶體支出自 2025 $216B 跳升至 2026 $633B；（2）TSMC Q1 2026 營收年增 +40%（AI 驅動為主、手機與車用衰退），毛利率回升至 50%+ 區間（Q1 實際 66.2% 雙創歷史新高）；管理層維持全年營收 >30% 成長指引；（3）AMD Q1 2026 營收 $10.3B（+38% YoY）、資料中心 +57% YoY 強勁；NVIDIA 採購承諾跨 $95B 揭露反映 GB200/GB300 + Rubin 系列訂單能見度；（4）SOX 指數 YTD 累計 +65% 為現代股市最大半導體 rally 之一；（5）TSMC 揭露 agentic AI 採用浪潮將進一步推升 token 生成需求、為 AI 算力第二波驅動。",
        "source": "Gartner / Deloitte Insights / Motley Fool / Distill Intelligence 2026-05", "impact": "正面",
    },
]

# ═══════════════════════════════════════════════════════════════════
# 輔助函式
# ═══════════════════════════════════════════════════════════════════
def sentiment_color(s):
    if "強烈看多" in s: return "#1B5E20"
    if "中性偏多" in s: return "#388E3C"
    if "看多"    in s: return "#2E7D32"
    if "中性偏空" in s: return "#BF360C"
    if "強烈看空" in s: return "#B71C1C"
    if "看空"    in s: return "#C62828"
    return "#546E7A"   # 中性

def pct_color(s):
    s = s.strip()
    return "#4CAF50" if s.startswith("+") else ("#F44336" if s.startswith("-") else "#888")

def sign_color(v):
    return "#4CAF50" if v >= 0 else "#F44336"

def bar5(values, width=18):
    mx = max(abs(v) for v in values) or 1
    cells = ""
    for v in values:
        pct = abs(v) / mx * 100
        color = "#4CAF50" if v >= 0 else "#F44336"
        label = f"+{v:,}" if v > 0 else f"{v:,}"
        cells += f'<div class="spark-bar"><div style="height:{pct:.0f}%;background:{color};width:100%;border-radius:2px 2px 0 0" title="{label}"></div><div class="spark-val" style="color:{color}">{label if abs(v)>=1000 else label}</div></div>'
    return f'<div class="spark-wrap">{cells}</div>'

def news_html(news_list):
    out = ""
    for n in news_list:
        ic = "#4CAF50" if n["impact"] == "正面" else ("#F44336" if n["impact"] == "負面" else "#FF9800")
        out += f'''<div class="news-card">
          <div class="news-hdr"><span class="tag" style="background:{n["tag_color"]}">{n["tag"]}</span>
            <span class="impact" style="border-color:{ic};color:{ic}">{n["impact"]}</span></div>
          <div class="news-title">{n["title"]}</div>
          <div class="news-body">{n["body"]}</div>
          <div class="news-src">來源：{n["source"]}</div></div>'''
    return out

def earnings_days():
    try:
        d = date.fromisoformat(EARNINGS["next_date"])
        return (d - date.today()).days
    except Exception:
        return "N/A"

# ═══════════════════════════════════════════════════════════════════
# HTML 模板
# ═══════════════════════════════════════════════════════════════════
tw_c  = pct_color(STOCK["tw_change_pct"])
ny_c  = pct_color(STOCK["nyse_change_pct"])
ed    = earnings_days()
f_net = INSTITUTIONAL["foreign"]
t_net = INSTITUTIONAL["trust"]
d_net = INSTITUTIONAL["dealer"]
total_net = f_net + t_net + d_net
sc        = sentiment_color(SUMMARY["sentiment"])
s_bullets = "".join(f"<li>{b}</li>" for b in SUMMARY["bullets"])
s_risk    = f'<div class="s-risk">&#9888; {SUMMARY["risk_alert"]}</div>' if SUMMARY.get("risk_alert") else ""

HTML = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>TSMC 台積電日報 — {TODAY}</title>
<style>
:root{{
  --navy:#0D1B2A; --blue:#1565C0; --mid:#1976D2; --light:#BBDEFB;
  --green:#4CAF50; --red:#F44336; --orange:#FF9800; --purple:#7B1FA2;
  --bg:#F0F4F8; --card:#FFFFFF; --text:#1A1A2E; --muted:#546E7A;
  --border:#E0E8F0;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:var(--bg);color:var(--text);font-size:14px;line-height:1.5}}

/* NAV TOC */
.toc{{background:var(--navy);padding:8px 40px;display:flex;gap:6px;flex-wrap:wrap;position:sticky;top:0;z-index:100}}
.toc a{{color:#90CAF9;font-size:11px;text-decoration:none;padding:3px 8px;border-radius:4px;white-space:nowrap}}
.toc a:hover{{background:rgba(255,255,255,.15);color:#fff}}

/* HEADER */
.hdr{{background:linear-gradient(135deg,var(--navy) 0%,#1a3a6e 100%);color:#fff;padding:24px 40px 18px}}
.hdr h1{{font-size:24px;font-weight:700;letter-spacing:.5px}}
.hdr-sub{{font-size:12px;color:#90CAF9;margin-top:3px}}
.hdr-meta{{display:flex;gap:10px;margin-top:12px;flex-wrap:wrap}}
.hdr-meta span{{background:rgba(255,255,255,.12);border-radius:6px;padding:3px 10px;font-size:11px}}
.data-note{{background:#FFF3E0;border-left:3px solid #FF9800;padding:6px 40px;font-size:11px;color:#E65100}}

/* SECTIONS */
.sec{{padding:16px 40px}}
.sec-title{{font-size:14px;font-weight:700;color:var(--blue);border-left:4px solid var(--blue);
  padding-left:10px;margin-bottom:14px;text-transform:uppercase;letter-spacing:.3px}}
.sub-title{{font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase;margin:12px 0 8px}}

/* STOCK PANEL */
.stock-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
.stock-box{{background:var(--card);border-radius:10px;padding:16px 18px;
  box-shadow:0 1px 5px rgba(0,0,0,.08);border-left:5px solid var(--mid)}}
.stock-exch{{font-size:11px;color:var(--muted);text-transform:uppercase;font-weight:600}}
.stock-row{{display:flex;align-items:baseline;gap:10px;margin:5px 0}}
.price{{font-size:30px;font-weight:700;color:var(--navy)}}
.chg{{font-size:15px;font-weight:600}}
.stock-detail{{display:grid;grid-template-columns:1fr 1fr;gap:3px 10px;margin-top:10px}}
.sd{{font-size:11.5px;color:var(--muted)}}.sd span{{color:var(--navy);font-weight:600}}

/* KPI */
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(148px,1fr));gap:10px}}
.kpi{{background:var(--card);border-radius:9px;padding:12px 14px;
  box-shadow:0 1px 5px rgba(0,0,0,.07);border-top:3px solid var(--blue)}}
.kpi-lbl{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.3px}}
.kpi-val{{font-size:20px;font-weight:700;margin:3px 0 2px;color:var(--navy)}}
.kpi-chg{{font-size:11.5px;font-weight:600}}
.kpi-note{{font-size:10px;color:var(--muted)}}

/* NEWS */
.news-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(310px,1fr));gap:12px}}
.news-card{{background:var(--card);border-radius:9px;padding:14px;
  box-shadow:0 1px 5px rgba(0,0,0,.07);border-bottom:3px solid var(--border)}}
.news-hdr{{display:flex;align-items:center;gap:7px;margin-bottom:7px}}
.tag{{color:#fff;border-radius:4px;padding:2px 7px;font-size:10.5px;font-weight:600}}
.impact{{border:1.5px solid;border-radius:10px;padding:1px 7px;font-size:10.5px;font-weight:600}}
.news-title{{font-size:13px;font-weight:700;margin-bottom:5px;line-height:1.45}}
.news-body{{font-size:12px;color:#444;line-height:1.65;margin-bottom:7px}}
.news-src{{font-size:10.5px;color:var(--muted)}}

/* TABLES */
table{{width:100%;border-collapse:collapse;background:var(--card);
  border-radius:9px;overflow:hidden;box-shadow:0 1px 5px rgba(0,0,0,.07)}}
th{{background:var(--mid);color:#fff;padding:8px 11px;font-size:11.5px;text-align:left;font-weight:600}}
td{{padding:8px 11px;font-size:12px;border-bottom:1px solid #F0F4F8}}
tr:last-child td{{border-bottom:none}}
tr:nth-child(even) td{{background:#F7FBFF}}
.badge{{display:inline-block;border-radius:4px;padding:2px 7px;font-size:10.5px;font-weight:600;color:#fff}}

/* INSTITUTIONAL */
.inst-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}}
.inst-card{{background:var(--card);border-radius:9px;padding:14px;
  box-shadow:0 1px 5px rgba(0,0,0,.07);text-align:center}}
.inst-name{{font-size:11px;color:var(--muted);font-weight:600;text-transform:uppercase}}
.inst-val{{font-size:22px;font-weight:700;margin:5px 0 3px}}
.inst-note{{font-size:10.5px;color:var(--muted)}}

/* SPARK BARS */
.spark-wrap{{display:flex;gap:4px;align-items:flex-end;height:50px;margin-top:8px}}
.spark-bar{{flex:1;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;height:100%}}
.spark-val{{font-size:8.5px;color:var(--muted);margin-top:1px;white-space:nowrap}}

/* TECHNICAL */
.tech-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:10px}}
.tech-card{{background:var(--card);border-radius:9px;padding:12px 14px;
  box-shadow:0 1px 5px rgba(0,0,0,.07);text-align:center}}
.tech-lbl{{font-size:10px;color:var(--muted);text-transform:uppercase}}
.tech-val{{font-size:18px;font-weight:700;margin:4px 0;color:var(--navy)}}
.tech-sig{{font-size:11px;font-weight:600;border-radius:10px;padding:2px 8px;display:inline-block}}
.ma-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-top:10px}}
.ma-box{{background:var(--card);border-radius:8px;padding:8px 10px;text-align:center;
  box-shadow:0 1px 4px rgba(0,0,0,.07)}}
.ma-lbl{{font-size:10px;color:var(--muted)}}
.ma-val{{font-size:14px;font-weight:700;margin-top:2px}}

/* DERIVATIVES */
.deriv-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:10px;margin-bottom:12px}}
.deriv-card{{background:var(--card);border-radius:9px;padding:12px 14px;
  box-shadow:0 1px 5px rgba(0,0,0,.07);text-align:center}}
.deriv-lbl{{font-size:10px;color:var(--muted);text-transform:uppercase}}
.deriv-val{{font-size:18px;font-weight:700;margin:4px 0;color:var(--navy)}}

/* ECOSYSTEM */
.eco-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:10px}}
.eco-card{{background:var(--card);border-radius:9px;padding:12px 14px;
  box-shadow:0 1px 5px rgba(0,0,0,.07)}}
.eco-type{{font-size:10px;color:#fff;border-radius:4px;padding:1px 6px;font-weight:600}}
.eco-name{{font-weight:700;font-size:13px;margin:5px 0 2px}}
.eco-price{{font-size:16px;font-weight:700;color:var(--navy)}}
.eco-note{{font-size:10.5px;color:var(--muted)}}

/* ANALYST */
.analyst-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:10px}}
.analyst-card{{background:var(--card);border-radius:9px;padding:14px;
  box-shadow:0 1px 5px rgba(0,0,0,.07);border-left:4px solid var(--blue)}}
.analyst-firm{{font-size:12px;font-weight:700;color:var(--navy)}}
.analyst-tp{{font-size:18px;font-weight:700;color:var(--blue);margin:5px 0 3px}}
.analyst-note{{font-size:11.5px;color:#555}}
.consensus-bar{{display:flex;border-radius:8px;overflow:hidden;height:22px;margin:10px 0 4px}}

/* EARNINGS COUNTDOWN */
.countdown-box{{background:linear-gradient(135deg,#0D1B2A,#1565C0);color:#fff;
  border-radius:12px;padding:20px 24px;display:grid;grid-template-columns:auto 1fr;gap:20px;align-items:center}}
.countdown-days{{font-size:56px;font-weight:700;line-height:1;color:#FDD835}}
.countdown-label{{font-size:12px;color:#90CAF9;text-transform:uppercase}}
.countdown-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:8px;margin-top:14px}}
.cg-item{{background:rgba(255,255,255,.1);border-radius:8px;padding:8px 12px}}
.cg-lbl{{font-size:10px;color:#90CAF9;text-transform:uppercase}}
.cg-val{{font-size:14px;font-weight:700;margin-top:3px}}

/* FAB UTILIZATION */
.fab-grid{{display:flex;flex-direction:column;gap:8px}}
.fab-row{{background:var(--card);border-radius:9px;padding:11px 14px;
  display:grid;grid-template-columns:90px 80px 140px 1fr;gap:10px;align-items:center;
  box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.fab-node{{font-size:13px;font-weight:700;color:var(--navy)}}
.util-bar{{height:10px;border-radius:5px;background:#E0E0E0;overflow:hidden;margin-top:3px}}
.util-fill{{height:100%;border-radius:5px}}

/* CALENDAR */
.cal-list{{display:flex;flex-direction:column;gap:8px}}
.cal-item{{background:var(--card);border-radius:9px;padding:11px 14px;
  display:grid;grid-template-columns:110px 1fr;gap:12px;align-items:start;
  box-shadow:0 1px 4px rgba(0,0,0,.06)}}
.cal-date{{font-size:13px;font-weight:700;color:#fff;border-radius:6px;
  padding:4px 8px;text-align:center}}
.cal-title{{font-size:13px;font-weight:600;color:var(--navy)}}
.cal-desc{{font-size:11.5px;color:var(--muted);margin-top:2px}}

/* RISK */
.risk-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:10px}}
.risk-card{{background:var(--card);border-radius:9px;padding:12px 14px;
  box-shadow:0 1px 5px rgba(0,0,0,.06);border-left:4px solid}}
.risk-lv{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.3px}}
.risk-name{{font-size:13px;font-weight:700;margin:4px 0}}
.risk-desc{{font-size:12px;color:#555;line-height:1.55}}

/* FOOTER */
footer{{background:var(--navy);color:#607D8B;text-align:center;padding:14px 40px;font-size:10.5px;margin-top:12px}}
footer strong{{color:#90A4AE}}

/* ALERT BOX */
.alert{{background:#FFF8E1;border-left:4px solid #FFC107;border-radius:6px;
  padding:8px 12px;font-size:11.5px;color:#795548;margin-bottom:10px}}

@media(max-width:700px){{
  .sec,.hdr{{padding:12px 14px}}
  .toc{{padding:6px 14px}}
  .stock-grid,.inst-grid{{grid-template-columns:1fr}}
  .fab-row{{grid-template-columns:1fr 1fr}}
  .cal-item{{grid-template-columns:1fr}}
}}
@media print{{
  .toc{{display:none}}
  body{{background:#fff}}
  .kpi,.news-card,.stock-box,.inst-card,.risk-card{{box-shadow:none;border:1px solid #ddd}}
}}

/* ── VALUATION ── */
.val-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px;margin-bottom:12px}}
.val-card{{background:var(--card);border-radius:9px;padding:12px 14px;
  box-shadow:0 1px 5px rgba(0,0,0,.07);text-align:center}}
.val-lbl{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.3px}}
.val-val{{font-size:20px;font-weight:700;color:var(--navy);margin:3px 0 2px}}
.val-sub{{font-size:11px;color:var(--muted)}}
.pct-bar-wrap{{background:#E0E8F0;border-radius:6px;height:10px;overflow:hidden;margin:6px 0 3px}}
.pct-bar-fill{{height:100%;border-radius:6px;background:linear-gradient(90deg,var(--blue),var(--mid))}}

/* ── CASHFLOW ── */
.cf-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px}}
.cf-card{{background:var(--card);border-radius:9px;padding:12px 14px;
  box-shadow:0 1px 5px rgba(0,0,0,.07)}}
.cf-lbl{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.3px}}
.cf-val{{font-size:15px;font-weight:700;color:var(--navy);margin:3px 0 1px}}
.cf-chg{{font-size:11px;font-weight:600}}

/* ── RELATIVE STRENGTH ── */
.rel-row{{display:grid;grid-template-columns:130px repeat(4,1fr);gap:0;margin-bottom:2px;align-items:center}}
.rel-hdr{{font-size:10.5px;font-weight:700;color:var(--muted);text-transform:uppercase;padding:4px 0}}
.rel-lbl{{font-size:12px;font-weight:600;color:var(--navy);padding:4px 0}}
.rel-val{{font-size:12px;font-weight:700;text-align:center;padding:4px 0}}
.adr-box{{background:linear-gradient(135deg,#0D1B2A,#1565C0);color:#fff;border-radius:10px;
  padding:14px 18px;display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-top:12px}}
.adr-lbl{{font-size:10px;color:#90CAF9;text-transform:uppercase}}
.adr-val{{font-size:18px;font-weight:700;margin-top:4px}}

/* ── COWOS ── */
.cowos-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(148px,1fr));gap:10px;margin-bottom:12px}}
.cowos-card{{background:var(--card);border-radius:9px;padding:12px 14px;
  box-shadow:0 1px 5px rgba(0,0,0,.07);text-align:center;border-top:3px solid var(--blue)}}
.cowos-lbl{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.3px}}
.cowos-val{{font-size:20px;font-weight:700;color:var(--navy);margin:3px 0 2px}}
.cowos-sub{{font-size:11px;color:var(--muted)}}
.cowos-cust-bar{{display:flex;height:24px;border-radius:8px;overflow:hidden;margin:10px 0 6px}}

/* ── EPS REVISION ── */
.rev-grid{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px}}
.rev-box{{background:var(--card);border-radius:9px;padding:14px;box-shadow:0 1px 5px rgba(0,0,0,.07)}}
.rev-period{{font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;margin-bottom:8px}}
.rev-consensus{{font-size:28px;font-weight:700;color:var(--navy)}}
.rev-range{{font-size:11px;color:var(--muted);margin-top:3px}}
.rev-trend-up{{display:inline-block;background:#E8F5E9;color:#2E7D32;
  border-radius:10px;padding:2px 9px;font-size:11px;font-weight:700}}
.trigger-list{{display:flex;flex-direction:column;gap:5px}}
.trigger-item{{display:flex;gap:8px;align-items:flex-start;font-size:12px;color:#444;line-height:1.5}}
.trigger-badge{{flex-shrink:0;font-size:10px;font-weight:700;border-radius:4px;
  padding:2px 6px;color:#fff;margin-top:2px}}

/* ── SUMMARY CARD ── */
.summary-card{{background:#fff;border-radius:10px;padding:18px 20px;
  box-shadow:0 2px 10px rgba(0,0,0,.09);border-left:5px solid #888;
  display:flex;gap:18px;align-items:flex-start}}
.summary-left{{flex-shrink:0;display:flex;flex-direction:column;align-items:center;
  gap:6px;min-width:76px;padding-top:2px}}
.s-badge{{color:#fff;font-size:12px;font-weight:700;border-radius:7px;
  padding:5px 10px;text-align:center;line-height:1.4;white-space:nowrap}}
.s-label{{font-size:9px;color:var(--muted);font-weight:600;text-transform:uppercase;
  letter-spacing:.3px}}
.summary-right{{flex:1;min-width:0}}
.s-headline{{font-size:14px;font-weight:700;color:var(--navy);line-height:1.45;
  margin-bottom:10px;padding-bottom:9px;border-bottom:1px solid var(--border)}}
.s-bullets{{list-style:none;display:flex;flex-direction:column;gap:5px;margin-bottom:11px}}
.s-bullets li{{font-size:12.5px;color:#333;line-height:1.6;padding-left:16px;position:relative}}
.s-bullets li::before{{content:"▸";position:absolute;left:0;top:2px;font-size:10px;color:var(--mid)}}
.s-bottom{{background:var(--bg);border-radius:6px;padding:9px 13px;
  font-size:12.5px;color:var(--navy);line-height:1.55;font-style:italic}}
.s-risk{{background:#FFF3E0;border-left:3px solid #FF6F00;border-radius:4px;
  padding:6px 10px;font-size:11.5px;color:#BF360C;font-weight:600;margin-top:8px}}
@media(max-width:600px){{
  .summary-card{{flex-direction:column;gap:10px}}
  .summary-left{{flex-direction:row;min-width:unset}}}}
</style>
</head>
<body>

<!-- TOC -->
<nav class="toc">
  <a href="#summary">&#128203; 今日摘要</a>
  <a href="#price">股價</a>
  <a href="#kpi">財務KPI</a>
  <a href="#inst">法人買賣超</a>
  <a href="#tech">技術分析</a>
  <a href="#deriv">期貨選擇權</a>
  <a href="#news">國際新聞</a>
  <a href="#customer">客戶動態</a>
  <a href="#ecosystem">產業生態系</a>
  <a href="#analyst">法人評級</a>
  <a href="#earnings">財報倒計時</a>
  <a href="#fab">稼動率</a>
  <a href="#risk">風險</a>
  <a href="#calendar">行事曆</a>
  <a href="#valuation">估值位階</a>
  <a href="#cashflow">現金流</a>
  <a href="#relative">相對強弱</a>
  <a href="#cowos">CoWoS</a>
  <a href="#fx">匯率敏感</a>
  <a href="#revision">EPS修正</a>
</nav>

<!-- HEADER -->
<div class="hdr">
  <h1>TSMC 台積電 (2330 / TSM) 每日股市分析日報</h1>
  <div class="hdr-sub">Taiwan Semiconductor Manufacturing Company — Daily Market Intelligence Report</div>
  <div class="hdr-meta">
    <span>報告日期：{TODAY}</span>
    <span>股票代碼：2330.TW / TSM (NYSE)</span>
    <span>下次財報：{EARNINGS["next_date"]} ({ed} 天後)</span>
    <span>⚠ 本報告僅供參考，不構成投資建議</span>
  </div>
</div>
<div class="data-note">⚠ {STOCK["data_note"]}</div>

<!-- ═══ 0. 今日重點摘要 ═══ -->
<div class="sec" id="summary">
  <div class="sec-title">&#128203; 今日重點摘要</div>
  <div class="summary-card" style="border-left-color:{sc}">
    <div class="summary-left">
      <div class="s-badge" style="background:{sc}">{SUMMARY["sentiment"]}</div>
      <div class="s-label">今日信號</div>
    </div>
    <div class="summary-right">
      <div class="s-headline">{SUMMARY["headline"]}</div>
      <ul class="s-bullets">{s_bullets}</ul>
      <div class="s-bottom">&#128161; {SUMMARY["bottom_line"]}</div>
      {s_risk}
    </div>
  </div>
</div>

<!-- ═══ 1. 股價快照 ═══ -->
<div class="sec" id="price">
  <div class="sec-title">📈 今日股價快照</div>
  <div class="stock-grid">
    <div class="stock-box">
      <div class="stock-exch">台灣證交所 TWSE：2330</div>
      <div class="stock-row">
        <div class="price">NT${STOCK["tw_price"]}</div>
        <div class="chg" style="color:{tw_c}">{STOCK["tw_change"]} ({STOCK["tw_change_pct"]})</div>
      </div>
      <div class="stock-detail">
        <div class="sd">開盤 <span>NT${STOCK["tw_open"]}</span></div>
        <div class="sd">最高 <span>NT${STOCK["tw_high"]}</span></div>
        <div class="sd">最低 <span>NT${STOCK["tw_low"]}</span></div>
        <div class="sd">成交量 <span>{STOCK["tw_volume"]} 張</span></div>
        <div class="sd">52W 高 <span>NT${STOCK["52w_high_tw"]}</span></div>
        <div class="sd">52W 低 <span>NT${STOCK["52w_low_tw"]}</span></div>
        <div class="sd">市值 <span>{STOCK["market_cap"]}</span></div>
        <div class="sd">P/E <span>{STOCK["pe_ratio"]}</span></div>
        <div class="sd">P/B <span>{STOCK["pb_ratio"]}</span></div>
        <div class="sd">殖利率 <span>{STOCK["div_yield"]}</span></div>
        <div class="sd">除息日 <span>{STOCK["ex_date"]}</span></div>
        <div class="sd">股利 <span>{STOCK["div_next"]}</span></div>
      </div>
    </div>
    <div class="stock-box">
      <div class="stock-exch">紐約證交所 NYSE：TSM（ADR）</div>
      <div class="stock-row">
        <div class="price">US${STOCK["nyse_price"]}</div>
        <div class="chg" style="color:{ny_c}">{STOCK["nyse_change"]} ({STOCK["nyse_change_pct"]})</div>
      </div>
      <div class="stock-detail">
        <div class="sd">P/E (USD) <span>{STOCK["pe_nyse"]}</span></div>
        <div class="sd">P/B <span>{STOCK["pb_nyse"]}</span></div>
        <div class="sd">52W 高 <span>US${STOCK["52w_high_nyse"]}</span></div>
        <div class="sd">52W 低 <span>US${STOCK["52w_low_nyse"]}</span></div>
        <div class="sd">Beta <span>{STOCK["beta"]}</span></div>
        <div class="sd">下次財報 <span>{STOCK["next_earnings"]}</span></div>
        <div class="sd">共識評級 <span style="color:var(--green)">{CONSENSUS["rating"]}</span></div>
        <div class="sd">均值目標 <span>{CONSENSUS["avg_target_us"]}</span></div>
        <div class="sd">上漲空間 <span style="color:var(--green)">{CONSENSUS["upside"]}</span></div>
        <div class="sd">分析師數 <span>{CONSENSUS["analyst_count"]} 位全買入</span></div>
        <div class="sd">最高目標 <span>{CONSENSUS["high_target"]}</span></div>
        <div class="sd">最低目標 <span>{CONSENSUS["low_target"]}</span></div>
      </div>
    </div>
  </div>
</div>

<!-- ═══ 2. 財務 KPI ═══ -->
<div class="sec" id="kpi">
  <div class="sec-title">💰 財務重點 KPI（Q4 2025 / FY 2025）</div>
  <div class="kpi-grid">
    <div class="kpi"><div class="kpi-lbl">Q4 EPS</div><div class="kpi-val">NT$19.50</div><div class="kpi-chg" style="color:var(--green)">+34.9% YoY</div><div class="kpi-note">超預估 7.23%</div></div>
    <div class="kpi"><div class="kpi-lbl">Q4 毛利率</div><div class="kpi-val">62.3%</div><div class="kpi-chg" style="color:var(--green)">+9.3pp YoY</div><div class="kpi-note">Q3: 57.8%</div></div>
    <div class="kpi"><div class="kpi-lbl">Q4 營業利益率</div><div class="kpi-val">54.0%</div><div class="kpi-chg" style="color:var(--green)">+7.3pp YoY</div><div class="kpi-note">Q3: 50.6%</div></div>
    <div class="kpi"><div class="kpi-lbl">Q4 ROE</div><div class="kpi-val">38.8%</div><div class="kpi-chg" style="color:var(--green)">+9.8pp YoY</div><div class="kpi-note">FY25: 35.4%</div></div>
    <div class="kpi"><div class="kpi-lbl">FY25 EPS</div><div class="kpi-val">NT$66.25</div><div class="kpi-chg" style="color:var(--green)">+46.4% YoY</div><div class="kpi-note">FY24: NT$45.25</div></div>
    <div class="kpi"><div class="kpi-lbl">FY25 營收</div><div class="kpi-val">$122B</div><div class="kpi-chg" style="color:var(--green)">+35.9% YoY</div><div class="kpi-note">TWD 3.8 兆</div></div>
    <div class="kpi"><div class="kpi-lbl">先進製程佔比</div><div class="kpi-val">77%</div><div class="kpi-chg" style="color:var(--green)">≤7nm</div><div class="kpi-note">N3+N5+N7</div></div>
    <div class="kpi"><div class="kpi-lbl">AI 加速器佔比</div><div class="kpi-val">~20%</div><div class="kpi-chg" style="color:var(--green)">60% CAGR 預測</div><div class="kpi-note">至 2029 年</div></div>
    <div class="kpi"><div class="kpi-lbl">Q1 2026 指引</div><div class="kpi-val">$35.2B</div><div class="kpi-chg" style="color:var(--green)">+38% YoY</div><div class="kpi-note">$34.6–35.8B</div></div>
    <div class="kpi"><div class="kpi-lbl">2026 CapEx</div><div class="kpi-val">$52–56B</div><div class="kpi-chg" style="color:var(--orange)">+37–47%</div><div class="kpi-note">70–80% 先進製程</div></div>
    <div class="kpi"><div class="kpi-lbl">2026 股利（預估）</div><div class="kpi-val">≥NT$23</div><div class="kpi-chg" style="color:var(--green)">+27.8%+ YoY</div><div class="kpi-note">FY25: NT$18</div></div>
    <div class="kpi"><div class="kpi-lbl">自由現金流</div><div class="kpi-val">NT$1兆</div><div class="kpi-chg" style="color:var(--green)">+15.2% YoY</div><div class="kpi-note">FY 2025</div></div>
  </div>
</div>

<!-- ═══ 3. 法人買賣超 ═══ -->
<div class="sec" id="inst">
  <div class="sec-title">🏦 三大法人買賣超（{INSTITUTIONAL["date"]}）</div>
  <div class="inst-grid">
    <div class="inst-card">
      <div class="inst-name">外資及陸資</div>
      <div class="inst-val" style="color:{sign_color(f_net)}">{f_net:+,} 張</div>
      <div class="inst-note">{"買超" if f_net>=0 else "賣超"} | 持股：{INSTITUTIONAL["foreign_ownership_pct"]}</div>
      {bar5(INSTITUTIONAL["foreign_5d"])}
    </div>
    <div class="inst-card">
      <div class="inst-name">投信</div>
      <div class="inst-val" style="color:{sign_color(t_net)}">{t_net:+,} 張</div>
      <div class="inst-note">{"買超" if t_net>=0 else "賣超"}</div>
      {bar5(INSTITUTIONAL["trust_5d"])}
    </div>
    <div class="inst-card">
      <div class="inst-name">自營商</div>
      <div class="inst-val" style="color:{sign_color(d_net)}">{d_net:+,} 張</div>
      <div class="inst-note">{"買超" if d_net>=0 else "賣超"}</div>
      {bar5(INSTITUTIONAL["dealer_5d"])}
    </div>
  </div>
  <div style="margin-top:10px;background:var(--card);border-radius:9px;padding:12px 14px;
    box-shadow:0 1px 5px rgba(0,0,0,.07);display:grid;grid-template-columns:repeat(3,1fr);gap:12px">
    <div><div style="font-size:11px;color:var(--muted)">三大法人合計</div>
      <div style="font-size:18px;font-weight:700;color:{sign_color(total_net)}">{total_net:+,} 張</div></div>
    <div><div style="font-size:11px;color:var(--muted)">外資持股比例</div>
      <div style="font-size:18px;font-weight:700">{INSTITUTIONAL["foreign_ownership_pct"]}</div>
      <div style="font-size:10px;color:var(--muted)">({INSTITUTIONAL["foreign_ownership_shares"]})</div></div>
    <div><div style="font-size:11px;color:var(--muted)">外資近 5 日趨勢</div>
      <div style="font-size:13px;font-weight:600;color:var(--red)">持續賣超</div>
      <div style="font-size:10px;color:var(--muted)">從高點獲利了結中</div></div>
  </div>
  <p style="font-size:10.5px;color:var(--muted);margin-top:6px">
    Source: 玩股網 / Goodinfo / TWSE 三大法人日報 | 每日至
    <a href="https://www.twse.com.tw/zh/trading/foreign/t86.html" style="color:var(--blue)">TWSE</a> 更新</p>
</div>

<!-- ═══ 4. 技術分析 ═══ -->
<div class="sec" id="tech">
  <div class="sec-title">📉 技術分析指標</div>
  <div class="alert">{TECHNICAL["note"]}</div>
  <div class="tech-grid">
    <div class="tech-card">
      <div class="tech-lbl">RSI (14)</div>
      <div class="tech-val">{TECHNICAL["rsi_14"]}</div>
      <span class="tech-sig" style="background:#FF9800;color:#fff">{TECHNICAL["rsi_signal"]}</span>
    </div>
    <div class="tech-card">
      <div class="tech-lbl">MACD</div>
      <div class="tech-val" style="color:{TECHNICAL['macd_color']}">{TECHNICAL["macd"]}</div>
      <span class="tech-sig" style="background:{TECHNICAL['macd_color']};color:#fff">{TECHNICAL["macd_signal"]}</span>
    </div>
    <div class="tech-card">
      <div class="tech-lbl">KDJ (9,3,3)</div>
      <div class="tech-val">K:{TECHNICAL["kdj_k"]:.0f}</div>
      <span class="tech-sig" style="background:#FF9800;color:#fff">{TECHNICAL["kdj_signal"]}</span>
      <div style="font-size:10px;color:var(--muted);margin-top:2px">D:{TECHNICAL["kdj_d"]:.0f} / J:{TECHNICAL["kdj_j"]:.0f}</div>
    </div>
    <div class="tech-card">
      <div class="tech-lbl">布林帶（20,2）</div>
      <div class="tech-val" style="font-size:14px">NT${STOCK["tw_price"]}</div>
      <div style="font-size:10.5px;color:var(--muted)">上：{TECHNICAL["bb_upper"]}　中：{TECHNICAL["bb_mid"]}</div>
      <div style="font-size:10.5px;color:var(--muted)">下：{TECHNICAL["bb_lower"]}</div>
      <span class="tech-sig" style="background:#FF9800;color:#fff;margin-top:4px">布林收窄整理</span>
    </div>
    <div class="tech-card">
      <div class="tech-lbl">成交量</div>
      <div class="tech-val">{STOCK["tw_volume"]}</div>
      <div style="font-size:10.5px;color:var(--muted)">5日均量：{TECHNICAL["vol_ma5"]} 張</div>
      <span class="tech-sig" style="background:#607D8B;color:#fff;margin-top:4px">量縮整理</span>
    </div>
    <div class="tech-card">
      <div class="tech-lbl">整體趨勢</div>
      <div class="tech-val" style="font-size:15px">{TECHNICAL["trend"]}</div>
      <div style="font-size:10.5px;color:var(--muted)">支撐：{TECHNICAL["support"]}</div>
      <div style="font-size:10.5px;color:var(--muted)">壓力：{TECHNICAL["resist"]}</div>
    </div>
  </div>
  <div class="sub-title">均線系統</div>
  <div class="ma-grid">
    {"".join(f'<div class="ma-box"><div class="ma-lbl">{lbl}</div><div class="ma-val" style="color:{("#F44336" if float(STOCK["tw_price"].replace(",","")) < float(v.replace(",","")) else "#4CAF50")}">{v}</div></div>' for lbl, v in [("MA5", TECHNICAL["ma5"]), ("MA20", TECHNICAL["ma20"]), ("MA60", TECHNICAL["ma60"]), ("MA120", TECHNICAL["ma120"])])}
  </div>
  <p style="font-size:10.5px;color:var(--muted);margin-top:8px">現價 NT${STOCK["tw_price"]} 已跌破 MA5({TECHNICAL["ma5"]}) 與 MA20({TECHNICAL["ma20"]})，短線偏空；但仍在 MA60({TECHNICAL["ma60"]}) 之上，中期趨勢尚未破壞。</p>
</div>

<!-- ═══ 5. 期貨選擇權 ═══ -->
<div class="sec" id="deriv">
  <div class="sec-title">📊 台積電期貨 & 選擇權動態</div>
  <div class="alert">{DERIVATIVES["note"]}</div>
  <div class="deriv-grid">
    <div class="deriv-card"><div class="deriv-lbl">期貨近月收盤</div><div class="deriv-val">NT${DERIVATIVES["futures_price"]}</div><div style="font-size:10.5px;color:var(--muted)">基差 {DERIVATIVES["futures_basis"]} 點（正價差）</div></div>
    <div class="deriv-card"><div class="deriv-lbl">未平倉量（口）</div><div class="deriv-val">{DERIVATIVES["futures_oi"]}</div><div style="font-size:10.5px;color:var(--muted)">{DERIVATIVES["date"]}</div></div>
    <div class="deriv-card"><div class="deriv-lbl">Put/Call Ratio</div><div class="deriv-val" style="color:var(--red)">{DERIVATIVES["pcr"]}</div><div style="font-size:10.5px;color:var(--red)">偏空（>1 = 空力較強）</div></div>
    <div class="deriv-card"><div class="deriv-lbl">Call 未平倉</div><div class="deriv-val" style="color:var(--green)">{DERIVATIVES["call_oi"]}</div><div style="font-size:10.5px;color:var(--muted)">全行使價合計</div></div>
    <div class="deriv-card"><div class="deriv-lbl">Put 未平倉</div><div class="deriv-val" style="color:var(--red)">{DERIVATIVES["put_oi"]}</div><div style="font-size:10.5px;color:var(--muted)">全行使價合計</div></div>
    <div class="deriv-card"><div class="deriv-lbl">最大痛苦點</div><div class="deriv-val">{DERIVATIVES["max_pain"]}</div><div style="font-size:10.5px;color:var(--muted)">市場多空分界</div></div>
  </div>
  <div class="sub-title">主要行使價未平倉分布</div>
  <table>
    <thead><tr><th>行使價</th><th>Call OI / Put OI</th><th>市場含意</th></tr></thead>
    <tbody>
      {"".join(f'<tr><td><strong>{s[0]}</strong></td><td>{s[1]}</td><td><span class="badge" style="background:{"#4CAF50" if "支撐" in s[2] else "#F44336"}">{s[2]}</span></td></tr>' for s in DERIVATIVES["key_strikes"])}
    </tbody>
  </table>
  <p style="font-size:10.5px;color:var(--muted);margin-top:6px">
    Source: 台灣期交所 <a href="https://www.taifex.com.tw" style="color:var(--blue)">taifex.com.tw</a></p>
</div>

<!-- ═══ 6. 國際重要新聞 ═══ -->
<div class="sec" id="news">
  <div class="sec-title">📰 國際重要新聞分析</div>
  <div class="news-grid">{news_html(NEWS)}</div>
</div>

<!-- ═══ 7. 主要客戶動態 ═══ -->
<div class="sec" id="customer">
  <div class="sec-title">🏢 主要客戶 2026 預估營收佔比</div>
  <div style="display:flex;height:30px;border-radius:8px;overflow:hidden;margin-bottom:12px">
    <div style="width:21%;background:#76b900" title="NVIDIA 21%"></div>
    <div style="width:17%;background:#999999" title="Apple 17%"></div>
    <div style="width:13%;background:#CC0000" title="Broadcom 13%"></div>
    <div style="width:9%;background:#00AAFF"  title="MediaTek 9%"></div>
    <div style="width:8%;background:#ED1C24"  title="AMD 8%"></div>
    <div style="width:7%;background:#3253DC"  title="Qualcomm 7%"></div>
    <div style="width:7%;background:#0071C5"  title="Intel 7%"></div>
    <div style="width:18%;background:#AAAAAA" title="其他 18%"></div>
  </div>
  <table>
    <thead><tr><th>客戶</th><th>預估佔比</th><th>預估營收 (USD)</th><th>主要晶片</th><th>趨勢</th></tr></thead>
    <tbody>
      <tr><td><strong style="color:#76b900">NVIDIA</strong></td><td><strong>~21%</strong></td><td>~$33B</td><td>H100/H200、Vera Rubin (N3)</td><td><span class="badge" style="background:#4CAF50">快速成長</span></td></tr>
      <tr><td><strong style="color:#999">Apple</strong></td><td>~17%</td><td>~$27B</td><td>A18 Pro / M5 / C1 Modem</td><td><span class="badge" style="background:#2196F3">穩定</span></td></tr>
      <tr><td><strong style="color:#CC0000">Broadcom</strong></td><td>~13%</td><td>—</td><td>ASIC/AI 加速器、網路晶片</td><td><span class="badge" style="background:#4CAF50">快速成長</span></td></tr>
      <tr><td><strong style="color:#00AAFF">MediaTek</strong></td><td>~9%</td><td>—</td><td>Dimensity 9400 等旗艦 SoC</td><td><span class="badge" style="background:#2196F3">穩定</span></td></tr>
      <tr><td><strong style="color:#ED1C24">AMD</strong></td><td>~8%</td><td>—</td><td>MI300X AI GPU / Zen5 CPU</td><td><span class="badge" style="background:#4CAF50">成長</span></td></tr>
      <tr><td><strong style="color:#3253DC">Qualcomm</strong></td><td>~7%</td><td>—</td><td>Snapdragon Elite 系列</td><td><span class="badge" style="background:#FF9800">持平</span></td></tr>
      <tr><td><strong style="color:#0071C5">Intel</strong></td><td>~7%</td><td>—</td><td>Arrow Lake / Panther Lake</td><td><span class="badge" style="background:#FF9800">持平</span></td></tr>
    </tbody>
  </table>
  <p style="font-size:10.5px;color:var(--muted);margin-top:6px">Source: Morgan Stanley / Creative Strategies / CNBC 2026-01</p>
</div>

<!-- ═══ 8. 產業生態系股價連動 ═══ -->
<div class="sec" id="ecosystem">
  <div class="sec-title">⛓️ 供應鏈 & 客戶股價連動（估算）</div>
  <div class="alert">⚠ 股價數據為估算值，每日請至看盤平台更新。色塊：<span class="badge" style="background:#2196F3">設備</span> <span class="badge" style="background:#9C27B0">記憶體</span> <span class="badge" style="background:#009688">封裝</span> <span class="badge" style="background:#1565C0">客戶</span> <span class="badge" style="background:#607D8B">競爭</span></div>
  <table>
    <thead><tr><th>類別</th><th>代碼</th><th>名稱</th><th>股價（估）</th><th>漲跌幅（估）</th><th>與 TSMC 關係</th></tr></thead>
    <tbody>
      {"".join(f'''<tr>
        <td><span class="badge" style="background:{"#2196F3" if r[0]=="設備" else "#9C27B0" if r[0]=="記憶" else "#009688" if r[0]=="封裝" else "#1565C0" if r[0]=="客戶" else "#607D8B"}">{r[0]}</span></td>
        <td><strong>{r[1]}</strong></td>
        <td>{r[2]}</td>
        <td>{r[3]}</td>
        <td style="color:{pct_color(r[4])};font-weight:600">{r[4]}</td>
        <td style="font-size:11px;color:#444">{r[5]}</td>
      </tr>''' for r in ECOSYSTEM_STOCKS)}
    </tbody>
  </table>
</div>

<!-- ═══ 9. 法人評級彙整 ═══ -->
<div class="sec" id="analyst">
  <div class="sec-title">🎯 法人評級彙整</div>
  <div style="background:var(--card);border-radius:10px;padding:16px;margin-bottom:14px;
    box-shadow:0 1px 5px rgba(0,0,0,.08)">
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px">
      <div style="text-align:center">
        <div style="font-size:11px;color:var(--muted)">共識評級</div>
        <div style="font-size:22px;font-weight:700;color:var(--green)">{CONSENSUS["rating"]}</div>
        <div style="font-size:11px;color:var(--muted)">{CONSENSUS["analyst_count"]} 位分析師</div>
      </div>
      <div style="text-align:center">
        <div style="font-size:11px;color:var(--muted)">平均目標價（TWD）</div>
        <div style="font-size:22px;font-weight:700;color:var(--blue)">{CONSENSUS["avg_target_tw"]}</div>
        <div style="font-size:11px;color:var(--green)">上漲空間 {CONSENSUS["upside"]}</div>
      </div>
      <div style="text-align:center">
        <div style="font-size:11px;color:var(--muted)">評級分布</div>
        <div class="consensus-bar">
          <div style="width:{int(CONSENSUS['buy']/CONSENSUS['analyst_count']*100)}%;background:var(--green)"></div>
          <div style="width:{int(CONSENSUS['hold']/CONSENSUS['analyst_count']*100)}%;background:#FF9800"></div>
        </div>
        <div style="font-size:10.5px;color:var(--muted)">買入 {CONSENSUS["buy"]} / 持有 {CONSENSUS["hold"]} / 賣出 {CONSENSUS["sell"]}</div>
      </div>
    </div>
  </div>
  <div class="analyst-grid">
    {"".join(f'''<div class="analyst-card">
      <div class="analyst-firm">{r[0]}</div>
      <div class="analyst-tp">{r[2]}</div>
      <div style="font-size:11px;margin-bottom:5px">
        <span class="badge" style="background:var(--green)">{r[1]}</span>
        <span style="color:var(--muted);font-size:10px;margin-left:6px">{r[3]}</span>
      </div>
      <div class="analyst-note">{r[4]}</div>
    </div>''' for r in ANALYST_RATINGS)}
  </div>
</div>

<!-- ═══ 10. 財報倒計時 ═══ -->
<div class="sec" id="earnings">
  <div class="sec-title">⏳ 財報倒計時 — {EARNINGS["quarter"]}</div>
  <div class="countdown-box">
    <div>
      <div class="countdown-days">{ed}</div>
      <div class="countdown-label">天後發布</div>
    </div>
    <div>
      <div style="font-size:16px;font-weight:700;margin-bottom:10px">{EARNINGS["quarter"]} 財報預計：{EARNINGS["next_date"]}</div>
      <div class="countdown-grid">
        <div class="cg-item"><div class="cg-lbl">營收指引（中位）</div><div class="cg-val">{EARNINGS["rev_guide_mid"]}</div></div>
        <div class="cg-item"><div class="cg-lbl">指引範圍</div><div class="cg-val">{EARNINGS["rev_guide_lo"]} – {EARNINGS["rev_guide_hi"]}</div></div>
        <div class="cg-item"><div class="cg-lbl">毛利率指引</div><div class="cg-val">{EARNINGS["gm_guide"]}</div></div>
        <div class="cg-item"><div class="cg-lbl">營業利益率指引</div><div class="cg-val">{EARNINGS["op_guide"]}</div></div>
        <div class="cg-item"><div class="cg-lbl">共識 EPS 預估</div><div class="cg-val">{EARNINGS["eps_consensus"]}</div></div>
        <div class="cg-item"><div class="cg-lbl">上季超預期幅度</div><div class="cg-val" style="color:#FDD835">{EARNINGS["last_beat"]}</div></div>
        <div class="cg-item"><div class="cg-lbl">近 8 季超預期率</div><div class="cg-val" style="color:#FDD835">{EARNINGS["beat_rate"]}</div></div>
        <div class="cg-item"><div class="cg-lbl">YoY 營收成長</div><div class="cg-val" style="color:#FDD835">~38%</div></div>
      </div>
    </div>
  </div>
</div>

<!-- ═══ 11. 晶圓廠稼動率 ═══ -->
<div class="sec" id="fab">
  <div class="sec-title">🏭 晶圓廠製程稼動率（業界分析師估算）</div>
  <div class="alert">⚠ 稼動率為業界分析師估算值，TSMC 不公開分製程稼動率數據</div>
  <div class="fab-grid">
    {"".join(f'''<div class="fab-row">
      <div class="fab-node">{r[0]}</div>
      <div><span class="badge" style="background:#607D8B">{r[1]}</span></div>
      <div>
        <div style="font-size:13px;font-weight:700;color:{r[4]}">{r[2]}</div>
        <div class="util-bar"><div class="util-fill" style="width:{r[2].split("–")[0].replace("~","").replace("%","").strip() if "%" in r[2] else "50"}%;background:{r[4]}"></div></div>
      </div>
      <div style="font-size:11.5px;color:#555">{r[3]}</div>
    </div>''' for r in FAB_UTILIZATION)}
  </div>
  <p style="font-size:10.5px;color:var(--muted);margin-top:8px">Source: TrendForce、Digitimes、業界分析師 2026 Q1 估算</p>
</div>

<!-- ═══ RISK + GEO ═══ -->
<div class="sec" id="risk">
  <div class="sec-title">⚠️ 風險評估矩陣 × 地緣政治</div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
    <div style="background:var(--card);border-radius:10px;padding:14px;
      box-shadow:0 1px 5px rgba(0,0,0,.07);border-left:4px solid var(--red)">
      <div style="font-weight:700;color:#B71C1C;margin-bottom:8px">🔴 地緣政治 & 法規風險</div>
      <ul style="font-size:12px;color:#444;line-height:1.9;padding-left:16px">
        <li>美中 AI 晶片出口管制：BIS 2026-01-15 改逐案審查，法遵成本上升</li>
        <li>NVIDIA H200/H20 中國版停產，影響 TSMC 訂單結構</li>
        <li>中國反制：鎵、鍺出口限制，半導體材料供給受壓</li>
        <li>關稅 25%：拉高客戶成本，NVIDIA GPU 已漲價 5–15%</li>
        <li>台海局勢不確定性，「矽盾」效果存疑</li>
      </ul>
    </div>
    <div style="background:var(--card);border-radius:10px;padding:14px;
      box-shadow:0 1px 5px rgba(0,0,0,.07);border-left:4px solid var(--green)">
      <div style="font-weight:700;color:#1B5E20;margin-bottom:8px">🟢 結構性長期利多</div>
      <ul style="font-size:12px;color:#444;line-height:1.9;padding-left:16px">
        <li>全球半導體市場 2026 估達 $975B，成長 26%</li>
        <li>AI 晶片佔總收入約 50%，2036 市場 $2 兆</li>
        <li>台日印三國框架：日本資本 × 台灣技術 × 印度人才</li>
        <li>美台協議 $2,500 億投資，Arizona/Japan 廠分散地緣風險</li>
        <li>Samsung 良率問題未解、Intel Foundry 轉型中 → TSMC 優勢穩固</li>
      </ul>
    </div>
  </div>
  <div class="risk-grid">
    {"".join(f'''<div class="risk-card" style="border-color:{r["color"]}">
      <div class="risk-lv" style="color:{r["color"]}">{"●●●" if r["level"]=="高" else "●●○" if r["level"]=="中" else "●○○"} {r["level"]}風險</div>
      <div class="risk-name">{r["name"]}</div>
      <div class="risk-desc">{r["desc"]}</div>
    </div>''' for r in [
        {"level":"高","color":"#F44336","name":"出口管制","desc":"AI 晶片對中出口政策持續演變，H200/H20 停產顯示政策不確定性仍高"},
        {"level":"高","color":"#F44336","name":"地緣政治","desc":"台海緊張局勢影響生產連續性，「矽盾」效果受關稅政策弱化"},
        {"level":"中","color":"#FF9800","name":"關稅衝擊","desc":"25% 晶片進口關稅拉高客戶成本，消費性需求可能降溫"},
        {"level":"中","color":"#FF9800","name":"毛利率稀釋","desc":"海外廠（AZ/JP）製程複製成本高，預估壓縮 1–2pp 毛利率"},
        {"level":"中","color":"#FF9800","name":"AI 需求集中","desc":"HPC/AI 佔比 58%，若 AI 景氣反轉衝擊明顯；消費性電子仍低迷"},
        {"level":"中","color":"#FF9800","name":"匯率風險","desc":"台幣升值壓縮美元計算的海外收益，每 1% 影響毛利率 0.3–0.4pp"},
        {"level":"低","color":"#4CAF50","name":"技術良率","desc":"N2 HVM 良率爬升符合預期，N2P/A16 按期程進行"},
        {"level":"低","color":"#4CAF50","name":"競爭威脅","desc":"Samsung 良率不穩、Intel 轉型需 2–3 年，TSMC 技術護城河穩固"},
    ])}
  </div>
</div>

<!-- ═══ 12. IR 行事曆 ═══ -->
<div class="sec" id="calendar">
  <div class="sec-title">📅 法說會 / IR 行事曆</div>
  <div class="cal-list">
    {"".join(f'''<div class="cal-item">
      <div class="cal-date" style="background:{r[3]}">{r[0]}</div>
      <div><div class="cal-title">{r[1]}</div><div class="cal-desc">{r[2]}</div></div>
    </div>''' for r in IR_CALENDAR)}
  </div>
  <p style="font-size:10.5px;color:var(--muted);margin-top:8px">
    Source: <a href="https://investor.tsmc.com/english/financial-calendar" style="color:var(--blue)">TSMC Investor Relations — Financial Calendar</a></p>
</div>

<!-- ═══ 14. 估值歷史位階 ═══ -->
<div class="sec" id="valuation">
  <div class="sec-title">📐 估值歷史位階（Valuation Percentile）</div>
  <div class="alert">⚠ {VALUATION["note"]}</div>
  <div class="val-grid">
    <div class="val-card">
      <div class="val-lbl">P/E 目前</div>
      <div class="val-val">{VALUATION["pe_current"]}x</div>
      <div class="val-sub">5Y 均 {VALUATION["pe_5y_avg"]}x</div>
      <div class="pct-bar-wrap"><div class="pct-bar-fill" style="width:{VALUATION["pe_5y_pct"]}%"></div></div>
      <div class="val-sub">5Y 百分位 <strong>{VALUATION["pe_5y_pct"]}%</strong>（高 {VALUATION["pe_5y_high"]}x / 低 {VALUATION["pe_5y_low"]}x）</div>
    </div>
    <div class="val-card">
      <div class="val-lbl">P/E 10年位階</div>
      <div class="val-val">{VALUATION["pe_10y_pct"]}%</div>
      <div class="val-sub">10Y 均 {VALUATION["pe_10y_avg"]}x</div>
      <div class="pct-bar-wrap"><div class="pct-bar-fill" style="width:{VALUATION["pe_10y_pct"]}%"></div></div>
      <div class="val-sub">10Y 歷史百分位</div>
    </div>
    <div class="val-card">
      <div class="val-lbl">P/B 目前</div>
      <div class="val-val">{VALUATION["pb_current"]}x</div>
      <div class="val-sub">5Y 均 {VALUATION["pb_5y_avg"]}x</div>
      <div class="pct-bar-wrap"><div class="pct-bar-fill" style="width:{VALUATION["pb_5y_pct"]}%"></div></div>
      <div class="val-sub">5Y 百分位 <strong>{VALUATION["pb_5y_pct"]}%</strong></div>
    </div>
    <div class="val-card">
      <div class="val-lbl">EV/EBITDA（估）</div>
      <div class="val-val">{VALUATION["ev_ebitda"]}x</div>
      <div class="val-sub">半導體業均約 12–16x</div>
    </div>
    <div class="val-card">
      <div class="val-lbl">PEG Ratio</div>
      <div class="val-val" style="color:var(--green)">{VALUATION["peg"]}</div>
      <div class="val-sub">&lt;1 = 相對成長便宜</div>
    </div>
  </div>
  <div style="background:var(--card);border-radius:9px;padding:12px 16px;
    box-shadow:0 1px 5px rgba(0,0,0,.07);font-size:12.5px;color:#444;line-height:1.75">
    <strong>估值解讀：</strong>目前 P/E {VALUATION["pe_current"]}x 處於 5 年 <strong>{VALUATION["pe_5y_pct"]}%</strong> 分位 / 10 年 <strong>{VALUATION["pe_10y_pct"]}%</strong> 分位高位，
    反映市場對 AI 超級週期的溢價預期。PEG {VALUATION["peg"]} 顯示相對成長速度仍屬合理；EV/EBITDA {VALUATION["ev_ebitda"]}x 高於業均，
    但考量 TSMC 壟斷性市占與訂單能見度，溢價具基本面支撐。若 AI 資本支出鬆動，P/E 均值回歸至 {VALUATION["pe_5y_avg"]}x
    意味台股下修至約 NT$1,200，為主要下行風險情境。
  </div>
</div>

<!-- ═══ 15. 現金流品質 ═══ -->
<div class="sec" id="cashflow">
  <div class="sec-title">💵 現金流品質分析（Quality of Earnings）</div>
  <div class="alert">⚠ {CASHFLOW["note"]}</div>
  <div class="cf-grid">
    <div class="cf-card">
      <div class="cf-lbl">FY25 合併營收</div>
      <div class="cf-val">{CASHFLOW["revenue_fy25"]}</div>
    </div>
    <div class="cf-card">
      <div class="cf-lbl">FY25 毛利率（估）</div>
      <div class="cf-val" style="color:var(--green)">{CASHFLOW["gm_fy25"]}</div>
      <div class="cf-chg" style="color:var(--muted)">Q4 2025：62.3%</div>
    </div>
    <div class="cf-card">
      <div class="cf-lbl">FY25 營業利益率（估）</div>
      <div class="cf-val" style="color:var(--green)">{CASHFLOW["op_margin_fy25"]}</div>
      <div class="cf-chg" style="color:var(--muted)">Q4 2025：54.0%</div>
    </div>
    <div class="cf-card">
      <div class="cf-lbl">FY25 CapEx（估）</div>
      <div class="cf-val">{CASHFLOW["capex_fy25"]}</div>
      <div class="cf-chg" style="color:var(--muted)">佔收入 {CASHFLOW["capex_pct_rev"]}</div>
    </div>
    <div class="cf-card">
      <div class="cf-lbl">FY25 自由現金流</div>
      <div class="cf-val" style="color:var(--green)">{CASHFLOW["fcf_fy25"]}</div>
      <div class="cf-chg" style="color:var(--muted)">FCF 殖利率 {CASHFLOW["fcf_yield"]}</div>
    </div>
    <div class="cf-card">
      <div class="cf-lbl">存貨天數 DOI（估）</div>
      <div class="cf-val">{CASHFLOW["doi_current"]} 天</div>
      <div class="cf-chg" style="color:var(--green)">YoY {CASHFLOW["doi_yoy"]} 天（去化）</div>
    </div>
    <div class="cf-card">
      <div class="cf-lbl">應收帳款天數 AR（估）</div>
      <div class="cf-val">{CASHFLOW["ar_days"]} 天</div>
      <div class="cf-chg" style="color:var(--green)">YoY {CASHFLOW["ar_days_yoy"]} 天（改善）</div>
    </div>
    <div class="cf-card">
      <div class="cf-lbl">庫存趨勢</div>
      <div class="cf-val" style="font-size:13px;color:var(--green)">{CASHFLOW["inventory_trend"]}</div>
      <div class="cf-chg" style="color:var(--muted)">下游需求復甦訊號</div>
    </div>
  </div>
</div>

<!-- ═══ 16. 相對強弱 / ADR 折溢價 ═══ -->
<div class="sec" id="relative">
  <div class="sec-title">📊 相對強弱 &amp; ADR 折溢價（Relative Strength）</div>
  <div class="alert">⚠ {RELATIVE["note"]}</div>
  <div style="background:var(--card);border-radius:9px;padding:14px 16px;
    box-shadow:0 1px 5px rgba(0,0,0,.07);overflow-x:auto;margin-bottom:12px">
    <div class="rel-row" style="border-bottom:2px solid var(--border);padding-bottom:6px;margin-bottom:6px">
      <div class="rel-hdr">標的</div>
      <div class="rel-hdr" style="text-align:center">今日</div>
      <div class="rel-hdr" style="text-align:center">5 日</div>
      <div class="rel-hdr" style="text-align:center">1 月</div>
      <div class="rel-hdr" style="text-align:center">YTD</div>
    </div>
    <div class="rel-row" style="background:#F7FBFF;border-radius:6px">
      <div class="rel-lbl">TSMC 台股 2330</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['tsmc_tw_1d'])}">{RELATIVE['tsmc_tw_1d']:+.2f}%</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['tsmc_tw_5d'])}">{RELATIVE['tsmc_tw_5d']:+.1f}%</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['tsmc_tw_1m'])}">{RELATIVE['tsmc_tw_1m']:+.1f}%</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['tsmc_tw_ytd'])}">{RELATIVE['tsmc_tw_ytd']:+.1f}%</div>
    </div>
    <div class="rel-row">
      <div class="rel-lbl">TAIEX 加權指數</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['taiex_1d'])}">{RELATIVE['taiex_1d']:+.2f}%</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['taiex_5d'])}">{RELATIVE['taiex_5d']:+.1f}%</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['taiex_1m'])}">{RELATIVE['taiex_1m']:+.1f}%</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['taiex_ytd'])}">{RELATIVE['taiex_ytd']:+.1f}%</div>
    </div>
    <div class="rel-row" style="background:#F7FBFF;border-radius:6px">
      <div class="rel-lbl">SOX 半導體指數</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['sox_1d'])}">{RELATIVE['sox_1d']:+.2f}%</div>
      <div class="rel-val">—</div>
      <div class="rel-val">—</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['sox_ytd'])}">{RELATIVE['sox_ytd']:+.1f}%</div>
    </div>
    <div class="rel-row">
      <div class="rel-lbl">TSM NYSE ADR</div>
      <div class="rel-val" style="color:{sign_color(RELATIVE['tsmc_nyse_1d'])}">{RELATIVE['tsmc_nyse_1d']:+.2f}%</div>
      <div class="rel-val">—</div>
      <div class="rel-val">—</div>
      <div class="rel-val">—</div>
    </div>
  </div>
  <div class="adr-box">
    <div>
      <div class="adr-lbl">ADR 市價（NYSE）</div>
      <div class="adr-val">US${RELATIVE["adr_price"]}</div>
      <div style="font-size:10px;color:#90CAF9;margin-top:2px">比例 {RELATIVE["adr_ratio"]}（1 ADR = 5 股）</div>
    </div>
    <div>
      <div class="adr-lbl">ADR 理論平價（估）</div>
      <div class="adr-val">US${RELATIVE["adr_parity"]:.2f}</div>
      <div style="font-size:10px;color:#90CAF9;margin-top:2px">5 × NT${STOCK["tw_price"]} ÷ {RELATIVE["usd_ntd"]}</div>
    </div>
    <div>
      <div class="adr-lbl">ADR 溢價</div>
      <div class="adr-val" style="color:#FDD835">{RELATIVE["adr_premium_pct"]}%</div>
      <div style="font-size:10px;color:#90CAF9;margin-top:2px">Beta(60d): {RELATIVE["beta_60d"]}</div>
    </div>
  </div>
</div>

<!-- ═══ 17. CoWoS 先進封裝產能 ═══ -->
<div class="sec" id="cowos">
  <div class="sec-title">🔧 CoWoS 先進封裝產能追蹤</div>
  <div class="alert">⚠ {COWOS["note"]}</div>
  <div class="cowos-grid">
    <div class="cowos-card">
      <div class="cowos-lbl">CoWoS-S 月產能</div>
      <div class="cowos-val">{COWOS["cowos_s_cap"]}</div>
      <div class="cowos-sub">wspm（主力 AI GPU）</div>
    </div>
    <div class="cowos-card">
      <div class="cowos-lbl">CoWoS-L 月產能</div>
      <div class="cowos-val">{COWOS["cowos_l_cap"]}</div>
      <div class="cowos-sub">wspm（大尺寸 HPC）</div>
    </div>
    <div class="cowos-card">
      <div class="cowos-lbl">CoWoS-R 月產能</div>
      <div class="cowos-val">{COWOS["cowos_r_cap"]}</div>
      <div class="cowos-sub">wspm（RDL 次世代）</div>
    </div>
    <div class="cowos-card">
      <div class="cowos-lbl">2025 年合計</div>
      <div class="cowos-val">{COWOS["total_2025"]}</div>
      <div class="cowos-sub">wspm</div>
    </div>
    <div class="cowos-card" style="border-top-color:var(--green)">
      <div class="cowos-lbl">2026 年目標</div>
      <div class="cowos-val" style="color:var(--green)">{COWOS["target_2026"]}</div>
      <div class="cowos-sub">wspm / YoY {COWOS["growth_yoy"]}</div>
    </div>
    <div class="cowos-card">
      <div class="cowos-lbl">稼動率</div>
      <div class="cowos-val" style="color:var(--green)">{COWOS["utilization"]}</div>
      <div class="cowos-sub">訂單積壓 {COWOS["backlog_qtrs"]} 季</div>
    </div>
  </div>
  <div class="sub-title">CoWoS 客戶分配佔比（估）</div>
  <div class="cowos-cust-bar">
    {"".join(f'<div style="flex:{c[1].strip("~%")};background:{c[3]}" title="{c[0]} {c[1]} — {c[2]}"></div>' for c in COWOS["customers"])}
  </div>
  <table>
    <thead><tr><th>客戶</th><th>CoWoS 佔比（估）</th><th>主要產品</th></tr></thead>
    <tbody>
      {"".join(f'<tr><td><strong style="color:{c[3]}">{c[0]}</strong></td><td><strong>{c[1]}</strong></td><td>{c[2]}</td></tr>' for c in COWOS["customers"])}
    </tbody>
  </table>
  <div style="margin-top:10px;background:#FFF3E0;border-left:3px solid var(--orange);
    border-radius:6px;padding:8px 12px;font-size:11.5px;color:#795548;line-height:1.65">
    <strong>瓶頸說明：</strong>{COWOS["bottleneck"]}<br>
    <strong>ASP 說明：</strong>{COWOS["asp_note"]}
  </div>
</div>

<!-- ═══ 18. 匯率敏感度 ═══ -->
<div class="sec" id="fx">
  <div class="sec-title">💱 匯率敏感度分析（FX Sensitivity）</div>
  <div class="alert">⚠ {FX_SENSITIVITY["note"]}</div>
  <div style="display:flex;gap:12px;margin-bottom:12px;flex-wrap:wrap">
    <div style="background:var(--card);border-radius:9px;padding:12px 16px;
      box-shadow:0 1px 5px rgba(0,0,0,.07);flex:1;min-width:150px">
      <div style="font-size:10px;color:var(--muted);text-transform:uppercase">今日 USD/NTD 匯率</div>
      <div style="font-size:24px;font-weight:700;color:var(--navy);margin:3px 0">{FX_SENSITIVITY["spot"]}</div>
      <div style="font-size:11px;color:var(--muted)">台幣越小 = 台幣越強</div>
    </div>
    <div style="background:var(--card);border-radius:9px;padding:12px 16px;
      box-shadow:0 1px 5px rgba(0,0,0,.07);flex:1;min-width:150px">
      <div style="font-size:10px;color:var(--muted);text-transform:uppercase">毛利率影響 / 每 1% 升值</div>
      <div style="font-size:24px;font-weight:700;color:var(--red);margin:3px 0">{FX_SENSITIVITY["gm_bps_per_pct"]} bps</div>
      <div style="font-size:11px;color:var(--muted)">TSMC 官方法說會指引</div>
    </div>
    <div style="background:var(--card);border-radius:9px;padding:12px 16px;
      box-shadow:0 1px 5px rgba(0,0,0,.07);flex:1;min-width:150px">
      <div style="font-size:10px;color:var(--muted);text-transform:uppercase">外幣避險比例（估）</div>
      <div style="font-size:24px;font-weight:700;color:var(--navy);margin:3px 0">~75%</div>
      <div style="font-size:11px;color:var(--muted)">{FX_SENSITIVITY["hedge_ratio"]}</div>
    </div>
  </div>
  <table>
    <thead><tr><th>USD/NTD 匯率</th><th>情境說明</th><th>毛利率影響（bps）</th><th>EPS 影響（估）</th></tr></thead>
    <tbody>
      {"".join(f'''<tr{"" if s[0] != FX_SENSITIVITY["spot"] else ' style="background:#FFF8E1;font-weight:700"'}>
        <td><strong>{s[0]}</strong></td>
        <td>{s[1]}</td>
        <td style="color:{'var(--red)' if '-' in str(s[2]) else ('var(--green)' if '+' in str(s[2]) else 'var(--muted)')}">{s[2]}</td>
        <td style="color:{'var(--red)' if '-' in str(s[3]) else ('var(--green)' if '+' in str(s[3]) else 'var(--muted)')}">{s[3]}</td>
      </tr>''' for s in FX_SENSITIVITY["scenarios"])}
    </tbody>
  </table>
</div>

<!-- ═══ 19. EPS 修正追蹤 ═══ -->
<div class="sec" id="revision">
  <div class="sec-title">✏️ EPS 修正追蹤（Earnings Revision Tracker）</div>
  <div class="alert">⚠ {EPS_REVISION["note"]}</div>
  <div class="rev-grid">
    <div class="rev-box">
      <div class="rev-period">Q1 2026 EPS 共識（{EPS_REVISION["q1_analysts"]} 位分析師）</div>
      <div class="rev-consensus">NT${EPS_REVISION["q1_consensus"]:.2f}</div>
      <div class="rev-range">區間：NT${EPS_REVISION["q1_low"]:.2f} – NT${EPS_REVISION["q1_high"]:.2f}</div>
      <div style="margin-top:8px;display:flex;align-items:center;gap:8px">
        <span class="rev-trend-up">{EPS_REVISION["q1_trend"]}</span>
        <span style="font-size:11px;color:var(--green);font-weight:600">vs 3 個月前 {EPS_REVISION["q1_vs_3m_ago"]}</span>
      </div>
    </div>
    <div class="rev-box">
      <div class="rev-period">FY 2026 EPS 共識</div>
      <div class="rev-consensus">NT${EPS_REVISION["fy_consensus"]:.2f}</div>
      <div class="rev-range">區間：NT${EPS_REVISION["fy_low"]:.2f} – NT${EPS_REVISION["fy_high"]:.2f}</div>
      <div style="margin-top:8px;display:flex;align-items:center;gap:8px">
        <span class="rev-trend-up">{EPS_REVISION["fy_trend"]}</span>
        <span style="font-size:11px;color:var(--green);font-weight:600">vs 3 個月前 {EPS_REVISION["fy_vs_3m_ago"]}</span>
      </div>
    </div>
  </div>
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:12px">
    <div style="background:var(--card);border-radius:9px;padding:12px 14px;
      box-shadow:0 1px 5px rgba(0,0,0,.07);text-align:center">
      <div style="font-size:10px;color:var(--muted);text-transform:uppercase">連續超預期</div>
      <div style="font-size:28px;font-weight:700;color:var(--green);margin:4px 0">{EPS_REVISION["beat_streak"]}</div>
      <div style="font-size:11px;color:var(--muted)">季（100% 超預期）</div>
    </div>
    <div style="background:var(--card);border-radius:9px;padding:12px 14px;
      box-shadow:0 1px 5px rgba(0,0,0,.07);text-align:center">
      <div style="font-size:10px;color:var(--muted);text-transform:uppercase">平均超預期幅度</div>
      <div style="font-size:28px;font-weight:700;color:var(--green);margin:4px 0">{EPS_REVISION["beat_avg_pct"]}</div>
      <div style="font-size:11px;color:var(--muted)">近 {EPS_REVISION["beat_streak"]} 季均值</div>
    </div>
    <div style="background:var(--card);border-radius:9px;padding:12px 14px;
      box-shadow:0 1px 5px rgba(0,0,0,.07);text-align:center">
      <div style="font-size:10px;color:var(--muted);text-transform:uppercase">下次財報</div>
      <div style="font-size:18px;font-weight:700;color:var(--blue);margin:4px 0">{EPS_REVISION["next_report"]}</div>
      <div style="font-size:11px;color:var(--muted)">Q1 2026 法說會</div>
    </div>
  </div>
  <div class="sub-title">修正催化因子</div>
  <div class="trigger-list">
    {"".join(f'''<div class="trigger-item">
      <span class="trigger-badge" style="background:{'var(--green)' if t[0]=='正面' else ('var(--red)' if t[0]=='負面' else 'var(--orange)')}">{t[0]}</span>
      <span>{t[1]}</span>
    </div>''' for t in EPS_REVISION["triggers"])}
  </div>
  <div style="margin-top:12px;background:linear-gradient(135deg,#0D1B2A,#1565C0);color:#fff;
    border-radius:10px;padding:14px 18px;font-size:12.5px;line-height:1.75">
    <strong style="color:#FDD835">分析師觀點：</strong>{EPS_REVISION["analyst_view"]}
  </div>
</div>

<!-- FOOTER -->
<footer>
  TSMC 台積電日報 — {TODAY}<br>
  資料來源：Yahoo Finance、CNBC、Morgan Stanley、Goldman Sachs、JPMorgan、TSMC IR、TrendForce、玩股網、TWSE、BIS、Deloitte、PwC<br>
  <strong>本報告由自動化腳本產生，不構成任何投資建議。投資人應自行評估風險，必要時請諮詢持牌財務顧問。</strong>
</footer>
</body>
</html>"""

BASE_DIR = r"C:\Users\K748\OneDrive - 財團法人中華民國對外貿易發展協會\FET\Stock分析"
out = os.path.join(BASE_DIR, f"TSMC_日報_{TODAY}.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(HTML)
print(f"Done: {out}")

# ── Update manifest.json for GitHub Pages ─────────────────────────────
manifest_path = os.path.join(BASE_DIR, "manifest.json")
try:
    if os.path.exists(manifest_path):
        with open(manifest_path, "r", encoding="utf-8") as f:
            manifest = json.load(f)
    else:
        manifest = {"reports": [], "latest": "", "updated": ""}

    filename = f"TSMC_日報_{TODAY}.html"
    if not any(r["date"] == TODAY for r in manifest["reports"]):
        dt = datetime.strptime(TODAY, "%Y-%m-%d")
        manifest["reports"].append({
            "date": TODAY,
            "filename": filename,
            "label": f"{dt.year}年{dt.month:02d}月{dt.day:02d}日"
        })

    # Keep sorted newest-first
    manifest["reports"].sort(key=lambda x: x["date"], reverse=True)
    manifest["latest"]  = manifest["reports"][0]["date"]
    manifest["updated"] = TODAY

    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print(f"Manifest updated ({len(manifest['reports'])} reports): {manifest_path}")
except Exception as e:
    print(f"Warning: Could not update manifest.json: {e}")
    
