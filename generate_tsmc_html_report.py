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
    "tw_price":       "1,895",        # 收盤價（數字加逗號，不含 NT$）
    "tw_change":      "+44",          # 漲跌點數（含+/-號）
    "tw_change_pct":  "+2.37%",       # 漲跌幅（含+/-號）
    "tw_volume":      "39,307",       # 成交量（張）—估算（接近均量）
    "tw_open":        "1,855",
    "tw_high":        "1,900",
    "tw_low":         "1,848",
    "market_cap":     "NT$49.01 兆",
    "pe_ratio":       "~28.6x",       # 台股 P/E（EPS TTM=66.25, P=1895）
    "pb_ratio":       "~9.10x",       # 台股 P/B（股淨比）
    "div_yield":      "~1.21%",       # 殖利率（NT$23/NT$1,895）
    "52w_high_tw":    "2,025",        # 52 週高（2026-02-25）
    "52w_low_tw":     "780",          # 52 週低（2025-04-07）
    # NYSE TSM ADR
    "nyse_price":     "349.49",       # 收盤（3/12 EST 盤中 3PM）
    "nyse_change":    "-7.95",        # 美伊戰爭風險拖累
    "nyse_change_pct":"-2.22%",
    "52w_high_nyse":  "390.21",       # 52 週高（2026-02-25）
    "52w_low_nyse":   "134.25",       # 52 週低（2025-04-07）
    "pe_nyse":        "~24.7x",       # NYSE USD 基礎 P/E（EPS FY26E $14.14）
    "pb_nyse":        "~8.84x",
    "beta":           "1.56",
    "div_next":       "US$0.968/股",  # 下次除息（現金股利）
    "ex_date":        "2026-03-17",
    "next_earnings":  "2026-04-16",
    # 資料日期備註
    "data_note":      "台股數據：2026-03-12 收盤（估）  |  NYSE 數據：2026-03-12 盤中（3PM EST）  |  前2月累計營收年增29.9%（NT$7,189億）",
}

# 今日重點摘要（每日更新 — 白話文，供非專業讀者快速掌握）
SUMMARY = {
    # 整體信號：強烈看多 / 看多 / 中性偏多 / 中性 / 中性偏空 / 看空 / 強烈看空
    "sentiment":  "中性偏多",
    # 一句話標題
    "headline":   "台股逆勢反彈 +2.37%，NYSE ADR 受地緣風險壓抑，短期台美市場出現分歧",
    # 3–5 條白話重點（不用術語，每條約 30 字內）
    "bullets": [
        "台股 2330 收 NT$1,895（+2.37%）：NVIDIA 超越 Apple 成台積電最大客戶、N2 良率達 70-80% 雙利多激勵",
        "NYSE TSM 跌 -2.22% 至 $349.49：美伊戰爭風險升溫，科技股普遍走弱",
        "三大法人合計買超 5,470 張，外資連日買超，中長期立場偏多",
        "Q1 財報（4/16）倒計時 35 天，EPS 共識 NT$20.38，近 8 季全部超預期",
        "本週關注：3/17 除息（NT$6/股）、台幣匯率走勢（目前 31.68）、美伊情勢後續",
    ],
    # 一句結論
    "bottom_line": "AI 需求動能仍強，短線地緣政治雜音視為逢低機會，中長線趨勢偏多，法人目標均價 NT$2,290（潛在漲幅 +20.8%）。",
    # 重大警示（選填，空白則不顯示）
    "risk_alert":  "",
}

# 三大法人買賣超（張）— 最新一日
INSTITUTIONAL = {
    "date":     "2026-03-12",
    "foreign":  +4800,    # 外資（正=買超）— NVIDIA超越Apple成最大客戶利多激勵（估）
    "trust":    +920,     # 投信
    "dealer":   -250,     # 自營商（含避險）
    # 近5日（由舊到新）
    "foreign_5d": [+2800, +5230, +1200, +3200, +4800],
    "trust_5d":   [+800,  +1120, +700,  +850,  +920],
    "dealer_5d":  [+600,  +430,  -200,  -180,  -250],
    "foreign_ownership_pct": "76.3%",  # 外資持股比例（估）
    "foreign_ownership_shares": "約 196.8 億股",
}

# 技術分析指標（需每日從看盤軟體更新）
TECHNICAL = {
    "rsi_14":       61.2,   "rsi_signal": "偏多",
    "macd":         +2.1,   "macd_signal": "MACD 轉正，多頭訊號確認",   "macd_color": "#4CAF50",
    "kdj_k":        67.8,   "kdj_d": 58.4, "kdj_j": 87.3, "kdj_signal": "偏多",
    "ma5":       "1,885",
    "ma20":      "1,942",
    "ma60":      "1,830",
    "ma120":     "1,720",
    "bb_upper":  "2,060",   # 布林上軌
    "bb_mid":    "1,942",   # 布林中軌（20MA）
    "bb_lower":  "1,824",   # 布林下軌
    "vol_ma5":  "36,000",   # 5日均量（張）
    "trend":    "NVIDIA超越Apple成TSMC最大客戶、N2良率70-80%利多助攻，台股反彈+2.37%；NYSE受美伊地緣風險壓抑回落",
    "support":  "1,870–1,880",
    "resist":   "1,920–1,950",
    "note": "⚠️ 技術指標數值為估算，請以看盤軟體即時數據為準",
}

# 期貨與選擇權（來源：台灣期交所 TAIFEX）
DERIVATIVES = {
    "date":          "2026-03-12",
    "futures_price": "1,898",      # 台積電期貨近月收盤（估）
    "futures_oi":    "70,500",     # 未平倉量（口）（估）
    "futures_basis": "+3",         # 正價差/逆價差
    "call_oi":       "33,800",
    "put_oi":        "40,200",
    "pcr":           "1.19",       # Put/Call 比率（>1 偏空）
    "pcr_signal":    "偏空",
    "max_pain":      "1,900",      # 最大痛苦價格
    "key_strikes": [
        ("1,800", "Call 4,200 / Put 8,900", "強支撐"),
        ("1,850", "Call 5,100 / Put 7,300", "重要支撐"),
        ("1,900", "Call 9,800 / Put 10,200","最大痛苦點"),
        ("1,950", "Call 7,200 / Put 4,100", "重要壓力"),
        ("2,000", "Call 5,600 / Put 2,800", "強壓力"),
    ],
    "note": "⚠️ 衍生品數據需每日至台灣期交所(taifex.com.tw)更新",
}

# 供應鏈與客戶股價（需每日更新）
ECOSYSTEM_STOCKS = [
    # 類別、股票代碼、名稱、收盤價、漲跌幅、備註
    ("設備", "ASML",    "艾司摩爾",    "687.50", "-1.8%", "EUV 設備龍頭"),
    ("設備", "AMAT",    "應用材料",    "172.30", "-2.1%", "CVD/PVD 設備"),
    ("設備", "LRCX",    "科林研發",    "812.40", "-1.5%", "蝕刻設備"),
    ("設備", "KLAC",    "科磊",        "598.10", "-0.9%", "量測設備"),
    ("記憶", "000660.KS","SK 海力士",  "298,500 KRW", "+0.5%", "HBM3e/4 供應商"),
    ("記憶", "MU",      "美光",        "102.80", "-1.2%", "HBM / DDR5"),
    ("封裝", "3711.TW", "日月光",      "225",    "-0.9%", "CoWoS 封裝"),
    ("客戶", "NVDA",    "輝達",        "872.30", "-3.1%", "#1 客戶 21%"),
    ("客戶", "AAPL",    "蘋果",        "228.50", "-0.7%", "#2 客戶 17%"),
    ("客戶", "AVGO",    "博通",        "198.60", "-1.8%", "#3 客戶 13%"),
    ("客戶", "AMD",     "超微",        "162.40", "-2.4%", "#4 客戶 8%"),
    ("客戶", "2454.TW", "聯發科",      "1,350",  "-1.1%", "手機 SoC 客戶 9%"),
    ("競爭", "INTC",    "英特爾",      "21.50",  "+0.9%", "Intel Foundry"),
    ("競爭", "005930.KS","三星電子",   "54,800 KRW", "+0.3%", "Samsung Foundry"),
]

# 法人評級彙整
ANALYST_RATINGS = [
    ("Goldman Sachs",  "強烈買入", "NT$2,330", "2026-01-05", "AI 需求驅動 30% YoY 增長"),
    ("JPMorgan",       "增持",     "NT$2,100", "2026-01-07", "N3/N2 需求強勁，毛利率改善"),
    ("Susquehanna",    "正面",     "US$400",   "2026-01",    "AI 基礎設施資本支出持續"),
    ("Bernstein",      "優於大盤", "US$330",   "2026-01",    "技術領先優勢穩固"),
    ("Morgan Stanley", "增持",     "NT$2,290", "2026-02",    "市場共識平均目標"),
    ("MarketScreener", "共識買入", "NT$2,290", "2026-03-09", "31 位分析師、0 位賣出"),
]

CONSENSUS = {
    "rating":       "強烈買入",
    "avg_target_tw":"NT$2,290",
    "high_target":  "NT$2,770",
    "low_target":   "NT$1,740",
    "avg_target_us":"US$410",
    "analyst_count": 31,
    "buy": 31, "hold": 0, "sell": 0,
    "upside":       "+20.8%",   # 從 NT$1,895 到 NT$2,290
}

# 財報倒計時
EARNINGS = {
    "next_date":   "2026-04-16",
    "quarter":     "Q1 2026",
    "rev_guide_lo":"$34.6B",
    "rev_guide_hi":"$35.8B",
    "rev_guide_mid":"$35.2B",
    "gm_guide":    "63–65%",
    "op_guide":    "54–56%",
    "eps_consensus":"NT$20.38 / US$3.26",
    "last_beat":   "+7.23%",
    "beat_rate":   "100%（近 8 季全部超預期）",
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
    ("2026-03-17", "除息日",       "2026 Q1 現金股利 NT$6.00/股",                   "#E91E63"),
    ("2026-04-16", "Q1 2026 財報", "預計法說會 15:00 TST；投資人電話會議同日",      "#F44336"),
    ("2026-06",    "2026 技術論壇","North America Technology Symposium（預估）",      "#9C27B0"),
    ("2026-07",    "Q2 2026 財報", "預計 7 月第三週（估）",                          "#FF9800"),
    ("2026-09",    "月營收持續",   "每月 10 日前公告上月合併營收",                    "#2196F3"),
    ("2026-10",    "Q3 2026 財報", "預計 10 月第三週（估）",                         "#FF9800"),
    ("2026-12",    "投資人日",     "TSMC North America Investor Day（不定期）",       "#9C27B0"),
]

# 總體經濟監控
MACRO = [
    ("SOX 半導體指數",  "4,356",   "-0.6%",  False),
    ("美元指數 DXY",    "103.2",   "-0.1%",  False),
    ("VIX 恐慌指數",    "22.1",    "+0.9",   False),  # 美伊風險持續
    ("WTI 原油",        "$69.3",   "+1.2%",  False),  # 地緣政治推升油價
    ("10Y 美債殖利率",  "4.31%",   "+0.02pp",False),
    ("台幣/美元",       "31.68",   "-0.00",  False),
    ("NVIDIA NVDA",     "$864",    "-0.9%",  False),
    ("ASML",            "$681",    "-0.9%",  False),
]

# ═══════════════════════════════════════════════════════════════════
#  進階分析師模組資料（每日同步更新）v3.0
# ═══════════════════════════════════════════════════════════════════

# 估值歷史位階 (Valuation Context)
VALUATION = {
    "pe_current":   28.6,   # 同 STOCK["pe_ratio"]
    "pe_5y_avg":    18.5,   # 5 年均 P/E（估算）
    "pe_5y_high":   37.8,   # 5 年高點（2021 AI 泡沫）
    "pe_5y_low":     9.8,   # 5 年低點（2022 下行周期）
    "pe_5y_pct":      74,   # 目前在 5 年區間的百分位（%）
    "pe_10y_avg":   16.2,   # 10 年均 P/E（估算）
    "pe_10y_pct":     80,   # 目前在 10 年區間的百分位（%）
    "pb_current":   9.20,   # 同 STOCK["pb_ratio"]
    "pb_5y_avg":     5.8,
    "pb_5y_pct":      78,
    "ev_ebitda":    17.8,   # EV/EBITDA（估算）
    "peg":          0.88,   # PEG = P/E ÷ 5 年 EPS CAGR ~33%
    "note": "P/E 歷史百分位為估算（Bloomberg/Refinitiv 基準），請每季更新基準值",
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
    "date":           "2026-03-12",
    "tsmc_tw_1d":     +2.37,   # 2330 當日（NVIDIA超越Apple利多催化）
    "taiex_1d":       +0.68,   # 加權指數（估算）
    "sox_1d":         -0.60,   # SOX（同 MACRO 資料）
    "tsmc_nyse_1d":   -2.22,   # TSM ADR 當日（美伊風險拖累）
    "tsmc_tw_5d":     +4.1,
    "taiex_5d":       +1.4,
    "tsmc_tw_1m":     -3.8,
    "taiex_1m":       -1.5,
    "tsmc_tw_ytd":    -4.6,
    "taiex_ytd":      -2.0,
    "sox_ytd":        -12.8,
    "beta_60d":        1.56,   # 60 日滾動 Beta vs TAIEX
    # ADR 折溢價
    "adr_price":     349.49,   # TSM NYSE 盤中（3PM EST 3/12）
    "adr_parity":    299.21,   # 理論值 = 5 × NT$1,895 ÷ 31.68
    "adr_ratio":     "5:1",    # 1 ADR = 5 台積電普通股
    "usd_ntd":        31.68,   # 同 MACRO 台幣匯率
    "adr_premium_pct": "+16.8",# (349.49 - 299.21) / 299.21 × 100
    "note": "ADR 理論值 = 5 × 台股收盤 ÷ 匯率（不含交易摩擦與流動性溢價）。TAIEX/多期間數據為估算",
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
    "spot":              31.68,  # 今日匯率（同 MACRO）
    "gm_bps_per_pct":     -40,  # 台幣每升值 1%，毛利率下降 ~40bps（法說會指引）
    "scenarios": [
        (30.0, "NTD +5.6%（強升）", "約 -224 bps", "約 -7.8%"),
        (31.0, "NTD +2.2%（小升）", "約 -88 bps",  "約 -3.1%"),
        (31.68,"基準（今日）",       "—",            "—"),
        (32.5, "NTD -2.6%（小貶）", "約 +104 bps",  "約 +3.6%"),
        (34.0, "NTD -7.3%（大貶）", "約 +292 bps",  "約 +10.2%"),
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
    "beat_streak":    8,       # 連續超預期季數
    "beat_avg_pct":  "+5.8%",  # 近 8 季平均超預期幅度
    "next_report":   "2026-04-16",
    "triggers": [
        ("正面", "2月月營收年增22.2%，催化法人上修EPS預估"),
        ("正面", "N3/N2 ASP調升，毛利率改善趨勢明確"),
        ("負面", "全球AI晶片出口管制擴大，訂單能見度風險"),
        ("中性", "台幣升值壓力，毛利率中性至下修風險"),
    ],
    "analyst_view": "目前 Q1 EPS 共識 NT$20.38，2 月營收年增 22.2% 後，預計法人將在 4/16 財報前陸續上修。EPS 有望挑戰 NT$21–22，此為股價突破 NT$2,025 歷史高點的核心動能。",
    "note": "資料來源：Bloomberg / MarketScreener（估算），請每日核實更新",
}

NEWS = [
    {
        "tag": "客戶動態", "tag_color": "#2196F3",
        "title": "NVIDIA 超越 Apple 成 TSMC 最大客戶，2026 年貢獻營收估達 22%",
        "body": "CNBC 確認 NVIDIA 已於 2026 年正式取代 Apple，成為台積電最大單一客戶。NVIDIA 預計全年貢獻 TSMC 約 $330 億美元營收（佔比 ~22%），Apple 則降至第二（$270 億，~18%）。NVIDIA CEO Jensen Huang 近期播客已確認此轉變。分析師 Ben Bajarin 指出，客戶結構重心從消費電子轉向 AI 基礎設施，有利 TSMC 長期 ASP 提升，且 NVIDIA 為 TSMC A16（1.6nm）製程的首要客戶，為未來 3 年營收成長提供強大能見度。",
        "source": "CNBC / MacRumors 2026-01-26～28", "impact": "正面",
    },
    {
        "tag": "產品發布", "tag_color": "#9C27B0",
        "title": "TSMC N2（2nm）HVM 良率達 70-80%，N3 全面滿載",
        "body": "2026 年 3 月報告顯示，TSMC N2 製程已在竹科 Fab 20 及高雄 Fab 22 順利進入高量產（HVM）階段，邏輯測試晶片良率達 70-80%，部分 SRAM 區塊良率超過 90%。N3/N3P 繼續以近 100% 稼動率滿載，主要服務 NVIDIA AI GPU 與 Apple SoC。AMD 將成為首個以 2nm 生產 AI 晶片的客戶，NVIDIA 則瞄準更先進的 A16（1.6nm）製程。整體先進製程良率超出市場預期，有助支撐 Q1 2026 毛利率指引 63-65%。",
        "source": "Cyberraiden / TrendForce 2026-03-11", "impact": "正面",
    },
    {
        "tag": "營運數據", "tag_color": "#4CAF50",
        "title": "前 2 月累計營收年增 29.9%，略低預期 33% 增速",
        "body": "台積電 2026 年前 2 月合計月營收 NT$7,189.1 億（約 US$226 億），年增 29.9%，略低於法人原預期 33%。2 月單月月減 20.8%（工作天數少），年增 22.2%。成長偏弱主因記憶體晶片需求拖累部分訂單，惟 AI 晶片（HPC 平台）持續強勁，N3/N5 滿載。TSMC 全年 2026 資本支出已提升至 $520-560 億美元（年增 37%），反映管理層對後續需求信心不減，Q1 財報（4/16）前法人將持續上修 EPS 預估。",
        "source": "Bloomberg / Yahoo Finance 2026-03-10", "impact": "正面",
    },
    {
        "tag": "擴產計畫", "tag_color": "#9C27B0",
        "title": "TSMC 快速推進台南 Mega Fab，搶攻 AI 晶片擴產先機",
        "body": "台積電已完成台南新巨型廠環評程序，加速進入建設階段，目標 2028 年完工。新廠主力製程針對 AI 硬體與高效能運算晶片。同時，美國 Arizona Phase 2（2nm）建設持續，Phase 1（3nm）已全面量產。2026 年 TSMC 全球 CapEx 達 $520-560 億美元（年增 37%），為歷史最高值，有望在 2027-28 年顯著拉升可用先進製程產能，滿足 NVIDIA、AMD、Apple 的長期訂單需求。",
        "source": "Simply Wall St / Digitimes 2026-03-11", "impact": "正面",
    },
    {
        "tag": "地緣政治", "tag_color": "#607D8B",
        "title": "美伊戰爭風險持續，NYSE TSM 3/12 盤中跌 2.22% 至 $349.49",
        "body": "美伊衝突風險令科技股承壓，NYSE TSM 於美東 3/12（周四）盤中跌 $7.95（-2.22%）至 $349.49，較 2/25 高點 $390.21 回落逾 10%。地緣風險疊加全球 AI 晶片出口管制預期，拖累半導體股廣泛走弱（SOX 下跌 0.6%）。台灣本地 2330 則受 NVIDIA 超越 Apple 成最大客戶、N2 良率亮眼等利多催化，3/12 逆勢上漲 +2.37% 至 NT$1,895，顯示境內外市場短期情緒分化。",
        "source": "StockTwits / Yahoo Finance / Bloomberg 2026-03-12", "impact": "負面",
    },
    {
        "tag": "法規政策", "tag_color": "#FF9800",
        "title": "美商務部擬立法管控全球 AI 晶片出口，訂單能見度存在中期風險",
        "body": "美國商務部起草法規，擬對全球所有目的地 AI 晶片出口實施美方逐案審批，賦予華盛頓廣泛否決權。若生效，NVIDIA、AMD 等 TSMC 主要客戶的海外銷售計畫將受重大影響，間接降低 TSMC 2027-28 年訂單能見度。TSMC 目前中國市場收入佔比約 15%，為最高風險暴露區域。短期影響有限（NVIDIA Blackwell/Rubin 訂單已鎖定），中期需密切關注具體條文。",
        "source": "Bloomberg 2026-03-05", "impact": "負面",
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
