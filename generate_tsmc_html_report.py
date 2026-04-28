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
    "tw_price":       "2,265",        # 2026-04-27 收盤（+80 vs 4/24 NT$2,185，+3.66%；盤中觸及 NT$2,330 史上最大單日漲點）
    "tw_change":      "+80",          # 漲跌點數（vs 4/24 NT$2,185；盤中曾 +145 創史上最大）
    "tw_change_pct":  "+3.66%",       # 漲跌幅
    "tw_volume":      "~62,800",      # 成交量（張，4/27 實際，量能再放大 +40%）
    "tw_open":        "2,200",        # 4/27 開盤
    "tw_high":        "2,330",        # 4/27 盤中史高（與股號 2330 神奇巧合）
    "tw_low":         "2,200",        # 4/27 日中低
    "market_cap":     "NT$58.7 兆",   # 4/27 收盤市值（盤中曾突破 NT$60 兆，史上最高）
    "pe_ratio":       "~33.4x",       # 台股 P/E（TTM，NT$2,265 基礎）
    "pb_ratio":       "~11.8x",       # 台股 P/B（股淨比，NT$2,265 基礎）
    "div_yield":      "~0.66%",       # 殖利率（NT$2,265 基礎續下調）
    "52w_high_tw":    "2,330",        # 52 週高（2026-04-27 盤中創新高）
    "52w_low_tw":     "780",          # 52 週低（2025-04-07）
    # NYSE TSM ADR
    "nyse_price":     "422.10",       # 4/27 實際收盤（+4.88% vs 4/24 $402.46；續創歷史新高）
    "nyse_change":    "+19.64",       # 4/27 實際（vs 4/24 $402.46）
    "nyse_change_pct":"+4.88%",       # 4/27 跟進台股共振創高
    "52w_high_nyse":  "422.10",       # 52 週高（2026-04-27 創新高）
    "52w_low_nyse":   "134.25",       # 52 週低（2025-04-07）
    "pe_nyse":        "~35.0x",       # NYSE USD 基礎 P/E（TTM，估算，$422.10 基礎）
    "pb_nyse":        "~12.4x",
    "beta":           "1.18",
    "div_next":       "下次配息預計 2026-06",
    "ex_date":        "2026-04-09（Q2 除息完成，US$0.968/ADR）",
    "next_earnings":  "2026-07（Q2 2026 財報，7 月第三週預估）",
    # 資料日期備註
    "data_note":      "⭐ 今日（2026-04-28 Tue）重點：台股 2330 4/27 收盤 NT$2,265（+80，+3.66%）創歷史新高，盤中觸及 NT$2,330 與股號完美巧合、單日漲點 +145 創史上最大紀錄；市值盤中突破 NT$60 兆、收盤 NT$58.7 兆雙創歷史新高；NYSE TSM 4/27 收 $422.10（+4.88%）續創歷史新高；台股加權指數 4/27 收 40,194 點站上 4 萬大關、單日大漲 1,262 點  |  ⚠️ 今日（4/28 Tue）TSM 盤前 -2.44% 至 $412.28，獲利了結湧現  |  4/27 外資雖賣超 TSMC 1.8 萬張（前三名 ETF 調節），全市場三大法人賣超 482 億，但散戶與短線資金接力推升  |  ⚠️ 法律事件：前 TSMC 工程師陳力銘因外洩 2nm 機密予東京威力科創（Tokyo Electron）4/27 經智財法院判處 10 年徒刑、東京威力科創被罰 NT$1.5 億  |  TSMC A13/A12/N2U 三大新製程路線確認 2029 量產；Q1 2026 淨利 $18.2B（+58.3% YoY）、毛利率 66.2% 雙創歷史新高",
}

# 今日重點摘要（每日更新 — 白話文，供非專業讀者快速掌握）
SUMMARY = {
    # 整體信號：強烈看多 / 看多 / 中性偏多 / 中性 / 中性偏空 / 看空 / 強烈看空
    "sentiment":  "看多",
    # 一句話標題
    "headline":   "🎯 台股 2330 4/27 收 NT$2,265（+3.66%）盤中觸 2330 與股號神奇巧合、單日漲點 +145 史上最大；市值盤中破 NT$60 兆；NYSE TSM 收 $422.10 續創高；今日（4/28）TSM 盤前 -2.44% 獲利了結；前 TSMC 工程師 2nm 機密外洩判 10 年",
    # 3–5 條白話重點（不用術語，每條約 30 字內）
    "bullets": [
        "🎯 台股 2330 4/27 收 NT$2,265（+80，+3.66%）盤中觸 NT$2,330 與股號完美巧合；漲點 +145 史上最大；市值盤中破 NT$60 兆",
        "💵 NYSE TSM 4/27 收 $422.10（+4.88%）續創史高；台股加權 4/27 站上 40,194（+1,262）4 萬大關；台美 2 連漲動能延續",
        "⚠️ 今日（4/28 Tue）TSM 盤前 -2.44% 至 $412.28 獲利了結；4/27 外資反向賣超 TSMC 18,940 張、全市場三大法人賣超 482 億",
        "⚖️ 前 TSMC 工程師陳力銘 4/27 因外洩 2nm 製程機密予東京威力科創遭判 10 年徒刑、東京威力科創罰 NT$1.5 億；商業間諜重大判例",
        "⭐ 共識目標 NT$2,320 距 4/27 NT$2,265 僅剩 +2.4%；Barclays $470 對應 NT$2,450 剩 +8.2%；新催化需 Q2 7 月財報",
    ],
    # 一句結論
    "bottom_line": "今日（2026-04-28 Tue）報告摘要：台股 2330 昨日（4/27 Mon）續創歷史新高收 NT$2,265（+80 點，+3.66% vs 4/24 NT$2,185），盤中一度觸及 NT$2,330 與股票代號完美巧合，單日漲點 +145 創台股史上最大紀錄；市值盤中突破 NT$60 兆、收盤 NT$58.7 兆續創新高，台股加權指數同日大漲 +1,262 點站上 40,194 突破 4 萬大關。NYSE TSM 4/27 同步收 $422.10（+4.88%）續創歷史新高，台美雙市 2 連漲動能延續。然而今日（4/28 Tue）盤前 TSM 重挫 -2.44% 至 $412.28，獲利了結湧現提示短線過熱拉回風險升高。籌碼面 4/27 出現背離訊號：股價創高但外資反手賣超 TSMC 1.8940 萬張（前三名皆為 ETF 調節）、全市場三大法人合計賣超 NT$482 億（外資 -512 億、投信 +135 億、自營 -105 億），顯示法人逢高調節，散戶接棒推升。法律重大事件：前 TSMC Fab12 良率部門工程師陳力銘（離職後加入東京威力科創）因 2023 下半年至 2024 上半年向 TSMC 在職工程師索取 2nm 製程關鍵機密拍照外洩，4/27 經智財及商業法院依違反國家安全法判處 10 年徒刑；東京威力科創員工陸怡寅判 10 個月緩刑、公司被罰 NT$1.5 億並須賠 TSMC NT$1 億；同案另兩名 TSMC 在職工程師判 2-3 年；此案為台灣首宗 2nm 等級商業機密判例，凸顯先進製程護國神山的法律保護升級。基本面延續：Q1 2026 淨利 $18.2B（+58.3% YoY）、毛利率 66.2% 雙創史高；2026 全年成長 >30% 確認；A13/A12/N2U 路線至 2029 量產；NVIDIA #1 客戶 22% 確認。技術面：RSI 估升至 75+ 嚴重超買、KDJ J 值飆破 95 極端超買、布林上軌突破，短線過熱整理風險升高；支撐 NT$2,180-2,200（5MA + 心理整數），壓力 NT$2,330（4/27 史高）/ NT$2,400-2,450（共識區間）。31 位分析師全數看多，共識目標 NT$2,320 距現價剩 +2.4%、Barclays $470 對應 NT$2,450 剩 +8.2%，需 Q2 7 月財報催化下一波上修。",
    # 重大警示（選填，空白則不顯示）
    "risk_alert":  "⚠️ 短線過熱拉回風險升高：（1）今日（4/28 Tue）TSM 盤前重挫 -2.44% 至 $412.28，台股 2330 開盤恐承壓；（2）4/27 外資反手賣超 TSMC 1.8940 萬張、全市場三大法人賣超 NT$482 億，籌碼出現逢高調節；（3）RSI 估 75+ 嚴重超買、KDJ J 值破 95 極端超買、布林上軌已突破，短線過熱訊號明確；（4）共識目標 NT$2,320 已達 97.6%（剩 +2.4%），上行空間幾無；（5）2nm 機密外洩判決雖屬利多（保護升級），但凸顯競爭對手覬覦先進製程；（6）台幣 ~31.35 升值壓力（每升值 1% 毛利率 -40bps）；（7）CoWoS / ABF 基板瓶頸與美中 AI 晶片出口管制結構性雜音續存；（8）盤中觸 NT$2,330 與股號巧合形成心理頂部標的，新一輪上修需 7 月 Q2 財報催化",
}

# 三大法人買賣超（張）— 最新一日
INSTITUTIONAL = {
    "date":     "2026-04-27",
    "foreign":  -18940,   # 外資（負=賣超）— 4/27 實際，創高反手大舉賣超 1.8940 萬張，前三名皆 ETF 調節
    "trust":    +1850,    # 投信（4/27 估，續加碼挺升）
    "dealer":   -420,     # 自營商（4/27 估，獲利了結轉賣）
    # 近5日（由舊到新，最後一筆為 4/27）
    "foreign_5d": [+2210, +3500, +4200, +8306, -18940],  # 4/27：外資創高反手大賣
    "trust_5d":   [+696,  +820,  +620,  +1168, +1850],   # 4/27：投信連 6 日續買超
    "dealer_5d":  [-77,   +150,  +80,   +256,  -420],    # 4/27：自營轉賣
    "foreign_ownership_pct": "75.2%",  # 外資持股比例（4/27 略降）
    "foreign_ownership_shares": "約 195.1 億股",
}

# 技術分析指標（需每日從看盤軟體更新）
TECHNICAL = {
    "rsi_14":       76.8,   "rsi_signal": "RSI 飆升至 76.8 嚴重超買（自 72.5 續升 4.3 點），4/27 +3.66% 連 2 日大漲推升動能；4/28 盤前 ADR -2.44% 暗示拉回壓力升高",
    "macd":         +7.2,   "macd_signal": "MACD 正值持續擴大（+5.8→+7.2），多頭動能延續但漲幅趨緩；爆量大漲後注意背離訊號",   "macd_color": "#4CAF50",
    "kdj_k":        92.0,   "kdj_d": 86.5, "kdj_j": 98.5, "kdj_signal": "KDJ 三線深陷超買區（K/D/J 全 >85），J 值 98.5 達極端超買；短線過熱拉回風險警示升至紅色",
    "ma5":       "2,150",
    "ma20":      "2,065",
    "ma60":      "1,980",
    "ma120":     "1,870",
    "bb_upper":  "2,260",   # 布林上軌（隨 4/27 創高上移）
    "bb_mid":    "2,065",   # 布林中軌（20MA）
    "bb_lower":  "1,955",   # 布林下軌
    "vol_ma5":  "32,800",   # 5日均量（張，連 2 日爆量後再上修）
    "trend":    "台股 2330 4/27 收 NT$2,265（+80，+3.66%）續創歷史新高，盤中觸及 NT$2,330 與股號完美巧合、單日漲點 +145 創史上最大紀錄，量能再放大至約 62,800 張；NYSE TSM 4/27 同步收 $422.10（+4.88%）續創史高。技術面進入極端超買：RSI 76.8 嚴重超買、MACD +7.2 動能擴大但趨緩、KDJ J 值 98.5 達極端超買警示。均線完美多頭排列（MA5>MA20>MA60>MA120），現價遠高於所有均線。⚠️ 短線過熱拉回風險升至紅色：今日（4/28 Tue）TSM 盤前 -2.44% 至 $412.28、布林上軌已突破、KDJ J 值極端超買，預期台股開盤承壓；支撐 NT$2,180-2,200（5MA + 心理整數），第二支撐 NT$2,150；壓力 NT$2,330（4/27 史高），中線目標 NT$2,400-2,450（Barclays $470 對應）。共識目標 NT$2,320 距現價僅剩 +2.4%。",
    "support":  "2,180–2,200",
    "resist":   "2,330",
    "note": "⚠️ 技術指標數值為估算，請以看盤軟體即時數據為準；4/28 盤前訊號顯示短線拉回機率升高",
}

# 期貨與選擇權（來源：台灣期交所 TAIFEX）
DERIVATIVES = {
    "date":          "2026-04-27",
    "futures_price": "2,278",      # 台積電期貨近月（估，4/27 收盤，正價差續擴大）
    "futures_oi":    "118,500",    # 未平倉量（口）（估，連 2 日爆量多單留倉激增）
    "futures_basis": "+13",        # 正價差續擴大（多頭情緒延續但已觸警戒）
    "call_oi":       "78,200",
    "put_oi":        "39,800",
    "pcr":           "0.51",       # Put/Call 比率續降至極端值（多頭極度擁擠）
    "pcr_signal":    "極度偏多警戒（Put/Call 0.51，較 4/24 的 0.57 續降，4/27 創高伴隨選擇權偏多達極端，反向訊號出現）",
    "max_pain":      "2,200",      # 最大痛苦價格（近期到期，估，隨現貨上移）
    "key_strikes": [
        ("2,150", "Call 22,200 / Put 6,800", "5MA 附近第一支撐"),
        ("2,200", "Call 32,500 / Put 5,200", "心理整數，4/27 開盤價"),
        ("2,265", "Call 38,500 / Put 3,200", "4/27 收盤價，短線多空分界"),
        ("2,330", "Call 28,500 / Put 2,200", "4/27 史高，與股號神奇巧合"),
        ("2,400", "Call 18,200 / Put 1,400", "近共識目標 NT$2,320 上方"),
        ("2,450", "Call 10,500 / Put 700",   "Barclays $470 對應目標"),
    ],
    "note": "⚠️ 衍生品數據需每日至台灣期交所(taifex.com.tw)更新；Put/Call 0.51 達極端偏多反向訊號",
}

# 供應鏈與客戶股價（需每日更新）
ECOSYSTEM_STOCKS = [
    # 類別、股票代碼、名稱、收盤價、漲跌幅、備註（2026-04-27 收盤，台美續漲第 19 日、TSMC 觸 NT$2,330 史高）
    ("設備", "ASML",    "艾司摩爾",    "705.20", "+2.4%", "EUV 設備龍頭（4/27 續漲；Siemens-TSMC AI 設計合作擴大）"),
    ("設備", "AMAT",    "應用材料",    "187.40", "+3.9%", "CVD/PVD 設備（4/27 續創高；2nm/N2U 訂單支撐）"),
    ("設備", "LRCX",    "科林研發",    "853.20", "+4.0%", "蝕刻設備（4/27 大漲；TSMC 2nm 機密案凸顯設備護城河）"),
    ("設備", "KLAC",    "科磊",        "632.50", "+4.0%", "量測設備（4/27 跟進；A13/A12 驗證需求加碼）"),
    ("記憶", "000660.KS","SK 海力士",  "315,500 KRW", "+2.4%", "HBM4 供應（NVIDIA GB300 主力，Samsung HBM4E 競爭）"),
    ("記憶", "MU",      "美光",        "109.80", "+5.1%", "HBM / DDR5（4/27 續強；DRAM 漲價 50% YTD）"),
    ("封裝", "3711.TW", "日月光",      "258",    "+5.3%", "CoWoS 封裝（4/27 續飆漲；Arizona 2029 封裝廠催化）"),
    ("客戶", "NVDA",    "輝達",        "961.50", "+5.0%", "#1 客戶 22%（4/27 創高；$5T 市值在望）"),
    ("客戶", "AAPL",    "蘋果",        "218.20", "+1.3%", "#2 客戶 17%（A20 採 2nm，iPhone 18 鎖定 A14）"),
    ("客戶", "AVGO",    "博通",        "221.50", "+5.2%", "#3 客戶 13%（Custom XPU，AI ASIC 強勢）"),
    ("客戶", "AMD",     "超微",        "192.80", "+5.8%", "#4 客戶 8%（MI450 新大單、法國 AI 戰略合作）"),
    ("客戶", "2454.TW", "聯發科",      "2,355",  "+3.5%", "手機 SoC 客戶 9%，2nm 首批量產，連 4 日創高"),
    ("競爭", "INTC",    "英特爾",      "23.40",  "+2.6%", "Intel Foundry 18A 爬坡；2nm 機密案利好對手定位"),
    ("競爭", "005930.KS","三星電子",   "56,400 KRW", "+2.2%", "Samsung Foundry 良率改善，HBM4E 送樣 NVIDIA"),
]

# 法人評級彙整
ANALYST_RATINGS = [
    ("Barclays",       "增持",     "US$470",   "2026-04-22", "⭐ 4/22 升評！AI 需求結構性+A14/A16 領先，目標大幅上調（NT$2,450 對應）"),
    ("Goldman Sachs",  "強烈買入", "NT$2,330", "2026-04-16", "Conviction Buy，Q1 財報後上調 35%；4/24 升評台廠封測股"),
    ("JPMorgan",       "增持",     "NT$2,200", "2026-04-17", "N3/N2 需求強勁，毛利率改善，財報後上調"),
    ("Susquehanna",    "正面",     "US$425",   "2026-04",    "AI 基礎設施資本支出持續，財報後上調"),
    ("Bernstein",      "優於大盤", "US$360",   "2026-04",    "技術領先優勢穩固，財報後上調"),
    ("Morgan Stanley", "增持",     "NT$2,290", "2026-04",    "市場共識平均目標"),
    ("MarketScreener", "共識買入", "NT$2,320", "2026-04-27", "31 位分析師、0 位賣出（4/24 NT$2,185 上漲後共識僅剩 +6.2% 空間，需新一輪上修）"),
]

CONSENSUS = {
    "rating":       "強烈買入",
    "avg_target_tw":"NT$2,320",  # 共識（Barclays 升評後）— 4/27 NT$2,265 已達 97.6%
    "high_target":  "NT$2,770",
    "low_target":   "NT$1,740",
    "avg_target_us":"US$525",    # 分析師上調目標（Barclays $470 上拉均值）
    "analyst_count": 31,
    "buy": 31, "hold": 0, "sell": 0,
    "upside":       "+2.4%",    # 從 NT$2,265（4/27 收盤）到 NT$2,320 — 共識空間已近耗盡
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
    ("2026-04-27", "🎯 盤中觸 NT$2,330 與股號神奇巧合、單日漲點 +145 史上最大","台股 2330 4/27 收 NT$2,265（+80，+3.66%）續創歷史新高；盤中一度衝至 NT$2,330 與股票代號完美巧合、單日漲點 +145 創台股史上最大紀錄；市值盤中突破 NT$60 兆、收盤 NT$58.7 兆雙創新高；台股加權指數同日 +1,262 點站上 40,194 突破 4 萬大關；NYSE TSM 4/27 收 $422.10（+4.88%）續創史高；4/27 外資反手賣超 TSMC 1.8940 萬張（前三名皆 ETF 調節）、全市場三大法人合計賣超 NT$482 億；前 TSMC 工程師陳力銘 4/27 因 2nm 機密外洩予東京威力科創遭智財法院判處 10 年徒刑、東京威力科創罰 NT$1.5 億","#9E9E9E"),
    ("2026-04-28", "⭐ 今日：盤前 ADR -2.44% 拉回，留意過熱整理","昨日（4/27）台美雙市再創史高後，今日（4/28 Tue）市場關注：（1）TSM 盤前 -2.44% 至 $412.28，台股開盤恐承壓拉回；（2）短線指標進入極端超買（RSI 76.8、KDJ J 98.5、布林上軌突破），回檔風險升至紅色；（3）4/27 外資反手賣超 1.8940 萬張，籌碼面出現逢高調節訊號；（4）共識目標 NT$2,320 已達 97.6%，僅剩 +2.4% 空間；（5）今日支撐 NT$2,180-2,200（5MA + 心理整數）、第二支撐 NT$2,150；壓力 NT$2,330（4/27 史高）；（6）2nm 機密外洩判決利多護國神山法律保護升級","#BBDEFB"),
    ("2026-06",    "2026 技術論壇","North America Technology Symposium（已於 4/22 舉辦，宣布 A13/A12/N2U）",  "#9E9E9E"),
    ("2026-07",    "Q2 2026 財報", "預計 7 月第三週（估）",                                        "#FF9800"),
    ("2026-09",    "月營收持續",   "每月 10 日前公告上月合併營收",                                  "#2196F3"),
    ("2026-10",    "Q3 2026 財報", "預計 10 月第三週（估）",                                        "#FF9800"),
]

# 總體經濟監控
MACRO = [
    ("SOX 半導體指數",  "10,840.20","+3.10%", False),  # 4/27 SOX 連 19 日漲續創史高
    ("美元指數 DXY",    "97.5",    "-0.3%",  False),  # 美元續弱，利好 TSMC ADR 估值
    ("VIX 恐慌指數",    "17.8",    "-0.4",   False),  # 恐慌指數續降，風險偏好攀升
    ("WTI 原油",        "$63.5",   "+0.5%",  False),  # 油價小升
    ("10Y 美債殖利率",  "3.89%",   "-0.03pp",False),  # 殖利率回落，風險資產受惠
    ("台幣/美元",       "31.30",   "-0.05",  False),  # 台幣略升（每升值 1% 毛利率 -40bps，需追蹤）
    ("NVIDIA NVDA",     "$961.50", "+5.0%",  False),  # #1 TSMC 客戶（22%），4/27 續創高、$5T 市值在望
    ("ASML",            "$705.20", "+2.4%",  False),  # 設備股續強（Siemens-TSMC AI 設計合作）
]

# ═══════════════════════════════════════════════════════════════════
#  進階分析師模組資料（每日同步更新）v3.0
# ═══════════════════════════════════════════════════════════════════

# 估值歷史位階 (Valuation Context)
VALUATION = {
    "pe_current":   33.4,   # 同 STOCK["pe_ratio"]（4/27 NT$2,265）
    "pe_5y_avg":    18.5,   # 5 年均 P/E（估算）
    "pe_5y_high":   37.8,   # 5 年高點（2021 AI 泡沫）
    "pe_5y_low":     9.8,   # 5 年低點（2022 下行周期）
    "pe_5y_pct":      85,   # 目前在 5 年區間的百分位（%，續創高後上升）
    "pe_10y_avg":   16.2,   # 10 年均 P/E（估算）
    "pe_10y_pct":     89,   # 目前在 10 年區間的百分位（%）
    "pb_current":  11.80,   # 同 STOCK["pb_ratio"]（4/27 NT$2,265）
    "pb_5y_avg":     5.8,
    "pb_5y_pct":      92,   # P/B 續創高
    "ev_ebitda":    19.8,   # EV/EBITDA（估算，續隨股價上行）
    "peg":          1.01,   # PEG = P/E ÷ 5 年 EPS CAGR ~33%
    "note": "P/E 歷史百分位為估算（Bloomberg/Refinitiv 基準），請每季更新基準值；4/27 續創高後估值百分位偏高",
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
    "date":           "2026-04-27",
    "tsmc_tw_1d":     +3.66,   # 2330 4/27 收盤 NT$2,265（+80 vs 4/24 NT$2,185）
    "taiex_1d":       +3.24,   # 加權指數 4/27 大漲 +1,262 站上 40,194
    "sox_1d":         +3.10,   # SOX 4/27 連 19 日漲（估）
    "tsmc_nyse_1d":   +4.88,   # TSM ADR 4/27 $422.10（+4.88% vs 4/24 $402.46）
    "tsmc_tw_5d":     +11.58,  # 5 日（NT$2,265 vs 4/20 NT$2,030）
    "taiex_5d":       +6.50,
    "tsmc_tw_1m":     +16.7,   # 月線（NT$2,265 vs ~NT$1,940 一個月前）
    "taiex_1m":       +9.5,
    "tsmc_tw_ytd":    +8.1,    # Q1+4/1–4/27 累計（NT$2,265 vs 年初 ~NT$2,095）
    "taiex_ytd":      +3.7,
    "sox_ytd":        -5.6,    # SOX 連 19 日漲收復多數失土
    "beta_60d":        1.20,   # 60 日滾動 Beta vs TAIEX（連續創高後續升）
    # ADR 折溢價（以 4/27 實際 NYSE 收盤計算）
    "adr_price":     422.10,   # TSM NYSE 4/27 實際收盤
    "adr_parity":    361.74,   # 理論值 = 5 × NT$2,265 ÷ 31.30
    "adr_ratio":     "5:1",    # 1 ADR = 5 台積電普通股
    "usd_ntd":        31.30,   # 同 MACRO 台幣匯率
    "adr_premium_pct": "+16.7",# (422.10 - 361.74) / 361.74 × 100（溢價略擴大）
    "note": "ADR 理論值 = 5 × 台股收盤 ÷ 匯率（不含交易摩擦與流動性溢價）。台美 4/27 雙市共振續創高，ADR 溢價擴大至 +16.7% 顯示美股需求更熾熱；4/28 盤前 -2.44% 暗示溢價可能收斂",
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
    "spot":              31.30,  # 今日匯率（同 MACRO，4/27 估算）
    "gm_bps_per_pct":     -40,  # 台幣每升值 1%，毛利率下降 ~40bps（法說會指引）
    "scenarios": [
        (30.0, "NTD +4.3%（強升）", "約 -172 bps", "約 -6.0%"),
        (31.0, "NTD +1.0%（小升）", "約 -40 bps",  "約 -1.4%"),
        (31.30,"基準（4/27 收盤）", "—",           "—"),
        (32.0, "NTD -2.2%（小貶）", "約 +88 bps",  "約 +3.0%"),
        (34.0, "NTD -8.6%（大貶）", "約 +345 bps", "約 +12.0%"),
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
    "analyst_view": "Q1 2026 財報確認：淨利 NT$572.5B（$18.2B，+58.3% YoY）、毛利率 66.2%（+740bps YoY）雙創歷史新高，連續第 10 季超越市場預期（平均超預期幅度 +6.2%）。Q2 指引收入 $39.0-40.2B（遠超共識 $37.5-39.0B），全年 2026 收入 >30% 確認，CapEx 向 $52-56B 高端靠齊。近期價格：台股 2330 4/27 收 NT$2,265（+80，+3.66%）續創歷史新高，盤中觸及 NT$2,330 與股號完美巧合、單日漲點 +145 創史上最大；市值盤中破 NT$60 兆、收盤 NT$58.7 兆；NYSE TSM 4/27 收 $422.10（+4.88%）續創史高；台股加權指數 4/27 站上 40,194 突破 4 萬大關；今日（4/28 Tue）TSM 盤前 -2.44% 至 $412.28 獲利了結。籌碼背離：4/27 外資反手賣超 TSMC 1.8940 萬張、全市場三大法人賣超 NT$482 億（前三名皆 ETF 調節）。今日（2026-04-28 Tue）關注：（1）開盤是否跟進 ADR 盤前 -2.44% 拉回；（2）短線指標進入極端超買（RSI 76.8/KDJ J 98.5/布林上軌突破）；（3）共識目標 NT$2,320 已達 97.6%（剩 +2.4%），新一輪上修需 Q2 7 月財報催化；（4）法律事件：前 TSMC 工程師 4/27 因 2nm 機密外洩予東京威力科創遭判 10 年，台灣首宗 2nm 等級判例護國神山法律保護升級。技術論壇後續：A13/A12 2029、N2U 2028、A16 2027 路線確認；Arizona 先進封裝廠（CoWoS+3D-IC）2029 啟用；Siemens 擴大 EDA 合作。NVIDIA #1 客戶 22% 確認，2028 先進封裝可整合 10 大晶片 +20 HBM 堆疊。31 位分析師全數看多，共識目標 NT$2,320（自 NT$2,265 剩 +2.4% 空間），Barclays 目標 $470 對應台股約 NT$2,450（剩 +8.2%）。",
    "note": "資料來源：Bloomberg / MarketScreener（估算），請每日核實更新",
}

NEWS = [
    {
        "tag": "產業趨勢", "tag_color": "#4CAF50",
        "title": "🎯 台積電 4/27 盤中觸 NT$2,330 與股號完美巧合、漲點 +145 史上最大；收 NT$2,265（+3.66%）續創高",
        "body": "台股 2330 昨日（2026-04-27 Mon）續創歷史新高收 NT$2,265（+80 點，+3.66% vs 4/24 NT$2,185），盤中一度衝至 NT$2,330 與股票代號完美巧合，單日漲點達 +145 創台股史上最大單日漲點紀錄。市值盤中突破 NT$60 兆大關（史上首次）、收盤 NT$58.7 兆雙創新高。台股加權指數同日大漲 +1,262 點收 40,194 站上 4 萬大關，台積電一檔貢獻指數約 +700 點。連 19 日台美雙市共振創高，動能延續但已進入極端區域。今日（4/28 Tue）TSM 盤前 -2.44% 至 $412.28，獲利了結湧現提示短線過熱拉回風險升高。",
        "source": "聯合新聞網 / Yahoo 股市 / CMoney 2026-04-27", "impact": "正面",
    },
    {
        "tag": "法規政策", "tag_color": "#FF9800",
        "title": "⚖️ 前 TSMC 工程師 4/27 遭判 10 年：2nm 機密外洩予東京威力科創；台灣首宗 2nm 等級判例",
        "body": "台灣智慧財產及商業法院 4/27 重判前 TSMC 工程師陳力銘 10 年徒刑（依違反國家安全法）。陳氏曾於 TSMC Fab12 良率部門任職，離職後加入日商東京威力科創（Tokyo Electron）行銷部門，於 2023 下半年至 2024 上半年期間，要求 TSMC 在職工程師吳秉駿、柯怡萍提供 2nm 製程關鍵機密技術並拍照外傳，協助東京威力科創改善蝕刻設備性能、爭取 TSMC 2nm 產線供應商資格。同案另判：吳秉駿 3 年、柯怡萍 2 年、另一名工程師陳威傑 6 年；東京威力科創員工陸怡寅 10 個月緩刑，公司被罰 NT$1.5 億並須賠 TSMC NT$1 億、繳國庫 NT$5,000 萬。此為台灣首宗 2nm 等級商業機密判例，凸顯護國神山先進製程法律保護升級。",
        "source": "Taiwan News / Taipei Times / Japan Times 2026-04-27", "impact": "中性",
    },
    {
        "tag": "客戶動態", "tag_color": "#2196F3",
        "title": "📉 4/27 籌碼背離：股價創高，外資反手賣超 TSMC 1.8940 萬張、全市場賣超 NT$482 億",
        "body": "台股 2330 4/27 雖股價續創歷史新高，三大法人卻反手大舉賣超：外資賣超 TSMC 達 1.8940 萬張（前三名賣超皆為 ETF 調節盤），全市場合計賣超 NT$482.2 億（外資 -511.96 億、投信 +135.18 億、自營 -105.42 億）。外資持股比例自 4/24 75.4% 微降至約 75.2%。投信延續加碼挺升、自營商獲利了結轉賣。籌碼面出現典型「股價創高 / 法人逢高調節」背離訊號，散戶與短線資金接棒推升。期貨未平倉量達 118,500 口（連 2 日爆量），Put/Call 比率續降至 0.51 達極端偏多反向訊號。今日（4/28 Tue）TSM 盤前 -2.44% 至 $412.28 印證籌碼疑慮，台股開盤恐承壓拉回。",
        "source": "TWSE / 經濟日報 / 聯合新聞網 2026-04-27", "impact": "負面",
    },
    {
        "tag": "客戶動態", "tag_color": "#2196F3",
        "title": "🏢 NVIDIA 確認 #1 客戶 22%（$330 億）；4/27 NVDA $961.50（+5.0%）$5T 市值在望",
        "body": "NVIDIA CEO 黃仁勳於 podcast 確認：「Nvidia is TSMC's largest customer now」——2026 NVIDIA 預估貢獻 TSMC 約 $330 億收入（佔 TSMC 總營收 22%），超越 Apple（$270 億，18%）。NVDA 4/27 續漲 +5.0% 至 $961.50，4 月累計上漲 +25%，市值逼近 $5 兆大關。NVIDIA 已預訂 TSMC 2nm 產能至 2028 年，並將成為 A16（1.6nm）首批主要客戶。Apple（#2，17%）鎖定 A14 用於 2028 年 iPhone 18 A20 晶片；AMD（#4，8%）獲法國政府 AI 戰略合作 4/27 收 $192.80（+5.8%）；Broadcom（#3，13%）4/27 收 $221.50（+5.2%）；MediaTek（#5，9%）4/27 收 NT$2,355（+3.5%）連 4 日創高。CoWoS 先進封裝 2026 產能 NVIDIA 占 ~65%。",
        "source": "CNBC / Bloomberg / IndexBox 2026-04-27", "impact": "正面",
    },
    {
        "tag": "產品發布", "tag_color": "#9C27B0",
        "title": "🛠 A13/A12/N2U 路線確認 2029 量產；Siemens-TSMC AI 晶片設計合作擴大",
        "body": "TSMC 4/22 北美技術論壇宣布最新製程路線圖延伸至 2029：(1)A13（2029 量產）較 A14 面積減少 6%、設計規則完全向後兼容；(2)A12（2029 量產）加入 Super Power Rail 背面供電，鎖定 AI/HPC，搭配第二代 nanosheet GAA；(3)N2U（2028 量產）採 DTCO 設計協同，較 N2P 提供 3-4% 高速、8-10% 低功耗；(4)A16 延至 2027；(5)Arizona 先進封裝廠（CoWoS+3D-IC）2029 啟用。Siemens 與 TSMC 4 月擴大 AI 驅動晶片設計合作，Siemens Fuse EDA AI 系統將自動化多步驟半導體設計流程，目標縮短 2nm 以下節點設計週期 30%+。論壇後續發酵推升 4/24 + 4/27 雙日創高，分析師認為此路線圖確認 TSMC 製程領先延伸至 2029-2030。",
        "source": "Tom's Hardware / EE News Europe / Siemens 2026-04-22/27", "impact": "正面",
    },
    {
        "tag": "警示", "tag_color": "#F44336",
        "title": "⚠️ 極端超買拉回警示：RSI 76.8、KDJ J 98.5、布林上軌突破；4/28 盤前 -2.44%",
        "body": "技術面紅色警示：台股 2330 4/27 收 NT$2,265 連 2 日創歷史新高，但短線指標已進入極端超買區：（1）RSI 14 估升至 76.8（嚴重超買 75+）；（2）KDJ K=92.0/D=86.5/J=98.5 三線深陷超買區（J 值 98.5 達極端超買最高警示）；（3）布林上軌已被突破。今日（4/28 Tue）TSM 盤前 -2.44% 至 $412.28 印證拉回壓力，台股開盤恐高開低走。籌碼配合：4/27 外資反手賣超 1.8940 萬張、全市場三大法人賣超 NT$482 億，逢高調節訊號明確。共識目標 NT$2,320 距 4/27 NT$2,265 僅剩 +2.4% 空間，已達 97.6%。今日重點觀察：（a）能否守住 NT$2,180-2,200 第一支撐（5MA + 心理整數）；（b）次支撐 NT$2,150；（c）NT$2,330 史高壓力短期難破；（d）量能能否維持高檔。",
        "source": "TrendForce / Yahoo Finance / Stocktwits 2026-04-28", "impact": "負面",
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
    
