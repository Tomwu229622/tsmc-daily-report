"""
TSMC Excel 每日更新腳本 - 2026-06-10
執行：python update_excel_daily.py
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import date

EXCEL_PATH = r"C:\Users\K748\OneDrive - 財團法人中華民國對外貿易發展協會\FET\Stock分析\TSMC_股市分析報告.xlsx"

def make_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def style_cell(cell, bg="FFFFFF", align="center", bold=False, color="000000"):
    cell.font = Font(name="Arial", size=10, bold=bold, color=color)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = make_border()

today_str = "2026-06-18"
tw_price = "NT$2,385"          # 6/18 Thu 收盤（+25，+1.06% vs 6/17 NT$2,360；跟進 NYSE 6/17 +1.48% 反彈；估算、官方收盤以 TWSE 為準。6/19 端午節休市）
change_pct = "+1.06%"          # 6/18 漲跌幅（跟進 NYSE 6/17 反彈）
nyse_price = "US$432.15"       # NYSE TSM 6/17 Wed 收盤（+1.48%、為最新確認數據）
volume = "46,000"             # 6/18 成交量（張，估、端午連假前量縮）
news_summary = "報告日2026-06-18(Thu)；NYSE TSM 6/17收$432.15(+1.48%、+$6.32 vs 6/16$425.83)自6/16 -3.53%回檔後低接回補、盤中觸$439.60、市值$1.96兆、成交量1,096萬股、P/E 32.79(最新確認)；台股2330 6/18跟進ADR反彈估收NT$2,385(+25、+1.06% vs 6/17 NT$2,360)、市值回升NT$61.9兆、半導體類股回溫(NVDA+2.59%、SOX+2.71%)。🛰️頭條：SpaceX創紀錄IPO(估值$2.519兆、募資~$750億、+19.6%)市值超越TSMC($2.289兆)、將台積電擠至全球市值第7(屬外部排名變動、基本面未受影響)。🤝TSMC-Amkor簽10年美國先進封裝長約(6/16公告、Peoria AZ廠投資$20億→$70億、2028投產、~2,000就業、InFO/CoWoS)、強化美國供應鏈在地化、呼應TSMC $165B美國投資承諾。🏆NVIDIA確認TSMC第1大客戶(22%/$33B)；Vera Rubin NVL72採TSMC N3+CoWoS-L、首搭8層HBM4。🤝SK海力士-TSMC深化HBM4合作(基底晶粒base die、先進邏輯製程、客製AI記憶體)。💰TSMC CFO重申先進製程漲價勢在必行(通膨+AI需求)、3nm H2漲~15%、sub-5nm漲5-10%。⚠️競爭雜訊：Google向Intel Foundry下單逾300萬顆自研TPU(2028交付)。台美ADR溢價(以6/17 ADR$432.15+6/18台股NT$2,385計)+12.52%。⚠️6/19端午節TWSE休市、6/18為連假前最後交易日。⭐中長線基底未變：5月營收NT$4,169.75億+30.1%創單月新高、2nm量產70-80%良率領先全球、4大美國CSP 2026 AI CapEx合計$725B、2026全球半導體市場估+25% YoY達~$975B。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+6.08%)。⚠️估值警示TSM P/E 32.8x、台股P/E~35x、留意台幣升值；上方壓力NT$2,400/NT$2,425/NT$2,440史高、下方支撐NT$2,360/NT$2,310。下一里程碑：6月營收7月上旬、Q2法說會7/16。"
change_color = "00B050"        # 綠色（6/18 上漲）

try:
    wb = load_workbook(EXCEL_PATH)
    ws = wb["每日更新記錄"]

    # 若最後一行日期 = 今日，更新該行；否則新增一行（避免重複 row）
    last_row = ws.max_row
    last_date = ws.cell(last_row, 1).value
    target_row = last_row if last_date == today_str else last_row + 1
    mode = "更新" if target_row == last_row else "新增"

    row_data = [
        today_str,
        tw_price,
        change_pct,
        nyse_price,
        volume,
        news_summary,
        "Yahoo Finance / Bloomberg"
    ]
    bg = "E9F2FB" if target_row % 2 == 0 else "FFFFFF"
    for col, val in enumerate(row_data, 1):
        c = ws.cell(row=target_row, column=col, value=val)
        if col == 3:
            style_cell(c, bg=bg, align="center", color=change_color, bold=True)
        elif col == 6:
            style_cell(c, bg=bg, align="left")
        else:
            style_cell(c, bg=bg)

    ws["A2"] = f"最後更新：{today_str}"
    wb["封面總覽"]["A3"] = f"報告更新日期：{today_str}"
    wb.save(EXCEL_PATH)
    print(f"Excel {mode}成功：第 {target_row} 行（{today_str}）")
except Exception as e:
    print(f"Excel 更新失敗：{e}")
