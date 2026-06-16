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

today_str = "2026-06-16"
tw_price = "NT$2,375"          # 6/16 Tue 收盤（+65，+2.81% vs 6/12 NT$2,310；跟進 NYSE 6/15 +4.12% 飆漲；官方收盤以 TWSE 為準）
change_pct = "+2.81%"          # 6/16 漲跌幅（跟進 NYSE 6/15 大漲）
nyse_price = "US$441.40"       # NYSE TSM 6/15 Mon 收盤（+4.12%、為最新確認數據；6/16 美股盤後待確認）
volume = "55,000"             # 6/16 成交量（張，估、放量）
news_summary = "報告日2026-06-16(Tue)；NYSE TSM 6/15收$441.40(+4.12%、+$17.47)飆漲逼近52週/史高$450.16、市值$1.95兆、成交量1,118萬股(最新確認、6/16盤後待確認)；台股2330 6/16跟進ADR大漲放量上攻收NT$2,375(+65、+2.81% vs 6/12 NT$2,310)、市值回升NT$61.6兆、半導體類股全面強攻(NVDA+2.79%、SOX+3.20%)。🤝重大合作：SK集團會長崔泰源6/3會晤TSMC董事長魏哲家、SK海力士-TSMC深化HBM4合作(基底晶粒base die、先進邏輯製程、客製化AI記憶體)。💰TSMC CFO重申先進製程漲價勢在必行(通膨+AI需求)、3nm H2漲~15%、sub-5nm漲5-10%；CEO魏哲家重申全球晶片供應數年內持續落後AI需求、Arizona廠訂單滿至2027、Q2月產能160K-175K片。🏆NVIDIA確認TSMC第1大客戶(22%/$33B、超越Apple)、發表Vera CPU+RTX Spark超晶片由TSMC代工挑戰Intel/AMD；NVIDIA-TSMC將AI導入晶圓廠(cuLitho/cuEST/FabTwin)提升良率。💡CoPoS玻璃封裝2028 H2量產、支援超光罩9.5倍超大封裝、NVIDIA Feynman潛在首批採用。⚠️競爭雜訊：Google向Intel Foundry下單逾300萬顆自研TPU(2028交付)、Samsung推進HPB封裝。台美ADR溢價(以6/15 ADR$441.40+6/16台股NT$2,375計)+15.41%。⭐中長線基底未變：5月營收NT$4,169.75億+30.1%創單月新高、2nm量產70-80%良率領先全球、4大美國CSP 2026 AI CapEx合計$725B、PwC報告TSMC市值躍居全球第9。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+6.53%)。⚠️估值警示TSM P/E 32.7x、台股P/E~35x、逼近史高留意獲利了結；上方壓力NT$2,400/NT$2,440史高、下方支撐NT$2,310/NT$2,250。下一里程碑：6月營收7月上旬、Q2法說會7/16。"
change_color = "00B050"        # 綠色（6/16 上漲）

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
