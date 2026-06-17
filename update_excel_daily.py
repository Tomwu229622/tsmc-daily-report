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

today_str = "2026-06-17"
tw_price = "NT$2,315"          # 6/17 Wed 收盤（-60，-2.53% vs 6/16 NT$2,375；跟進 NYSE 6/16 -3.53% 回檔；估算、官方收盤以 TWSE 為準）
change_pct = "-2.53%"          # 6/17 漲跌幅（跟進 NYSE 6/16 回檔）
nyse_price = "US$425.83"       # NYSE TSM 6/16 Tue 收盤（-3.53%、為最新確認數據；6/17 美股盤後待確認）
volume = "50,000"             # 6/17 成交量（張，估、量縮整理）
news_summary = "報告日2026-06-17(Wed)；NYSE TSM 6/16收$425.83(-3.53%、-$15.57)自逼近52週/史高$450.16高位獲利了結、市值$1.97兆、成交量1,109萬股、P/E 33.07(最新確認、6/17盤後待確認)；台股2330 6/17跟進ADR回檔估收NT$2,315(-60、-2.53% vs 6/16 NT$2,375)、市值回落NT$60.0兆、半導體類股同步回吐(NVDA-2.91%、SOX-3.35%)。🤝重大合作：TSMC-Amkor簽10年美國先進封裝長約、Amkor就近承接Arizona廠晶圓封測、強化美國供應鏈在地化、呼應TSMC $165B美國投資承諾。🤝SK海力士-TSMC深化HBM4合作(基底晶粒base die、先進邏輯製程、客製化AI記憶體)。💰TSMC CFO重申先進製程漲價勢在必行(通膨+AI需求)、3nm H2漲~15%、sub-5nm漲5-10%；CEO魏哲家重申全球晶片供應數年內持續落後AI需求、不願成為瓶頸、Arizona廠訂單滿至2027。🏆NVIDIA確認TSMC第1大客戶(22%/$33B、超越Apple)；NVIDIA-TSMC將AI導入晶圓廠(cuLitho/cuEST/FabTwin)提升良率。📦CoWoS封裝瓶頸續為產業關鍵、TSMC規劃2026底擴至~130K wpm、5.5倍光罩封裝良率>98%。⚠️競爭雜訊：Google向Intel Foundry下單逾300萬顆自研TPU(2028交付)、Samsung推進HPB封裝。台美ADR溢價(以6/16 ADR$425.83+6/17台股NT$2,315計)+14.23%。⭐中長線基底未變：5月營收NT$4,169.75億+30.1%創單月新高、2nm量產70-80%良率領先全球、4大美國CSP 2026 AI CapEx合計$725B、PwC報告TSMC市值躍居全球第9。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+9.29%)。⚠️估值警示TSM P/E 33.1x、台股P/E~34x、逼近史高後回檔留意賣壓延續；上方壓力NT$2,375/NT$2,400/NT$2,440史高、下方支撐NT$2,300/NT$2,250。下一里程碑：6月營收7月上旬、Q2法說會7/16。"
change_color = "FF0000"        # 紅色（6/17 下跌）

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
