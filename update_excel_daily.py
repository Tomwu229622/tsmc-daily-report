"""
TSMC Excel 每日更新腳本 - 2026-07-01
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

today_str = "2026-07-03"
tw_price = "NT$2,465（7/2收、7/3盤前未開）"   # 台股7/3 Fri盤前尚未開盤(08:51 GMT+8)；最新確認收盤為7/2 Thu NT$2,465(-40、-1.60%、守穩NT$2,450支撐)；7/3開盤恐跟進NYSE 7/2續跌承壓
change_pct = "-1.60%（7/2收）"    # 台股7/3盤前未開；顯示最新確認收盤7/2漲跌幅(-1.60%)
nyse_price = "US$434.16"       # NYSE TSM 7/2 Thu收盤(-$10.07、-2.27%、前收$444.23、自6/30史新高$477.57連二日回落、跌幅自7/1 -6.98%大幅收斂、市值約$2.25兆、P/E~37.7)；NYSE 7/3 Fri因美國國慶(7/4週六順延至7/3)休市
volume = "35,000（7/2）"        # 台股7/2全日成交量(張、估)；7/3盤前未開
news_summary = "報告日2026-07-03(Fri)。修正延續、惟跌勢趨緩＋分析師逆勢續挺：📉NYSE TSM 7/2 Thu收$434.16(-2.27%、-$10.07)自6/30史高$477.57連二日回落，惟跌幅自7/1 -6.98%大幅收斂、殺盤趨緩、屬估值修正非基本面惡化；NYSE 7/3因美國國慶休市。⚠️台股2330 7/3盤前尚未開盤、最新收7/2 NT$2,465(-1.60%)、守穩NT$2,450，7/3開盤恐跟進NYSE續跌承壓。🚀Nomura上調目標至NT$3,425(+38.9%空間)、Barclays亦上調；平均目標$476.24(自$434.16具+9.7%空間)、拉回視為估值修正；⚠️惟Goldman 7/1移出APAC Conviction List(戰術調整非降評)。🤝NVIDIA-TSMC導入AI入廠提升良率；BofA上修2026全球半導體市場至$1.3兆(原$1.0兆)、2030上看$2兆。⭐漲價續發酵、全先進製程調漲5-10%(佔晶圓營收~74%)、估全年毛利率+2pp以上。中長線基底未變：NVIDIA #1客戶(22%/$33B)/A16首發、Vera Rubin 6顆晶片全TSMC代工(3nm+HBM4)、2nm 70-80%良率、費半Q2 +87.8%史上最強單季。⚠️估值警示TSM P/E~37.7x、台股~36.3x；6月營收(7月上旬)與7/16財報為關鍵catalyst；下方支撐NT$2,450/NT$2,410、上方壓力NT$2,505/NT$2,535史高。"
change_color = "FF0000"        # 紅色（NYSE 7/2 續跌 -2.27%、台股 7/2 收盤下跌 -1.60%、7/3 盤前未開）

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
