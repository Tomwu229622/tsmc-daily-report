"""
TSMC Excel 每日更新腳本 - 2026-05-05
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

today_str = "2026-05-05"
tw_price = "NT$2,275"        # 5/4 Mon 收盤（+140 vs 4/30 NT$2,135，+6.56%；爆量飆漲創 52 週新高、單日漲幅 2024 以來最大）
change_pct = "+6.56%"        # 台股 2330 5/4 收盤漲跌幅
nyse_price = "US$401.61"     # NYSE TSM 5/4 收盤（+0.99%；連 4 反彈、收復 $400 關卡）
volume = "92,800 張"         # 5/4 成交量（爆量近 2 倍 vs 4/30 48,745 張）
news_summary = "台股2330 5/4爆量飆漲收NT$2,275(+6.56%,+140)創52週新高、單日漲幅2024以來最大;台股加權指數同日+1,778.51(+4.57%)收40,705.14史上首度收盤站穩4萬大關;NYSE TSM 5/4 $401.61(+0.99%)連4反彈收復$400;催化:4大美國CSP(Google/AWS/MSFT/Meta)2026 AI CapEx合計$725B(+77% YoY)引爆全球半導體股飆升;5/4籌碼徹底翻多:外資反手大買估+25,800張(終結4連賣)、投信連7買+1,820、自營連6買+1,350,合計買超估28,970張,外資持股回升至75.1%;AI晶片股延續強勢:NVDA $952.80(+1.34%)、AMD +0.97%、AVGO +1.72%、SOX 10,795(+2.42%);Goldman Sachs維持目標NT$2,330、Barclays $470、共識NT$2,320(剩+2.0%);AMD Q1財報今日(5/5)盤後;TSMC 4月月營收5/10前公布(估NT$330B+);Q1基本面強勁不變"
change_color = "00B050"  # 上漲綠色

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
