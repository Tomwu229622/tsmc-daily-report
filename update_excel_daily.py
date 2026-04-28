"""
TSMC Excel 每日更新腳本 - 2026-03-19
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

today_str = "2026-04-28"
tw_price = "NT$2,265"        # 4/27 收盤（+80 vs 4/24 NT$2,185，+3.66%；盤中觸 NT$2,330 史上最大漲點）
change_pct = "+3.66%"        # 台股 2330 4/27 收盤漲跌幅
nyse_price = "US$422.10"     # NYSE TSM 4/27 收盤（+4.88% 續創史高；4/28 盤前 -2.44%）
volume = "62,800 張"         # 4/27 成交量（量能再放大 +40%）
news_summary = "台股2330 4/27收NT$2,265(+3.66%,+80)續創史高,盤中觸NT$2,330與股號神奇巧合、單日漲點+145史上最大;市值盤中破NT$60兆(收盤NT$58.7兆);加權指數站上40,194破4萬;NYSE TSM 4/27 $422.10(+4.88%)續創史高;籌碼背離:外資反手賣超TSMC 1.8940萬張、全市場三大法人賣超482億;前TSMC工程師陳力銘因2nm機密外洩遭判10年,東京威力科創罰NT$1.5億;4/28盤前TSM -2.44%獲利了結;極端超買警示RSI 76.8/KDJ J 98.5"
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
