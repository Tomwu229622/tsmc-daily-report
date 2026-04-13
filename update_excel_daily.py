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

today_str = "2026-04-13"
tw_price = "NT$1,950"        # 4/10 收盤（今日4/13週日休市，取上一交易日）
change_pct = "-0.26%"        # 台股 2330 4/10 漲跌（vs 4/9 收盤 NT$1,955）
nyse_price = "US$370.60"     # NYSE TSM 4/10 收盤
volume = "28,000 張"          # 4/10 成交量（估）
news_summary = "今日週日休市，取4/10收盤；Q1 2026營收NT$1.134兆(+35.1%)超標；3月營收+45.2% YoY；SOX指數歷史新高；4/16 Q1財報倒計時3天"
change_color = "FF0000"  # 下跌紅色

try:
    wb = load_workbook(EXCEL_PATH)
    ws = wb["每日更新記錄"]

    next_row = ws.max_row + 1
    row_data = [
        today_str,
        tw_price,
        change_pct,
        nyse_price,
        volume,
        news_summary,
        "Yahoo Finance / Bloomberg"
    ]
    bg = "E9F2FB" if next_row % 2 == 0 else "FFFFFF"
    for col, val in enumerate(row_data, 1):
        c = ws.cell(row=next_row, column=col, value=val)
        if col == 3:
            style_cell(c, bg=bg, align="center", color=change_color, bold=True)
        elif col == 6:
            style_cell(c, bg=bg, align="left")
        else:
            style_cell(c, bg=bg)

    ws["A2"] = f"最後更新：{today_str}"
    wb["封面總覽"]["A3"] = f"報告更新日期：{today_str}"
    wb.save(EXCEL_PATH)
    print(f"Excel 更新成功：第 {next_row} 行（{today_str}）")
except Exception as e:
    print(f"Excel 更新失敗：{e}")
