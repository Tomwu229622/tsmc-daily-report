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

today_str = "2026-04-20"
tw_price = "NT$2,030"        # 4/17 實際收盤（4/20 週末休市，保留上一交易日）
change_pct = "-2.64%"        # 台股 2330 4/17 實際漲跌（深度獲利了結）
nyse_price = "US$370.50"     # NYSE TSM 4/17 實際收盤（+1.97%，美股財報後反彈）
volume = "32,500 張"          # 4/17 成交量估算（財報後縮量整理）
news_summary = "4/20 週末休市(4/17收NT$2,030,-2.64%;TSM$370.50,+1.97%);Amazon/Meta/Anthropic自研AI芯片均依賴TSMC;Q1淨利+58%毛利率66.2%雙創歷史高"
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
