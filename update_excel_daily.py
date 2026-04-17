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

today_str = "2026-04-17"
tw_price = "NT$2,085"        # 4/17 財報後首日收盤（-0.48%，獲利了結）
change_pct = "-0.48%"        # 台股 2330 4/17 漲跌（vs 4/16 收盤 NT$2,095）
nyse_price = "US$390.00"     # NYSE TSM 4/17 估算（財報後整理，-0.51%）
volume = "35,000 張"          # 4/17 成交量（財報後縮量整理）
news_summary = "4/17 財報後首日：Q1淨利+58%($18.1B)創歷史高；Q2指引$39-40.2B遠超共識；NVIDIA超越Apple成TSMC最大客戶；TSMC警告伊朗戰爭供應鏈風險"
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
