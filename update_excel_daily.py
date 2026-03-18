"""
TSMC Excel 每日更新腳本 - 2026-03-18
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

today_str = "2026-03-18"
tw_price = "NT$1,870"        # 3/17 收盤（3/18 數據確認中）
change_pct = "-5.03%"        # NYSE TSM 3/18 漲跌（台股數據待確認）
nyse_price = "US$336.71"     # NYSE TSM 3/18 收盤
volume = "23,688 張"          # 3/17 成交量（3/18 待確認）
news_summary = "NYSE TSM 3/18大跌-5.03%至$336.71，半導體板塊廣泛修正；NVIDIA Feynman確認採TSMC A16；記憶體短缺延伸至2027"
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
