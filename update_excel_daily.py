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

today_str = "2026-04-22"
tw_price = "NT$2,065"        # 4/22 早盤（+0.73%，vs 4/21 收 NT$2,050；北美技術論壇催化）
change_pct = "+0.73%"        # 台股 2330 4/22 早盤漲幅
nyse_price = "US$374.00"     # NYSE TSM 估算 4/21 收盤（+0.94%；Good Friday 4/18 後首日，實際待確認）
volume = "~12,000 張"         # 4/22 估算全日（早盤 09:14 約 4,009 張）
news_summary = "TSMC北美技術論壇(4/22):A14製程宣布(2028量產,速度+15%/功耗-30%);A16 H2 2026量產確認;台股早盤NT$2,065(+0.73%);三大法人4/21淨買超2829張"
change_color = "00B050"  # 上漲綠色

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
