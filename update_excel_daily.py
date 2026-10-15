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

today_str = '2026-10-15'
tw_price = 'NT$2,490（10/14收，平盤0.00、0.00%；開2,505高2,535低2,490，開高後收在全日最低、黑K，盤中2,535未守住）；10/13收NT$2,490（-1.97%）'
change_pct = '0.00%（10/14收，平盤，相對大盤-0.59pp；加權+284.52點、+0.59%收48,625.83、結束連4日下跌）；10/13為-1.97%'
nyse_price = 'US$430.00（10/14收，-3.40%；9/16以來最低收盤、五日-8.94%；ADR溢價壓縮至+10.76%）；10/13收US$445.12（+0.17%）'
volume = '28,112（10/14收，-31.1%；高於前報基準25,048張12.2%、高於新5日均量26,510張6.0%；成交筆數91,202筆；2026年188日中第133）；10/13為40,796張'
news_summary = """報告日2026-10-15（台美皆為10/14）。10/14開高後收全日最低2,490（平盤），外資賣超6,752張（連6賣），立場維持看空；MACD死叉；美股CPI偏熱、費半-4.30%，ADR-3.40%收430.00、溢價+10.76%，夜盤個股期2,474；Q3法說會今日14:00。"""
change_color = '000000'
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
