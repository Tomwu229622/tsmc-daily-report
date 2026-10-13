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

today_str = '2026-10-13'
tw_price = 'NT$2,540（10/12收，-10、-0.39%；開2,510高2,550低2,510，開低40元回補10/5缺口後拉回、紅K但跌破2,545防線5元；10/9國慶補假休市後首日）；10/8收NT$2,550（-1.35%）'
change_pct = '-0.39%（10/12收，相對大盤-0.37pp；加權-11.48點、-0.02%收49,301.96、連3日下跌）；10/8為-1.35%'
nyse_price = 'US$444.35（10/12收，-1.98%；連5日下跌、五日-8.53%、9/18以來最低；ADR溢價壓縮至+11.61%）；10/9收US$453.31（-1.02%）'
volume = '23,556（10/12收，+1.8%，高於前報基準20,696張13.8%、高於新5日均量22,249張5.9%；成交筆數125,197筆；2026年186日中第153）；10/8為23,145張'
news_summary = """報告日2026-10-13（台美皆為10/12）。10/12開低回補10/5缺口後收2,540（-0.39%），跌破防線2,545達5元，立場依約自中性下修為中性偏空；外資連4日賣超合計24,559張；台股收盤後ADR再-1.98%連5跌收444.35、溢價+11.61%，夜盤個股期2,520低於現貨20點；10/15法說會距今2天。"""
change_color = 'FF0000'
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
