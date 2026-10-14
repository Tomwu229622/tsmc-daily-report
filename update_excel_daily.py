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

today_str = '2026-10-14'
tw_price = 'NT$2,490（10/13收，-50、-1.97%；開2,500高2,510低2,490，開低40元收在全日最低、小黑K，跌破10/12低點2,510與MA10）；10/12收NT$2,540（-0.39%）'
change_pct = '-1.97%（10/13收，相對大盤-0.02pp；加權-960.65點、-1.95%收48,341.31、連4日下跌）；10/12為-0.39%'
nyse_price = 'US$445.12（10/13收，+0.17%；結束連5日下跌、五日-7.71%；ADR溢價擴大至+14.73%）；10/12收US$444.35（-1.98%）'
volume = '40,796（10/13收，+73.2%，9/18以來最大；高於前報基準22,249張83.4%、高於新5日均量25,048張62.9%；成交筆數329,150筆；2026年187日中第66）；10/12為23,556張'
news_summary = """報告日2026-10-14（台美皆為10/13）。10/13放量收最低2,490（-1.97%），外資賣超25,641張（2026年第6大、連5賣）且跌破2,510，立場依約自中性偏空下修為看空；KD中期轉弱確認；台股收盤後ADR+0.17%止跌、溢價+14.73%，夜盤個股期2,520高於現貨30點；10/15法說會距今1天。"""
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
