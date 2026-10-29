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

today_str = '2026-10-29'
tw_price = 'NT$2,425（10/28收，+25、+1.04%；開高2,420、低2,410高於前日高2,405留下向上跳空缺口、高2,430，實體5元小紅K；站上MA5 2,407（結束連11日在下），但仍連7日低於MA60 2,429.00（差4元）、連10日低於MA10；2,450–2,490缺口第8日未回補）；10/27收NT$2,400（+0.63%）'
change_pct = '+1.04%（10/28收，相對大盤-1.00pp（精算-0.9961）；加權+939.35點、+2.04%收47,035.58，全市場成交7,503.41億）；10/27為+0.63%'
nyse_price = 'US$439.51（10/28收，+0.60%；連4日上漲；ADR溢價同日對齊+17.52%（ADR 10/28對台股10/28與匯率32.42），10/27同日對齊+18.14%；盤後431.09（-1.92%，原因未查證））；10/27收US$436.87（+2.95%）'
volume = '21,867（10/28收，-42.72%；低於前報基準29,032張24.7%、新5日均量27,867張；2026年196日中第166）；10/27為38,173張'
news_summary = """10/28收2,425漲1.04%站上MA5，但量縮、落後大盤1.00pp；外資賣超3,175張連14賣，未回MA60、維持看空；ADR+0.60%、溢價+17.52%。"""
change_color = '00B050'
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
