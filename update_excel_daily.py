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

today_str = '2026-11-04'
tw_price = 'NT$2,295（11/03收，-35、-1.50%；開低2,300、高2,335、低2,280，實體5元黑K；7/30以來最低收盤、連4日下跌；續在半年線MA120 2,390.04與五條均線之下，連11日低於MA60 2,430.5（差135.5元）；MA10下穿MA120；2,345–2,360缺口第2日、2,380–2,390缺口第3日、2,450–2,490缺口第12日未回補）；11/02收NT$2,330（-1.69%）'
change_pct = '-1.50%（11/03收，相對大盤+2.09pp（精算+2.0878）；加權-1,612.60點、-3.59%收43,307.54，全市場成交10,020.36億）；11/02為-1.69%'
nyse_price = 'US$404.80（11/03收，+1.55%；終結連3日下跌；ADR溢價同日對齊+15.57%（ADR 11/03對台股11/03與匯率32.76），11/02同日對齊+11.51%；盤後約404.50（19:59 EST））；11/02收US$398.63（-1.85%）'
volume = '46,137（11/03收，+12.08%；高於前報基準31,245張47.66%、新5日均量32,837.8張；2026年200日中第47）；11/02為41,165張'
news_summary = """11/03收2,295(-1.50%)創7/30來低、相對大盤抗跌；外資連18賣11,703張，維持看空；ADR+1.55%。"""
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
