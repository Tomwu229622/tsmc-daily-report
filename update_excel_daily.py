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

today_str = '2026-10-19'
tw_price = 'NT$2,420（10/16收，-20、-0.82%；開2,430高2,445低2,420，開低走低收最低小黑K，2,450–2,490缺口未回補，收盤低於MA60 2,421僅1元）；10/15收NT$2,440（-2.01%）'
change_pct = '-0.82%（10/16收，相對大盤-1.12pp；加權+141.39點、+0.30%收47,111.26，全市場成交9,938.15億）；10/15為-2.01%'
nyse_price = 'US$412.61（10/16收，-2.28%；連3跌、8/24以來最低、五日-8.98%；ADR溢價收斂至+10.28%）；10/15收US$422.25（-1.80%）'
volume = '43,047（10/16收，-26.4%；高於前報基準34,819張23.6%、高於新5日均量38,800張10.9%；2026年190日中第57）；10/15為58,487張'
news_summary = """10/16法說後首日收最低2,420（-0.82%），外資賣超22,046張連8賣、低於MA60 1元，維持看空；費半-3.83%、ADR-2.28%，夜盤偏弱。"""
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
