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

today_str = '2026-11-03'
tw_price = 'NT$2,330（11/02收，-40、-1.69%；跳空開低2,310即全日最低、高2,345，實體20元紅K；8/4以來最低收盤；續在半年線MA120 2,390.00與五條均線之下，連10日低於MA60 2,430.92（差100.92元）；再留2,345–2,360向下缺口，2,380–2,390缺口第2日、2,450–2,490缺口第11日未回補）；10/30收NT$2,370（-1.86%）'
change_pct = '-1.69%（11/02收，相對大盤-1.04pp（精算-1.0426）；加權-291.69點、-0.65%收44,920.14，全市場成交7,996.88億）；10/30為-1.86%'
nyse_price = 'US$398.63（11/02收，-1.85%；連3日下跌；ADR溢價同日對齊+11.51%（ADR 11/02對台股11/02與匯率32.59），10/30同日對齊+11.32%；盤後約399.3–399.8（19:59 EST））；10/30收US$406.13（-5.95%）'
volume = '41,165（11/02收，+12.42%；高於前報基準28,971張42.09%、新5日均量31,245張；2026年199日中第65）；10/30為36,616張'
news_summary = """11/02收2,330跌1.69%創8/4來低；外資賣超15,005張連17賣，KD死叉，維持看空；費半-3.37%、ADR-1.85%。"""
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
