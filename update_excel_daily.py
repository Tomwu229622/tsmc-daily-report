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

today_str = '2026-11-06'
tw_price = 'NT$2,240（11/05收，-20、-0.88%；開低2,250、高2,285、低2,235（與11/04持平），實體10元黑K；7/30以來最低收盤、連6日下跌；續在半年線MA120 2,390.13與五條均線之下、MA20下穿MA60，連13日低於MA60 2,426.00（差186元）；2,345–2,360缺口第4日、2,380–2,390缺口第5日、2,450–2,490缺口第14日未補；RSI 25.76再創2026年新低；外資賣超10,252張連20賣（較前日收斂11,351張）；轉弱程度再加深）'
change_pct = '-0.88%（11/05收，相對大盤+2.31pp（精算+2.3146），恢復相對抗跌；加權-1,374.19點、-3.20%收41,574.72，全市場成交9,870.99億）；11/04為-1.53%'
nyse_price = 'US$353.10（11/05收，-10.78%，2年235個下跌日第2大；費半-12.25%為2年最大跌幅；ADR溢價同日對齊+3.66%（ADR 11/05對台股11/05與匯率32.88），11/04同日對齊+14.94%；盤後約360.50（19:59 EST））；11/04收US$395.75（-2.24%）'
volume = '45,175（11/05收，-16.15%；為前報基準39,240張的1.15倍（+15.12%）、新5日均量44,594.4張；2026年202日中第53）；11/04為53,879張'
news_summary = """11/05收2,240(-0.88%)7/30來低但抗跌；外資連20賣10,252張；費半-12.25%，維持看空。"""
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
