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

today_str = '2026-11-05'
tw_price = 'NT$2,260（11/04收，-35、-1.53%；開高2,320即全日最高、低2,235，實體60元黑K；收盤與盤中低點皆7/30以來最低、連5日下跌；續在半年線MA120 2,390.25與五條均線之下，連12日低於MA60 2,428.08（差168.08元）；2,345–2,360缺口第3日、2,380–2,390缺口第4日、2,450–2,490缺口第13日未補；RSI 27.28跌破30再創2026年新低；外資賣超21,603張連19賣；轉弱程度再加深）'
change_pct = '-1.53%（11/04收，相對大盤-0.70pp（精算-0.697）；加權-358.63點、-0.83%收42,948.91，全市場成交11,769.30億）；11/03為-1.50%'
nyse_price = 'US$395.75（11/04收，-2.24%；反彈1日後再跌；ADR溢價同日對齊+14.94%（ADR 11/04對台股11/04與匯率32.82），11/03同日對齊+15.57%；盤後約391.6–391.7（19:59 EST））；11/03收US$404.80（+1.55%）'
volume = '53,879（11/04收，+16.78%；為前報基準32,838張的1.64倍（+64.08%）、新5日均量39,240.2張；2026年201日中第27）；11/03為46,137張'
news_summary = """11/04收2,260(-1.53%)創7/30來低、落後大盤；外資連19賣21,603張，維持看空；ADR-2.24%。"""
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
