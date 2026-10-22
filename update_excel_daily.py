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

today_str = '2026-10-21'
tw_price = 'NT$2,405（10/20收，-15、-0.62%；開2,400高2,415低2,395，開低20元後小紅K，9/16以來最低收盤；2,450–2,490缺口第4日未回補，收盤連3日低於MA60且差16.5元＝中期轉弱有效）；10/19收NT$2,420（平盤）'
change_pct = '-0.62%（10/20收，相對大盤+0.51pp；加權-529.15點、-1.13%收46,292.93，全市場成交8,029.47億）；10/19為0.00%'
nyse_price = 'US$425.08（10/20收，+1.16%；連2日上漲、五日-4.50%；ADR溢價擴大至+15.31%，新台幣收32.617、取32.62）；10/19收US$420.22（+1.84%）'
volume = '22,402（10/20收，-19.1%；低於前報基準39,627張43.5%、低於新5日均量35,949張37.7%；10/7以來最低，2026年192日中第161）；10/19為27,695張'
news_summary = """10/20收2,405跌0.62%，外資連10賣，連3日破季線轉弱有效、RSI破40，維持看空；收盤後高雄F22工安4死；ADR+1.16%。"""
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
