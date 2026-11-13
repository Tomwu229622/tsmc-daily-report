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

today_str = '2026-11-13'
tw_price = 'NT$2,340（11/12收，-55、-2.30%；開2,385（低開10元）、高2,395、低2,330，黑K實體45元；2026年96個下跌日第14大，非極端；終止連3日上漲；跌回MA20 2,355.25與半年線MA120 2,391.71，MA60 2,415.92續在其下（連18日），守住MA10 2,288.50與MA5 2,278.00；2,285–2,305向上缺口未回補、無新缺口；MACD柱+1.19翻正、KD黃金交叉第5日、RSI 47.73跌回50之下；外資轉賣3,955張（門檻內）；轉弱程度由明顯緩和略為回落為緩和）'
change_pct = '-2.30%（11/12收，相對大盤+0.34pp（精算+0.3365），延續相對抗跌；加權-1,165.28點、-2.63%收43,091.87，全市場成交10,568.42億）；11/11為+4.81%'
nyse_price = 'US$405.49（11/12收，-1.90%，2年233個下跌日並列第84大，非極端；終止連4日上漲；費半-0.54%；ADR溢價同日對齊+13.99%（ADR 11/12對台股11/12與匯率32.89），11/11同日對齊+13.43%；盤後約409.0（+0.88%，19:55 EST））；11/11收US$413.35'
volume = '34,359（11/12收，+3.94%；為前報基準54,281張的0.63倍（-36.70%）＝量縮下跌、新5日均量52,117.8張；2026年207日中第107大）；11/11為33,056張'
news_summary = """11/12收2,340(-2.30%)終止連3漲、跌回半年線但MACD翻正；外資轉賣3,955張；輝達宣布黃仁勳辭任；維持看空。"""
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
