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

today_str = '2026-11-18'
tw_price = 'NT$2,170（11/17收，-45、-2.03%；開2,195（低開20元）、高2,270、低2,155，黑K實體25元、上影線75元；2026年98個下跌日並列第25大；連2日下跌，收盤與盤中最低皆為11/06以來最低；續在MA5 2,294、MA10 2,262.50、MA20 2,324.50、半年線MA120 2,392.25（連4日）與MA60 2,408.58（連21日）之下；無新缺口、守在布林下軌2,149.09之上；MACD柱-13.80續負、KD死亡交叉、RSI 36.16；外資賣超14,606張（逾門檻）、投信轉買2,275張、三大合計-5,506張；放量下跌；轉弱程度加劇）'
change_pct = '-2.03%（11/17收，相對大盤+0.28pp（精算+0.2783），延續相對抗跌；加權-947.97點、-2.31%收40,092.11，全市場成交10,128.68億）；11/16為-5.74%'
nyse_price = 'US$359.69（11/17收，-3.68%，2年235個下跌日第29大；連4日下跌；費半+0.34%；ADR溢價同日對齊+8.90%（ADR 11/17對台股11/17與匯率32.85），11/16同日對齊+10.90%；輝達財報後盤後約376（約+4.6%，19:59 EST））；11/16收US$373.45'
volume = '61,611（11/17收，+17.12%；為前報基準37,514張的1.64倍（+64.24%）＝放量下跌、新5日均量42,520.8張；2026年210日中第14大）；11/16為52,605張'
news_summary = """11/17收2,170(-2.03%)續跌、KD死叉；外資賣超14,606張逾門檻、投信轉買；輝達Q3優於共識、ADR-3.68%盤後約376；維持看空。"""
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
