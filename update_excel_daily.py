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

today_str = '2026-11-17'
tw_price = 'NT$2,215（11/16收，-135、-5.74%；開2,340（低開10元）、高2,350、低2,215，收在全日最低，黑K實體125元；2026年97個下跌日第2大；一日內跌回MA5 2,317、MA10 2,275、MA20 2,337之下，收盤低於全部五條均線，半年線MA120 2,392.96（連3日）與MA60 2,412.08（連20日）續在其下；回補2,285–2,305向上缺口、無新缺口；MACD柱-1.69重新轉負、RSI 38.79跌回40下、KD黃金交叉第7日；外資賣超8,466張（門檻內）、投信轉賣終止連31買、三大合計-7,684張；放量下跌；轉弱程度由緩和轉為加劇）'
change_pct = '-5.74%（11/16收，相對大盤+0.29pp（精算+0.2886），恢復相對抗跌；加權-2,635.04點、-6.03%收41,040.08，全市場成交9,031.15億）；11/13為+0.43%'
nyse_price = 'US$373.45（11/16收，-6.34%，2年234個下跌日第10大；連3日下跌；費半-5.28%；ADR溢價同日對齊+10.90%（ADR 11/16對台股11/16與匯率32.89），11/13同日對齊+11.44%；盤後約373（約持平，19:59 EST））；11/13收US$398.73'
volume = '52,605（11/16收，+69.84%；為前報基準36,169張的1.45倍（+45.44%）＝放量下跌、新5日均量37,513.8張；2026年209日中第32大）；11/13為30,973張'
news_summary = """11/16收2,215(-5.74%)長黑跌回MA5/10/20、MACD轉負；外資賣超8,466張、投信轉賣；費半-5.28%、ADR-6.34%；維持看空。"""
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
