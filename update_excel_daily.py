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

today_str = '2026-11-16'
tw_price = 'NT$2,350（11/13收，+10、+0.43%；開2,375（高開35元）、高2,390、低2,340，黑K實體25元；2026年100個上漲日第87大，非極端；以1.75元之差收回MA20 2,348.25、MA5上穿MA10，半年線MA120 2,393.08（連2日）與MA60 2,415.17（連19日）續在其下，守住MA10 2,286.50與MA5 2,321.00；2,285–2,305向上缺口未回補、無新缺口；MACD柱+9.85續正、KD黃金交叉第6日、RSI 48.64；外資賣超2,618張（門檻內）、三大合計+743張；BWIBBU換季P/E 23.92x屬分母更新；轉弱程度維持緩和）'
change_pct = '+0.43%（11/13收，相對大盤-0.93pp（精算-0.9262），轉為落後；加權+583.25點、+1.35%收43,675.12，全市場成交9,515.80億）；11/12為-2.30%'
nyse_price = 'US$398.73（11/13收，-1.67%，2年234個下跌日第102大，非極端；連2日下跌；費半-3.43%；ADR溢價同日對齊+11.44%（ADR 11/13對台股11/13與匯率32.84），11/12同日對齊+13.99%；盤後約394.5（-1.06%，19:59 EST））；11/12收US$405.49'
volume = '30,973（11/13收，-9.85%；為前報基準52,118張的0.59倍（-40.57%）＝量縮上漲、新5日均量36,169.4張；2026年208日中第132大）；11/12為34,359張'
news_summary = """11/13收2,350(+0.43%)收回MA20但落後大盤；外資賣超2,618張、三大合計轉買；費半-3.43%；維持看空。"""
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
