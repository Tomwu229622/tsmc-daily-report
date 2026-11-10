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

today_str = '2026-11-10'
tw_price = 'NT$2,235（11/09收，+100、+4.68%；開2,210（跳空+75）、高2,290、低2,190，紅K實體25元；2026年97個上漲日第8大漲幅，終止連7日下跌；收回布林下軌之上（實際2,175.37、估算2,152.32）、站回MA5 2,233（僅高2元）、盤中回補11/06的2,195–2,235缺口；續在MA10 2,310.5、MA20 2,383.25、MA60 2,419.67、半年線MA120 2,389.13之下，MA20下穿MA120；MACD柱-50.45收斂、KD黃金交叉第2日、RSI 35.47回到30之上；外資賣超8,758張連22賣（較前日減少14,086張、回到門檻內）；轉弱程度由再加深轉為緩和）'
change_pct = '+4.68%（11/09收，相對大盤-0.25pp（精算-0.2531），續落後；加權+1,960.06點、+4.94%收41,661.70，2026年第2大漲幅，全市場成交8,770.26億）；11/06為-4.69%'
nyse_price = 'US$388.57（11/09收，+3.50%，2年264個上漲日第44大，非極端；收復11/05跌幅的83.17%；費半+6.85%（2年第5大漲幅）；ADR溢價同日對齊+14.33%（ADR 11/09對台股11/09與匯率32.88），11/06同日對齊+15.25%；盤後約390.01（19:57 EST））；11/06收US$375.44（+6.33%）'
volume = '45,883（11/09收，-58.56%；為前報基準59,414張的0.77倍（-22.77%）＝量縮上漲、新5日均量60,357.8張；2026年204日中第50大）；11/06為110,715張'
news_summary = """11/09收2,235(+4.68%)終止連7跌、站回MA5但量縮；外資連22賣8,758張；ADR+3.50%，維持看空。"""
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
