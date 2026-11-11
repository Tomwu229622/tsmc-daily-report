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

today_str = '2026-11-11'
tw_price = 'NT$2,285（11/10收，+50、+2.24%；開2,265（跳空+30）、高2,285、低2,230，收在全日最高、紅K實體20元；2026年98個上漲日第29，非極端；連2日上漲、11/03以來最高收盤；站穩MA5 2,231（高54元）、守在布林下軌之上（實際2,172.82、估算2,164.88）；仍低於MA10 2,299（差14元），續在MA20 2,370、MA60 2,417.83、半年線MA120 2,389.29之下；MACD柱-35.48續收斂且DIF回升、KD黃金交叉第3日、RSI 41.68回到40之上；外資賣超8,181張連23賣（較前日減少577張、維持門檻內）；轉弱程度續緩和）'
change_pct = '+2.24%（11/10收，相對大盤-0.28pp（精算-0.2837），續落後；加權+1,050.22點、+2.52%收42,711.92，全市場成交10,340.12億）；11/09為+4.68%'
nyse_price = 'US$396.48（11/10收，+2.02%（前收修訂為388.62；對前報發布值388.57為+2.04%），2年265個上漲日第103，非極端；連3日上漲；費半-0.35%；ADR溢價同日對齊+14.07%（ADR 11/10對台股11/10與匯率32.87），11/09同日對齊+14.33%；盤後約402.65（19:59 EST））；11/09收US$388.57（前報發布值）'
volume = '36,576（11/10收，-20.28%；為前報基準60,358張的0.61倍（-39.40%）＝量縮上漲、新5日均量58,445.6張；2026年205日中第93大）；11/09為45,883張'
news_summary = """11/10收2,285(+2.24%)連2漲、站穩MA5但量縮；外資連23賣8,181張；10月營收年增59.71%；維持看空。"""
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
