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

today_str = '2026-11-25'
tw_price = 'NT$2,230（11/24收，+200、+9.85%；開高低收皆2,230＝漲停價，一字鎖漲停（K線無實體、振幅0）；2026年103個上漲日第2大；收回MA5 2,107與MA10 2,200.50、收回11/19收盤2,100；回補2,060–2,085缺口、留下2,050–2,230向上缺口；仍在MA20 2,249.75、半年線MA120 2,383.96（連9日）與MA60 2,385.67（連26日）之下；MACD柱-23.79續負收斂、KD死叉延續（差距-0.34）、RSI 47.52回到40之上；外資賣超4,098張（回到門檻內）、投信買超917張、三大合計-2,793張；字面量縮上漲、實質為鎖漲停機械結果；雙條件1/2、維持看空、轉弱程度緩和）'
change_pct = '+9.85%（11/24收，相對大盤+0.86pp（精算+0.8593），恢復相對強勢；加權+3,310.46點、+8.99%收40,122.35，全市場成交4,325.60億）；11/23為+1.50%'
nyse_price = 'US$349.20（11/24收，-3.96%，2年236個下跌日第29大；收在11/05收盤353.10之下1.10%；費半-7.88%；ADR溢價同日口徑+3.16%（ADR 11/24對台股11/24與匯率32.94，台股補漲與時差造成收斂），11/23同日口徑+17.86%；盤後約348（約-0.4%，19:59 ET））；11/23收US$363.60'
volume = '13,213（11/24收，-80.95%；為前報基準65,306張的0.20倍＝字面量縮上漲，實質為一字鎖漲停的機械結果、漲停價委買未成交張數未能查得；新5日均量55,626.4張；2026年215日中由小到大第2）；11/23為69,364張'
news_summary = """11/24一字鎖漲停收2,230(+9.85%)、收回MA5與MA10、回補2,060–2,085缺口、RSI回到40之上；外資賣超4,098張回到門檻內、投信續買；量縮屬鎖漲停機械結果；美股費半-7.88%、ADR-3.96%（皆在台股收盤後）；維持看空、轉弱緩和。"""
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
