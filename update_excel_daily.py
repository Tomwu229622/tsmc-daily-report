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

today_str = '2026-11-24'
tw_price = 'NT$2,030（11/23收，+30、+1.50%；開1,990（低開10元）、高2,050、低1,970，紅K實體40元、上影線20元、下影線20元；2026年102個上漲日第45大；結束連2日下跌；收回布林下軌（實際2,012.30，%B 0.04）與11/06盤中最低2,020；2,060–2,085缺口未補；續在MA5 2,095、MA10 2,206、MA20 2,258.25、半年線MA120 2,385（連8日）與MA60 2,388.50（連25日）之下；MACD柱-50.98續負、KD死叉延續、RSI 31.01回到30之上；外資賣超11,039張（仍在門檻外）、投信買超1,294張、三大合計-7,080張；放量上漲；轉弱程度緩和）'
change_pct = '+1.50%（11/23收，相對大盤-0.66pp（精算-0.6608），轉為落後；加權+778.60點、+2.16%收36,811.89，全市場成交11,137.43億）；11/20為-4.76%'
nyse_price = 'US$363.60（11/23收，+14.08%，2年264個上漲日第1大；收在11/05收盤353.10之上2.97%；費半+14.03%；ADR溢價同日口徑+17.86%（ADR 11/23對台股11/23與匯率32.90，美股交易晚於台股收盤，擴大屬時差），11/20同日口徑+4.98%；盤後約370（約+1.8%，20:00 EST））；11/20收US$318.71'
volume = '69,364（11/23收，-21.34%；為前報基準61,954張的1.12倍（+11.96%）＝放量上漲（禁平盤下放空首日，保留但書）、新5日均量65,306.0張；2026年214日中第13大）；11/20為88,179張'
news_summary = """11/23收2,030(+1.50%)放量上漲、收回布林下軌與11/06盤中最低2,020、RSI回到30之上；外資賣超11,039張仍逾門檻、投信轉買；美股費半+14.03%、ADR+14.08%（皆在台股收盤後）；維持看空、轉弱緩和。"""
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
