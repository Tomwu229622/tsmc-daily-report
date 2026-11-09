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

today_str = '2026-11-09'
tw_price = 'NT$2,135（11/06收，-105、-4.69%；開盤即跌停2,020（＝全日最低）、高2,195，紅K實體115元但收低於前收；2026年95個下跌日第2大跌幅、連7日下跌；收盤與4/30持平、為4/23之後最低；跌破布林下軌（實際2,186.90、估算2,214.52）、留2,195–2,235新向下缺口；續在半年線MA120 2,389.42與五條均線之下，連14日低於MA60 2,422.08（差287.08元）；MACD柱-58.83、RSI 19.57雙創2026年新低，KD低檔黃金交叉；外資賣超22,844張連21賣（較前日擴大12,592張）；轉弱程度再加深）'
change_pct = '-4.69%（11/06收，相對大盤-0.18pp（精算-0.1822），由抗跌轉為落後；加權-1,873.08點、-4.51%收39,701.64，2026年第3大跌幅，全市場成交13,210.72億）；11/05為-0.88%'
nyse_price = 'US$375.44（11/06收，+6.33%，2年264個上漲日第7大，收復11/05跌幅的52.38%；費半+4.21%；ADR溢價同日對齊+15.25%（ADR 11/06對台股11/06與匯率32.77），11/05同日對齊+3.66%；盤後約378.58（19:59 EST））；11/05收US$353.10（-10.78%）'
volume = '110,715（11/06收，+145.08%；為前報基準44,594張的2.48倍（+148.27%）、新5日均量59,414.2張；2026年203日中第1大、2年第4）；11/05為45,175張'
news_summary = """11/06收2,135(-4.69%)跌破布林下軌、量創年內最大；外資連21賣22,844張；ADR+6.33%，維持看空。"""
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
