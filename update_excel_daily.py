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

today_str = '2026-11-23'
tw_price = 'NT$2,000（11/20收，-100、-4.76%；開2,035（低開65元）、高2,060、低1,970，黑K實體35元、上影線25元、下影線30元；2026年100個下跌日第3大；留下2,060–2,085新向下缺口；收盤跌破11/06盤中最低2,020、為4/13以來最低收盤；續在MA5 2,132、MA10 2,226.50、MA20 2,276、半年線MA120 2,387.21（連7日）與MA60 2,394.25（連24日）之下；收盤低於前報估算布林下軌2,073.91與實際下軌2,047.91（%B -0.11）；MACD柱-47.55續負、KD死叉擴大、RSI 27.85跌破30；外資賣超11,862張（回到門檻外）、投信賣超3,848張、三大合計-10,901張；放量下跌；轉弱程度加劇）'
change_pct = '-4.76%（11/20收，相對大盤+1.47pp（精算+1.4671），延續相對抗跌；加權-2,393.60點、-6.23%收36,033.29，全市場成交9,353.10億）；11/19為-3.45%'
nyse_price = 'US$318.71（11/20收，-9.48%，2年237個下跌日第3大；收在11/05收盤353.10之下9.74%；費半-10.73%；ADR溢價同日對齊+4.98%（ADR 11/20對台股11/20與匯率32.94），11/19同日對齊+10.46%；盤後約329（約+3.1%，19:59 EST））；11/19收US$352.09'
volume = '88,179（11/20收，+55.84%；為前報基準50,513張的1.75倍（+74.57%）＝放量下跌、新5日均量61,954.2張；2026年213日中第5大、11/06以來最大量）；11/19為56,582張'
news_summary = """11/20收2,000(-4.76%)放量下跌、留2,060–2,085新向下缺口、跌破11/06盤中最低；外資賣超11,862張回到門檻外、投信轉賣；美股費半-10.73%、ADR-9.48%；維持看空、轉弱加劇。"""
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
