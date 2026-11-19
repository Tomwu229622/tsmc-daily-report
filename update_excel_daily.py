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

today_str = '2026-11-19'
tw_price = 'NT$2,175（11/18收，+5、+0.23%；開2,200（高開30元）、高2,210、低2,150，黑K實體25元、上影線10元、下影線25元；2026年101個上漲日並列第94大；結束連2日下跌，盤中最低為11/06以來最低；續在MA5 2,250、MA10 2,254、MA20 2,312.25、半年線MA120 2,391.13（連5日）與MA60 2,405.67（連22日）之下，MA5下穿MA10；無新缺口、守在布林下軌2,131.10之上；MACD柱-19.22續負、KD死叉延續、RSI 36.67；外資賣超15,452張（連2日逾門檻）、投信買超3,383張、三大合計-6,885張；月結算日量能字面放量上漲（實質未確認）；轉弱程度持平）'
change_pct = '+0.23%（11/18收，相對大盤-0.03pp（精算-0.0260），轉為落後；加權+102.80點、+0.26%收40,194.91，全市場成交10,751.74億）；11/17為-2.03%'
nyse_price = 'US$362.62（11/18收，+0.81%，2年264個上漲日第196大；結束連4日下跌；費半+0.10%；ADR溢價同日對齊+9.87%（ADR 11/18對台股11/18與匯率32.95），11/17同日對齊+8.90%；盤後約367（約+1.2%，19:59 EST））；11/17收US$359.69'
volume = '50,794（11/18收，-17.56%；為前報基準42,521張的1.19倍（+19.46%）＝字面放量上漲（月結算日，實質未確認）、新5日均量46,068.4張；2026年211日中第37大）；11/17為61,611張'
news_summary = """11/18收2,175(+0.23%)止跌但收黑K、MA5下穿MA10；外資賣超15,452張連2日逾門檻、投信續買；輝達財報後首盤-2.89%、ADR+0.81%；維持看空。"""
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
