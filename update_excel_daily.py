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

today_str = '2026-11-20'
tw_price = 'NT$2,100（11/19收，-75、-3.45%；開2,165（低開10元）、高2,190、低2,085，黑K實體65元、上影線25元、下影線15元；2026年99個下跌日第8大；收盤跌破11/06收盤2,135、為4/23以來最低收盤；續在MA5 2,202、MA10 2,240、MA20 2,297、半年線MA120 2,389.71（連6日）與MA60 2,401.08（連23日）之下；收盤低於前報估算布林下軌2,115.46（字面跌破）、高於實際下軌2,099.08達0.92元；MACD柱-30.33續負、KD死叉擴大、RSI 32.45；外資賣超8,634張（回到門檻內）、投信買超1,579張、三大合計-2,504張；放量下跌；轉弱程度加劇）'
change_pct = '-3.45%（11/19收，相對大盤+0.95pp（精算+0.9503），恢復相對抗跌；加權-1,768.02點、-4.40%收38,426.89，全市場成交9,429.89億）；11/18為+0.23%'
nyse_price = 'US$352.09（11/19收，-2.90%，2年236個下跌日並列第52大；收在11/05收盤353.10之下；費半-2.41%；ADR溢價同日對齊+10.46%（ADR 11/19對台股11/19與匯率32.94），11/18同日對齊+9.87%；盤後約345（約-1.9%，19:59 EST））；11/18收US$362.62'
volume = '56,582（11/19收，+11.40%；為前報基準46,068張的1.23倍（+22.82%）＝放量下跌、新5日均量50,513.0張；2026年212日中第21大）；11/18為50,794張'
news_summary = """11/19收2,100(-3.45%)放量下跌、跌破11/06收盤；外資賣超縮至8,634張回到門檻內、投信續買；美股跌逾2%、ADR-2.90%；維持看空、轉弱加劇。"""
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
