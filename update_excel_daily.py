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

today_str = '2026-10-30'
tw_price = 'NT$2,415（10/29收，-10、-0.41%；開低2,420、高2,425、低2,390回補10/28向上缺口，實體5元小黑K、下影線25元；續站上MA5 2,409（連2日），但連8日低於MA60 2,432.5（差17.5元）、連11日低於MA10 2,422，MA10下穿MA60屬滾出值機械效果；2,450–2,490缺口第9日未回補）；10/28收NT$2,425（+1.04%）'
change_pct = '-0.41%（10/29收，相對大盤+1.01pp（精算+1.0062）；加權-667.25點、-1.42%收46,368.33，全市場成交8,238.13億）；10/28為+1.04%'
nyse_price = 'US$431.81（10/29收，-1.75%；結束連4日上漲；ADR溢價同日對齊+16.19%（ADR 10/29對台股10/29與匯率32.49），10/28同日對齊+17.52%；盤後419.42（-2.87%，19:59 ET））；10/28收US$439.51（+0.60%）'
volume = '18,404（10/29收，-15.84%；低於前報基準27,867張34.0%、新5日均量27,067張；2026年197日中第182）；10/28為21,867張'
news_summary = """10/29收2,415跌0.41%，大盤重挫1.42%中相對抗跌1.01pp；外資賣超縮至10張仍連15賣，MA60差距擴至17.5元、維持看空；ADR-1.75%、溢價+16.19%。"""
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
