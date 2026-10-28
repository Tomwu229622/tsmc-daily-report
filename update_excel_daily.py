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

today_str = '2026-10-28'
tw_price = 'NT$2,400（10/27收，+15、+0.63%；連休四日後復市首日，開2,380＝最低、高2,405、上影線5元小紅K；守住半年線MA120 2,386.21（高13.79元），但仍連11日低於MA5、連6日低於MA60（差25.25元）；2,450–2,490缺口第7日未回補）；10/22收NT$2,385（-1.45%）'
change_pct = '+0.63%（10/27收，相對大盤+0.51pp（精算+0.5099）；加權+54.80點、+0.12%收46,096.23，全市場成交10,316.84億）；10/22為-1.45%'
nyse_price = 'US$436.87（10/27收，+2.95%；連3日上漲；ADR溢價同日對齊+18.14%（ADR 10/27對台股10/27與匯率32.45），10/22同日對齊+12.42%、10/26混合口徑+15.69%）；10/26收US$424.37（+1.40%）'
volume = '38,173（10/27收，+28.11%；高於前報基準30,007張27.2%、新5日均量29,032張；2026年195日中第80）；10/22為29,797張'
news_summary = """10/27復市收2,400漲0.63%、放量38,173張，外資賣超縮至1,618張仍連13賣、三大法人轉買2,709張，未回MA60、維持看空；美股費半+3.60%、ADR+2.95%，溢價+18.14%。"""
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
