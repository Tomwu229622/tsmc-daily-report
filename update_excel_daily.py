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

today_str = '2026-10-27'
tw_price = 'NT$2,385（10/22收，-35、-1.45%；開2,435＝最高、收在最低2,385，無影線長黑K、9/16以來最低收盤，跌破法說後低點2,390；收盤連5日低於MA60且差38.25元、僅高於半年線1元；2,450–2,490缺口第6日未回補；10/23颱風休市、10/26補假）；10/21收NT$2,420（+0.62%）'
change_pct = '-1.45%（10/22收，相對大盤+0.36pp；加權-846.34點、-1.81%收46,041.43，全市場成交9,334.37億）；10/21為+0.62%'
nyse_price = 'US$424.37（10/26收，+1.40%；10/22 -4.52%→412.35、10/23 +1.49%→418.51；三日累計-1.74%；ADR溢價混合口徑+15.69%（ADR 10/26對台股10/22與匯率32.51），10/22同日對齊+12.42%）；10/21收US$431.89（+1.60%）'
volume = '29,797（10/22收，+9.97%；低於前報基準35,745張16.6%、約當新5日均量30,007張；2026年194日中第127）；10/21為27,095張'
news_summary = """10/22收2,385跌1.45%、破法說低點，外資賣1.27萬張連12賣，維持看空；休市期間費半先跌6.72%後反彈，ADR三日-1.74%。"""
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
