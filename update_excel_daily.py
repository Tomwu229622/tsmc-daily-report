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

today_str = '2026-10-22'
tw_price = 'NT$2,420（10/21收，+15、+0.62%；開2,425高2,435低2,400，開高20元後走低收小黑K，盤中一度低於前收2,405；收盤連4日低於MA60但差距收斂至2.67元、距MA5僅1元；2,450–2,490缺口第5日未回補）；10/20收NT$2,405（-0.62%）'
change_pct = '+0.62%（10/21收，相對大盤-0.66pp；加權+594.84點、+1.28%收46,887.77，全市場成交8,746.14億）；10/20為-0.62%'
nyse_price = 'US$431.89（10/21收，+1.60%；連3日上漲、五日+0.44%；ADR溢價擴大至+15.97%，新台幣收32.488、取32.49）；10/20收US$425.08（+1.16%）'
volume = '27,095（10/21收，+20.95%；低於新5日均量35,745張24.2%、低於前報基準35,949張24.6%；2026年193日中第140）；10/20為22,402張'
news_summary = """10/21收2,420漲0.62%、落後大盤，外資賣超縮至6,153張仍連11賣，維持看空；MACD柱首度收斂；ADR+1.60%；高雄F22事故4死。"""
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
