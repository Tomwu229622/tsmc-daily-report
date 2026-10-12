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

today_str = '2026-10-12'
tw_price = 'NT$2,550（10/8收，-35、-1.35%；開2,550高2,575低2,550，開低走低收在全日最低、倒T字線；10/9國慶補假休市）；10/7收NT$2,585（平盤）'
change_pct = '-1.35%（10/8收，相對大盤-0.36pp；加權-492.93點、-0.99%收49,313.44、連2日下跌）；10/7為0.00%'
nyse_price = 'US$453.31（10/9收，-1.02%；10/8收US$457.99、-3.01%，兩日累計-4.00%、連4日下跌；ADR溢價壓縮至+13.34%）；10/7收US$472.20（-2.09%）'
volume = '23,145（10/8收，+36.6%，高於5日均量20,696張11.8%、成交筆數141,976筆倍增；2026年185日中第153）；10/7為16,941張'
news_summary = """報告日2026-10-12（連假後首報，涵蓋台股10/8與美股10/8、10/9）。10/8開低走低收最低2,550（-1.35%），量增至23,145張、外資賣超12,302張（8/04以來最大）觸發退回條件、立場退回中性，KD死叉、J跌破100；盤後9月營收5,118.57億、Q3合計1.4942兆以32元換算超指引上緣；休市期間ADR兩日-4.00%、夜盤個股期2,535低於現貨15點，今日為營收利多與美股利空首度對決。"""
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
