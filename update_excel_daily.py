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

today_str = '2026-10-08'
tw_price = 'NT$2,585（10/7收，平盤0.00；開2,565高2,585低2,560，與10/6歷史最高收盤同價、未再創新高）；10/6收NT$2,585（+10）'
change_pct = '0.00%（10/7收，平盤，相對大盤+0.03pp；加權-0.03%、連4日創收盤新高中止）；10/6為+0.39%（+10.00）'
nyse_price = 'US$472.20（10/7收，-2.09%，跌幅為費半-1.15%的1.8倍；ADR溢價+16.18%、連2日收斂）；10/6收US$482.30（-0.72%）'
volume = '16,941（10/7收，-18.6%，低於前報基準23,669張28.4%、低於新5日均量20,200張16.1%；2026年184日中第11低）；10/6為20,802張'
news_summary = """報告日2026-10-08。10/7開低拉回收平盤2,585，平歷史最高收盤但未創新高，量再縮至16,941張、外資連2日小賣1,163張，上調看多三條件全部未過、退回防線守住，維持中性偏多但信心標低；ADR-2.09%、夜盤個股期2,570低於現貨，連假前偏弱。"""
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
