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

today_str = '2026-10-07'
tw_price = 'NT$2,585（10/6收，+10；開2,575高2,590低2,565，收盤與盤中高連2日創歷史新高）；10/5收NT$2,575（+75）'
change_pct = '+0.39%（10/6收，+10.00，領先大盤0.17pp；加權+0.22%連4日創收盤新高）；10/5為+3.00%（+75.00）'
nyse_price = 'US$482.30（10/6收，-0.72%，未創新高；ADR溢價+18.59%）；10/5收US$485.80（+2.75%）'
volume = '20,802（10/6收，-22.4%，低於前報門檻24,887張16.4%、低於5日均量23,669張12.1%；2026年183日中第23低）；10/5為26,800張'
news_summary = """報告日2026-10-07。10/6收2,585(+0.39%)連2日創歷史收盤新高，但量縮至20,802張、外資賣超1,672張，前報上調看多條件只過價格，維持中性偏多；退回中性防線皆守住，收盤續高於布林上軌，過熱未解。"""
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
