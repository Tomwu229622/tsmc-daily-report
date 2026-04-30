"""
TSMC Excel 每日更新腳本 - 2026-03-19
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

today_str = "2026-04-30"
tw_price = "NT$2,180"        # 4/29 收盤（-35 vs 4/28 NT$2,215，-1.58%；連 3 日修正、跌破心理 NT$2,200）
change_pct = "-1.58%"        # 台股 2330 4/29 收盤漲跌幅
nyse_price = "US$393.79"     # NYSE TSM 4/29 收盤（+0.37%；ADR 隔夜微反彈、未站回 $400）
volume = "42,860 張"         # 4/29 成交量（續縮量約 9% vs 4/28 46,932 張）
news_summary = "台股2330 4/29收NT$2,180(-1.58%,-35)連3日修正、跌破5MA NT$2,250+心理NT$2,200,距4/27史高NT$2,265累計-3.75%;NYSE TSM 4/29 $393.79(+0.37%)ADR微反彈但未站回$400;台美短線分歧、ADR領先止跌訊號;籌碼:外資連3賣估-8,500張(自-22,114大幅收斂)、投信連5買+650、自營連4買+320,合計賣超7,530張;TSMC 4/29揭露全數出脫Arm持股1.11M股@$207.65共$231M(財務性投資非戰略);NVDA +0.79%、AMD +1.21%、AVGO +1.14%、SOX收平,AI賣壓緩解;5座2nm廠2026量產啟動、首年+45% vs 3nm;基本面Q1淨利$18.2B/毛利率66.2%雙創史高不變"
change_color = "FF0000"  # 下跌紅色

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
