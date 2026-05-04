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

today_str = "2026-05-04"
tw_price = "NT$2,135"        # 4/30 收盤（-45 vs 4/29 NT$2,180，-2.06%；連 4 日修正、跌破 NT$2,150 第一支撐；5/1 勞動節休市）
change_pct = "-2.06%"        # 台股 2330 4/30 收盤漲跌幅
nyse_price = "US$397.67"     # NYSE TSM 5/1 收盤（+0.41%；ADR 連 3 反彈、逼近 $400 關卡）
volume = "47,200 張"         # 4/30 成交量（量微增 10% vs 4/29 42,860 張）
news_summary = "台股2330 4/30收NT$2,135(-2.06%,-45)連4日修正、跌破NT$2,150第一支撐,距4/27史高NT$2,265累計-5.74%;5/1台股勞動節休市;NYSE TSM 5/1 $397.67(+0.41%)連3日反彈逼近$400;ADR對台股溢價自+13.1%擴大至+16.2%;4/30籌碼:外資連4賣估-10,200張、投信連6買+480、自營連5買+260,合計賣超9,460張,外資持股74.9%;美股5/1 AI晶片股延續反彈:NVDA $940.20(+0.84%)、AMD +0.60%、AVGO +0.95%、SOX 10,540(+0.28%)連3日漲;OpenAI雜訊基本消化;TSMC 4月月營收5/10前公布(預估NT$330B+);基本面Q1淨利$18.2B/毛利率66.2%雙創史高不變"
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
