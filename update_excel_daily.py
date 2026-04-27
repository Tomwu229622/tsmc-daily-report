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

today_str = "2026-04-27"
tw_price = "NT$2,185"        # 4/24 收盤（+105 vs 4/23 NT$2,080，+5.05%；爆量飆漲創歷史新高）
change_pct = "+5.05%"        # 台股 2330 4/24 收盤漲跌幅
nyse_price = "US$402.46"     # NYSE TSM 4/24 收盤（+5.17% 首次站上 $400 創歷史新高）
volume = "44,962 張"         # 4/24 成交量實際（量能翻倍）
news_summary = "台股2330 4/24爆量飆漲收NT$2,185(+5.05%,+105)創歷史新高,量能44,962張較前日翻倍,市值NT$56.6兆(+2.7兆);NYSE TSM 4/24 $402.46(+5.17%)首次站上$400創歷史新高;SOX寫30年新高(+4.32%)連18日漲創紀錄;三大法人合計買超9,729張(外資+8,306,投信+1,168,自營+256);TSMC A13/A12/N2U三大新製程確認2029量產,A16延至2027;Arizona先進封裝廠2029啟用;NVIDIA #1客戶22%確認"
change_color = "00B050"  # 上漲綠色

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
