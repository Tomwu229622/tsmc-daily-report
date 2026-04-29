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

today_str = "2026-04-29"
tw_price = "NT$2,215"        # 4/28 收盤（-50 vs 4/27 NT$2,265，-2.21%；OpenAI 不達標引爆 AI 晶片股拋售）
change_pct = "-2.21%"        # 台股 2330 4/28 收盤漲跌幅
nyse_price = "US$392.34"     # NYSE TSM 4/28 收盤（-3.12%；跌破 $400 關卡）
volume = "46,932 張"         # 4/28 成交量（量縮 25% vs 4/27 6.28 萬張）
news_summary = "台股2330 4/28收NT$2,215(-2.21%,-50)自史高拉回;NYSE TSM 4/28 $392.34(-3.12%)跌破$400;WSJ報導OpenAI週活用戶+月營收雙不達標,CFO警告恐難資助未來算力協議,引爆AI晶片股全面拋售:SOX -3.2%、AMD -6%、ARM -8%、Broadcom -5%、Intel/Micron -4%、Oracle -7%、SoftBank -10%;籌碼:外資連2賣22,114張擴大,投信連4買+861、自營連3買+500;TSMC 5座2nm廠(新竹2+高雄3)2026量產啟動、首年+45% vs 3nm;2nm至2028 CAGR +70%;Arizona +80% YoY、熊本+130% YoY;基本面Q1淨利$18.2B/毛利率66.2%雙創史高不變"
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
