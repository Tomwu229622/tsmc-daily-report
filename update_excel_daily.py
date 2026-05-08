"""
TSMC Excel 每日更新腳本 - 2026-05-08
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

today_str = "2026-05-08"
tw_price = "NT$2,310"        # 5/7 Thu 收盤（+60，+2.67% vs 5/6 NT$2,250；盤中觸 NT$2,345 創史高、市值衝 60.81 兆）
change_pct = "+2.67%"        # 台股 2330 5/7 收盤漲跌幅
nyse_price = "US$416.58"     # NYSE TSM 5/7 收盤（-0.70% vs 5/6 $419.50；高位整理、盤中觸 52 週新高 $420.00）
volume = "62,800 張"         # 5/7 成交量（爆量補漲跟進 NYSE 前夜 +6.36%）
news_summary = "台股2330 5/7爆量飆漲收NT$2,310(+60,+2.67%)創歷史新高、盤中觸NT$2,345(+95,+4.22%)史高、市值收盤59.9兆/盤中達60.81兆雙創高、量62,800張;TWii 5/7+794.93(+1.93%)收41,933.78創史高、盤中觸42,156.06首破4萬2大關;最後一盤爆7,112張賣單仍守天價;NYSE TSM 5/7高位整理收$416.58(-2.92,-0.70% vs 5/6 $419.50)、開$418.09高$420.00(盤中52週新高首破$420)、低$414.02消化前夜+6.36%;ADR溢價自+15.1%收斂至+12.2%、台美再度共振屬健康;5/7全市場三大法人合計買超+583.31億——外資爆量+464.12億、投信+160.62億(連9買)、自營-41.43億;TSMC為外資+投信當日買超第一大標的、外資5日累計買TSMC估+37,000張、外資持股升至75.3%;SOX 5/7-0.40%收11,205;NVDA+1.30%收$211.20、AMD-0.87%收$228.50、AVGO+0.57%、ASML-0.59%;Barclays $470對應NT$2,450(剩+6.1%)、共識NT$2,320(剩+0.4%即將突破);TSMC 4月月營收5/10前公布(估NT$330B+)、距今2天;NVDA 5/28 Q1 FY27財報距20天;技術面RSI估72.5進超買區、KDJ J108嚴重過熱、MACD+5.2紅柱續擴"
change_color = "00B050"      # 上漲綠色（+2.67%）

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
