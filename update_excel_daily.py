"""
TSMC Excel 每日更新腳本 - 2026-05-07
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

today_str = "2026-05-07"
tw_price = "NT$2,250"        # 5/6 Wed 收盤（持平 vs 5/5 NT$2,250，0.00%；TWii 大漲下台股 TSMC 持平、待 5/7 跟進 NYSE +5.55%）
change_pct = "0.00%"         # 台股 2330 5/6 收盤漲跌幅（持平）
nyse_price = "US$416.30"     # NYSE TSM 5/6 收盤（+5.55%；創 52 週新高 $417.68 盤中、AMD Q1 + AI CapEx $725B 引爆）
volume = "28,400 張"         # 5/6 成交量（量能略放大、盤中觸 NT$2,275）
news_summary = "台股2330 5/6收NT$2,250(0.00%持平)盤中觸NT$2,275、量28,400張;TWii 5/6+369.56(+0.91%)爆量1.4491兆收41,138.85創歷史新高連3日創高;NYSE TSM 5/6大漲+5.55%收$416.30(+19.41 vs 5/5 $396.89)盤中觸52週新高$417.68;ADR溢價自+9.7%擴至+15.1%、台美分歧達極端、預期5/7台股強勢補漲;驅動:AMD 5/6飆漲+25%反映Q1大超預期(EPS $1.37 vs $1.29、營收$10.25B、Q2指引$11.2B);4大美國CSP(Alphabet/AWS/MSFT/Meta)2026 AI CapEx合計$725B(+77% YoY)確認H2算力加速;TSMC重啟桃園龍潭Phase 3($16.9B、A14/1.6nm用地、2026 H2量產);SOX 5/6+5.41%收11,250.45;NVDA 5/6+6.30%收$208.50、AMD+24.93%收$230.50、AVGO+5.89%、ASML+4.88%;5/6全市場外資爆量大買+751.05億、投信+17.8億、自營-96.11億合計+672.85億;外資TSMC部位估強化買進、持股估升至75.1%;Goldman Sachs NT$2,330、Barclays $470對應NT$2,450、共識NT$2,320(剩+3.1%);TSMC 4月月營收5/10前公布(估NT$330B+)、距今3天;NVDA 5/28 Q1 FY27財報距21天"
change_color = "808080"      # 持平灰色（0.00%）

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
