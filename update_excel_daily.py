"""
TSMC Excel 每日更新腳本 - 2026-05-06
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

today_str = "2026-05-06"
tw_price = "NT$2,250"        # 5/5 Tue 收盤（+25 vs 5/4 NT$2,275，+1.10%；盤中觸 NT$2,300 心理整數新高、量縮健康整理）
change_pct = "+1.10%"        # 台股 2330 5/5 收盤漲跌幅
nyse_price = "US$396.89"     # NYSE TSM 5/5 收盤（-1.18%；拉回跌破 $400 關卡、台美短線分歧）
volume = "21,565 張"         # 5/5 成交量（vs 5/4 92,800 張量縮 77%、爆量後健康整理）
news_summary = "台股2330 5/5收NT$2,250(+1.10%,+25)盤中觸NT$2,300心理整數新高、量縮21,565張(vs 5/4 92,800張縮77%)健康整理;加權指數同日+64.15(+0.16%)收40,769.29連2日站穩4萬;NYSE TSM 5/5 $396.89(-1.18%)拉回跌破$400、台美短線分歧;ADR溢價維持+9.7%;短線最大利多:AMD Q1盤後大超預期EPS $1.37 vs $1.29、營收$10.25B vs $9.89B、Q2指引$11.2B、AH +15%確認AI需求結構性轉強;TSMC重啟桃園龍潭Phase 3 fab($16.9B、A14/1.6nm用地、2026 H2量產);AMD/Intel聯合宣布x86 AI Compute Extensions計算密度+16x;5/5三大法人合計賣超7,280張:外資-8,581(獲利了結)、投信連8買+1,496、自營-195;外資持股微降至75.0%;Goldman Sachs維持NT$2,330、Barclays $470、共識NT$2,320(剩+3.1%);TSMC 4月月營收5/10前公布(估NT$330B+);NVDA 5/28財報"
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
