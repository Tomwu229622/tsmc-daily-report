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

today_str = "2026-04-24"
tw_price = "NT$2,080"        # 4/23 收盤（+30 vs 4/22 NT$2,050，+1.46%；跟進 ADR 大漲）
change_pct = "+1.46%"        # 台股 2330 4/23 收盤漲跌幅
nyse_price = "US$387.53"     # NYSE TSM 4/23 收盤（-0.05% 技術整理）
volume = "~22,500 張"        # 4/23 成交量估算
news_summary = "台股2330 4/23收NT$2,080(+1.46%)跟進ADR大漲;NYSE TSM 4/23 $387.53(-0.05%)技術整理;TSMC宣布暫不採用ASML高NA EUV(€350M/台過貴)→A13/N2U設計協同突破,ASML股價-2.6%;Siemens擴大與TSMC AI晶片設計合作;2028先進封裝將整合10大晶片+20 HBM;NVIDIA確認#1客戶22%,黃仁勳稱AI晶片瓶頸為2-3年問題"
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
