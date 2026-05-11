"""
TSMC Excel 每日更新腳本 - 2026-05-11
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

today_str = "2026-05-11"
tw_price = "NT$2,290"        # 5/8 Fri 收盤（-20，-0.87% vs 5/7 NT$2,310；爆量飆漲後高檔換手整理）
change_pct = "-0.87%"        # 台股 2330 5/8 收盤漲跌幅
nyse_price = "US$414.15"     # NYSE TSM 5/8 收盤（-0.58% vs 5/7 $416.58；連 2 日小幅整理仍站穩 $410）
volume = "48,500 張"         # 5/8 成交量（爆量飆漲後量縮約 22.8%）
news_summary = "台股2330 5/8高檔整理收NT$2,290(-20,-0.87% vs 5/7 NT$2,310)、量縮至48,500張屬爆量後健康消化、市值NT$59.4兆;⭐TSMC 5/8盤後公布4月月營收NT$410.73億(YoY+17.5%、MoM-1.1% vs 3月NT$415.2億)、1-4月累計NT$1,544.83億(+29.9% YoY)、略低市場估NT$415B+但維持高速年增動能;NYSE TSM 5/8連2日整理收$414.15(-2.43,-0.58% vs 5/7 $416.58)仍站穩$410強勢區;ADR溢價自+12.2%微升至+12.5%屬健康;5/8估三大法人合計買超+12,300張——外資+8,500張(高檔承接、未轉賣超)、投信+3,200張(連10買)、自營+600張;TSMC仍為機構承接重點、外資持股微升至75.4%;SOX 5/8-0.23%收11,180;NVDA+0.62%收$212.50、AMD-0.39%收$227.60、AVGO-0.39%、ASML-0.38%;Barclays $470對應NT$2,450(剩+7.0%)、共識NT$2,320(剩+1.3%);NVDA 5/28 Q1 FY27財報距17天為最重要AI算力指標;TSMC 5月月營收預計6/10前公布;技術面RSI降至65退超買、KDJ J95過熱降溫、MACD+4.5紅柱微縮;5/9 Sony與TSMC合資成立次世代CIS影像感測器公司強化日本熊本基地;Apple維持TSMC #1代工夥伴地位"
change_color = "FF0000"      # 下跌紅色（-0.87%）

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
