"""
TSMC Excel 每日更新腳本 - 2026-05-12
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

today_str = "2026-05-12"
tw_price = "NT$2,235"        # 5/11 Mon 收盤（-55，-2.40% vs 5/8 NT$2,290；4月營收略低預期+獲利了結回檔）
change_pct = "-2.40%"        # 台股 2330 5/11 收盤漲跌幅
nyse_price = "US$404.54"     # NYSE TSM 5/11 收盤（-1.69% vs 5/8 $411.51；跟進台股回檔仍站穩 $400）
volume = "55,200 張"         # 5/11 成交量（回檔放量）
news_summary = "台股2330 5/11跌破5MA收NT$2,235(-55,-2.40% vs 5/8 NT$2,290)、量放55,200張、市值NT$57.9兆;受4月營收MoM-1.1%略低市場估NT$415B+衝擊+爆量飆漲後獲利了結賣壓共振;NYSE TSM 5/11跟進回檔-1.69%收$404.54(-6.97)仍站穩$400強勢區;ADR溢價自+12.5%維持+12.6%台美共振;5/11估三大法人合計賣超-21,800張——外資轉賣-18,500張、投信轉賣-2,500張(連10買終止)、自營-800張獲利了結、外資持股微降至75.0%;⭐重大利多:Applied Materials(AMAT)與TSMC 5/11宣布矽谷EPIC Center共建$5B AI晶片研發合作中心、美國史上最大半導體設備R&D投資、AMAT當日大漲+3.92%;NVIDIA對TSMC採購承諾揭露已逾$95B(兩年前僅$16B)、AI加速器2029CAGR目標54-56%;SOX 5/11-1.74%收10,985、NVDA+4.46%逆勢創高$221.98、AMD+1.80%逆勢漲$463.40、AVGO-0.26%、ASML-1.78%;Barclays$470對應NT$2,450(剩+9.6%)、共識NT$2,320(剩+3.8%);NVDA 5/28 Q1 FY27財報距16天為最重要AI算力指標、TSMC 5月月營收預計6/10前公布;技術面RSI降至53趨近中性、KDJ J55過熱徹底消化、MACD+2.5紅柱明顯收斂、潛在死叉風險;Siemens EDA與TSMC論壇深化AI設計工具合作"
change_color = "FF0000"      # 下跌紅色（-2.40%）

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
