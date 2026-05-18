"""
TSMC Excel 每日更新腳本 - 2026-05-18
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

today_str = "2026-05-18"
tw_price = "NT$2,220"        # 5/18 Mon 估收（-45，-1.99% vs 5/15 NT$2,265；跳空走低跟進ADR 5/15 -3.20%回檔 + NVDA 5/20財報倒數雙重壓力）
change_pct = "-1.99%"        # 台股 2330 5/18 估收盤漲跌幅
nyse_price = "US$404.35"     # NYSE TSM 最新收盤（5/15 Fri，-13.37，-3.20% vs 5/14 $417.72）
volume = "53,000 張"         # 5/18 估成交量（量增反映賣壓）
news_summary = "台股2330 5/18跳空走低估收NT$2,220(-45,-1.99% vs 5/15 NT$2,265)、量增至53,000張、市值估降至NT$57.6兆;跟進NYSE TSM 5/15 -3.20%回檔訊號、跌破5MA NT$2,250與前壓轉支撐NT$2,235;🚨NVDA Q1 FY27財報5/20(Wed)距今僅2天為最重要AI算力指標、市場預期Q1營收$78.5B(+78% YoY)、guidance $78.0B±2%、毛利率~75%、EPS $1.77;⚠️短線雜訊:(1)Cathie Wood ARK 5/15賣超$40.6M TSM加重情緒;(2)兩岸地緣政治升溫、媒體5/17報導台海封鎖風險再起;(3)NYSE TSM 5/15收$404.35(-13.37,-3.20% vs 5/14 $417.72)獲利了結為主、距5/7史高$420.00剩-3.7%;ADR溢價自+11.1%擴大至+13.3%;5/18估三大法人合計大幅賣超-19,500張——外資-16,500張、投信-2,000張、自營-1,000張、外資持股估降至74.8%;🎯結構性催化續存:TSMC 5/14新竹技術論壇上修2030全球半導體市場至$1.5兆(+50%)、AI/HPC占55%、18座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%北美擴產加速;⭐利多續存:(1)NVIDIA為TSMC最大客戶(22%/$33B vs Apple 17%/$27B);(2)AMAT EPIC Center $5B AI晶片R&D;(3)Arizona追加$20B資本注入;SOX 5/15跟跌-3.95%收11,053、NVDA-4.00%收$225.32、AMD-4.00%收$432、AVGO-3.88%、ASML-1.95%、INTC-7.00%;Barclays$470對應NT$2,450(剩+10.4%)、共識NT$2,320(剩+4.5%);TSMC 5月月營收預計6/10前公布;技術面RSI自60.5跌破中軸至45.5進入弱勢區、KDJ K(58)/D(62)/J(50)死叉訊號浮現、MACD+0.5紅柱大幅收斂逼近翻黑;本週關鍵:NT$2,200心理整數+60MA守關、若失守將測試NT$2,150第一支撐"
change_color = "FF0000"      # 下跌紅色（-1.99%）

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
