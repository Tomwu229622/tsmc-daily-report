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

today_str = "2026-05-13"
tw_price = "NT$2,210"        # 5/13 Wed 收盤（-45，-2.00% vs 5/12 NT$2,255；Apple 傳轉單英特爾 + 獲利了結）
change_pct = "-2.00%"        # 台股 2330 5/13 收盤漲跌幅
nyse_price = "US$404.54"     # NYSE TSM 最新收盤（5/12 Tue，5/13 美股盤中未收）
volume = "50,200 張"         # 5/13 成交量（回檔放量）
news_summary = "台股2330 5/13跌破5MA收NT$2,210(-45,-2.00% vs 5/12 NT$2,255)、量放50,200張、市值NT$57.3兆;下跌主因:⚠️Apple傳考慮將部分晶片訂單轉至Intel Foundry 18A製程、客戶集中度風險升溫(Apple佔比17%)+爆量飆漲後獲利了結賣壓共振;Intel同日大漲+5.66%;NYSE TSM最新收盤(5/12)$404.54(-1.73% vs 5/11 $411.68)仍站穩$400強勢區;ADR溢價自+9.6%擴大至+13.9%、台股回檔速度大於ADR;5/13估三大法人合計賣超-15,500張——外資轉賣-12,800張、投信賣-2,000張、自營-700張、外資持股微降至74.8%;⭐利多續存:(1)TSMC Arizona子公司董事會5/12批准追加$20B資本注入、美國Phase 2/3擴張(美國史上單一最大資本承諾);(2)AMAT $5B EPIC Center AI晶片R&D合作;(3)Sony 5/9合資CIS熊本擴張;(4)2nm 2026-2028產能CAGR +70%、5座2nm廠今年量產啟動;(5)NVIDIA採購承諾揭露逾$95B、AI加速器2029CAGR目標54-56%;SOX 5/12+0.69%收10,895、NVDA-1.19%、AMD-0.98%、AVGO-0.89%、ASML-0.41%;Barclays$470對應NT$2,450(剩+10.9%)、共識NT$2,320(剩+5.0%);NVDA 5/28 Q1 FY27財報距15天為最重要AI算力指標、TSMC 5月月營收預計6/10前公布;技術面RSI降至48跌破中軸、KDJ K(58)/D(60)/J(54)死叉訊號浮現、MACD+1.5紅柱續收斂逼近翻黑;明日5/14關鍵:守NT$2,200心理整數+NT$2,190 20MA雙重支撐"
change_color = "FF0000"      # 下跌紅色（-2.00%）

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
