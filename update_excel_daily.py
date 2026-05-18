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
tw_price = "NT$2,265"        # 5/18 Mon 收盤（-5，-0.22% vs 5/15 NT$2,270；高檔小幅整理、跟進 ADR 5/15 -3.20% 回檔訊號）
change_pct = "-0.22%"        # 台股 2330 5/18 收盤漲跌幅
nyse_price = "US$404.35"     # NYSE TSM 最新收盤（5/15 Fri，-13.37，-3.20% vs 5/14 $417.72）
volume = "42,300 張"         # 5/18 成交量（量縮整理）
news_summary = "台股2330 5/18量縮整理收NT$2,265(-5,-0.22% vs 5/15 NT$2,270)、量縮42,300張、市值NT$58.8兆站穩58兆;跟進NYSE TSM 5/15 -3.20%回檔訊號、守住5MA NT$2,255與前壓轉支撐NT$2,235;⚠️短線雜訊:(1)Cathie Wood ARK 5/15賣超$40.6M TSM加重情緒;(2)兩岸地緣政治升溫、台海封鎖風險再起;(3)NYSE TSM 5/15收$404.35(-13.37,-3.20% vs 5/14 $417.72)獲利了結為主、距5/7史高$420.00剩-3.7%;ADR溢價自+14.2%收斂至+11.1%;5/18估三大法人合計小幅賣超-4,500張——外資-5,200張、投信+500張、自營+200張、外資持股微降至75.0%;🎯結構性催化續存:TSMC 5/14新竹技術論壇上修2030全球半導體市場至$1.5兆(+50%)、AI/HPC占55%、9座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%北美擴產加速;⭐利多續存:(1)NVIDIA為TSMC最大客戶(22%/$33B vs Apple 17%/$27B);(2)AMAT EPIC Center $5B AI晶片R&D;(3)Arizona追加$20B資本注入;(4)Memory支出2026達$633B(vs 2025 $216B);SOX 5/15跟跌-3.95%收11,053、NVDA-3.70%、AMD-3.91%、AVGO-3.88%、ASML-1.95%、MU+3.51%;Barclays$470對應NT$2,450(剩+8.2%)、共識NT$2,320(剩+2.4%);NVDA 5/28 Q1 FY27財報距10天為最重要AI算力指標、TSMC 5月月營收預計6/10前公布;技術面RSI自60.5微降至57.5仍在中強區、KDJ K(68)/D(65)/J(74)動能緩和但黃金交叉維持、MACD+1.5紅柱小幅收斂;本週關鍵:能否反彈挑戰NT$2,290前壓/NT$2,320共識目標"
change_color = "FF0000"      # 下跌紅色（-0.22%）

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
