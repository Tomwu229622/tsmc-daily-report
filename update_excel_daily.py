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

today_str = "2026-05-15"
tw_price = "NT$2,275"        # 5/15 Fri 收盤（+50，+2.25% vs 5/14 NT$2,225；跟進 ADR 大漲訊號）
change_pct = "+2.25%"        # 台股 2330 5/15 收盤漲跌幅
nyse_price = "US$417.72"     # NYSE TSM 最新收盤（5/14 Thu，+17.92，+4.48% vs 5/13 $399.80）
volume = "58,500 張"         # 5/15 成交量（放量上漲）
news_summary = "台股2330 5/15大漲收NT$2,275(+50,+2.25% vs 5/14 NT$2,225)、放量58,500張、市值NT$59.0兆;跟進NYSE TSM 5/14 +4.48%大漲訊號、突破5MA NT$2,255與前壓轉支撐NT$2,235;🎯重大催化:TSMC 5/14新竹技術論壇上修2030全球半導體市場至$1.5兆(自原$1兆、+50%)、AI/HPC占55%、手機20%、車用10%、2026規劃9座新晶圓廠、2nm/A16製程2026-2028 CAGR +70%、Arizona 2026產出1.8倍YoY且良率追平台灣、AI加速器晶圓需求2022-2026成長11倍;NYSE TSM 5/14收$417.72(+17.92,+4.48% vs 5/13 $399.80)距5/7史高$420.00僅-0.5%;ADR溢價擴大至+14.2%;5/15估三大法人合計買超+18,500張——外資大買+16,800張、投信+1,200張、自營+500張、外資持股回升至75.1%;⭐利多續存:(1)NVIDIA為TSMC最大客戶(22%/$33B vs Apple 17%/$27B、$95B採購承諾);(2)AMAT EPIC Center $5B AI晶片R&D合作;(3)Arizona追加$20B資本注入;(4)台灣對美半導體投資$200B+;(5)SOX YTD +65%、全球半導體2026預估$975B創高;SOX 5/14 +4.90%收11,508、NVDA+4.28%、AMD+4.98%、AVGO+4.85%、ASML+3.78%、MU+10.78%;Barclays$470對應NT$2,450(剩+7.7%)、共識NT$2,320(剩+2.0%);NVDA 5/28 Q1 FY27財報距13天為最重要AI算力指標、TSMC 5月月營收預計6/10前公布;技術面RSI升至60.5進中強區、KDJ K(72)/D(64)/J(88)黃金交叉延續、MACD+1.7紅柱重新擴張;下週一5/18關鍵:站穩NT$2,275+挑戰NT$2,290前壓/NT$2,320共識目標"
change_color = "00B050"      # 上漲綠色（+2.25%）

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
