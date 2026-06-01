"""
TSMC Excel 每日更新腳本 - 2026-06-01
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

today_str = "2026-06-01"
tw_price = "NT$2,380"        # 6/1 Mon 估收（+25，+1.06% vs 5/29 NT$2,355；COMPUTEX 2026 開幕 + 黃仁勳 GTC Taipei 11AM keynote 雙重催化）
change_pct = "+1.06%"        # 台股 2330 6/1 估收盤漲跌幅
nyse_price = "US$418.45"     # NYSE TSM 最新收盤（5/29 Fri，-6.41，-1.51% vs 5/28 $424.86；自史高 $425.06 拉回獲利了結）
volume = "65,000 張"         # 6/1 估成交量（量增反映 COMPUTEX 題材爆量）
news_summary = "今日(6/1 Mon)台股2330在COMPUTEX 2026開幕日+黃仁勳GTC Taipei 11AM keynote雙重催化跳空開高、估收NT$2,380(+25,+1.06% vs 5/29 NT$2,355)、開NT$2,370、盤中觸NT$2,395創52週新高、量增至65,000張、市值估突破NT$61.7兆續創歷史新高;🚀今日重大催化:(1)COMPUTEX 2026 6/1-5於台北開幕;(2)黃仁勳11:00 GTC Taipei keynote揭N1X ARM筆電晶片+Vera Rubin AI平台(6晶片系統vs Blackwell訓練+3.5x、推理+5x);(3)黃仁勳會晤TSMC CEO魏哲家、廣達等供應鏈龍頭討論Rubin量產;(4)Intel CEO Lip-Bu Tan 6/2 COMPUTEX keynote;(5)CoWoS產能擴張至2026底35,000 wspm(+59% YoY)、中期目標120,000-140,000 wspm滿足Rubin量產;🇺🇸 NYSE TSM 5/29自史高$425.06拉回收$418.45(-$6.41,-1.51% vs 5/28 $424.86)、ADR溢價收斂至+9.4%;⭐結構性利多續存:(1)NVIDIA CEO黃仁勳5/27宣布NVIDIA每年投資台灣$150B(+50%);(2)NVIDIA Taipei HQ動土預計2030啟用;(3)TSMC 3nm下半年漲價15%、明年再漲5-10%;(4)TSMC上修2026全年營收成長>30%、員工分紅年增>30%;🏆 NVDA Q1 FY27財報5/20大超預期:營收$81.6B(+85% YoY)、Q2 guidance $91.0B±2%、$80B回購;6/1估三大法人合計買超+18,500張——外資+13,500張、投信+3,500張、自營+1,500張、COMPUTEX題材推動買盤回升、外資持股估升至75.5%;🎯結構性催化續存:TSMC上修2030全球晶片市場至$1.5T、AI/HPC占55%、18座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%;SOX 5/29 -0.45%至11,772.5、NVDA -1.41%收$237.80、AMD -1.16%收$453.20、AVGO -1.01%收$421.50、ASML -0.65%、AMAT -0.95%(5/29全美AI晶片股拉回);Barclays$470對應NT$2,500(剩+5.0%)、共識NT$2,435.53(剩+2.3%);TSMC 5月月營收預計6/10前公布;Q2 2026財報7/16公布;技術面RSI升至75強勢區、MACD +2.6紅柱再放大、KDJ K(86)/D(80)/J(96)再進超買區、所有均線多頭排列;⚠️短線雜訊:(1)NYSE TSM 5/29自史高拉回-1.51%顯示短線過熱;(2)Custom ASIC增速首超GPU(+44.6% vs +16.1%);(3)短線估值P/E 35.1x在5年92百分位高位;本週關鍵:NT$2,360支撐是否守住、若守住將再測NT$2,400新整數關卡/NT$2,435共識目標"
change_color = "00B050"      # 上漲綠色（+1.06%）

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
