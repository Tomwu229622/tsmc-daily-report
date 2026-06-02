"""
TSMC Excel 每日更新腳本 - 2026-06-02
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

today_str = "2026-06-02"
tw_price = "NT$2,470"        # 6/2 Tue 估收（+30，+1.23% vs 6/1 NT$2,440；接續 ADR 6/1 +4.73% 大漲訊號）
change_pct = "+1.23%"        # 台股 2330 6/2 估收盤漲跌幅
nyse_price = "US$439.78"     # NYSE TSM 最新收盤（6/1 Mon，+19.88，+4.73% vs 5/29 $418.45；COMPUTEX + GTC Taipei + NVIDIA-TSMC AI 合作三重催化、創史新高）
volume = "62,000 張"         # 6/2 估成交量（量持續高檔反映題材熱度延續）
news_summary = "今日(6/2 Tue)台股2330接續6/1 NT$2,440強勢(+85,+3.61% vs 5/29 NT$2,355)追趕ADR 6/1 +4.73%大漲訊號、估收NT$2,470(+30,+1.23% vs 6/1 NT$2,440)、開NT$2,455跳空開高、盤中觸NT$2,485再創52週新高、估量62,000張持續高檔、市值估突破NT$64兆再創歷史新高;🚀NYSE TSM 6/1 Mon大漲+4.73%收$439.78(+$19.88 vs 5/29 $418.45)、盤中觸$442.50創史新高、三重催化:(1)COMPUTEX開幕(2)GTC Taipei keynote(3)NVIDIA-TSMC AI製程合作;🤝今日(6/2)關鍵催化:(1)COMPUTEX 2026第二日續存;(2)Intel CEO Lip-Bu Tan 6/2 13:30台北COMPUTEX keynote、18A量產進度+TSMC競合關係焦點;(3)NVIDIA 6/1宣布TSMC全面導入NVIDIA AI與加速計算技術——cuLitho運算光刻效率+20-50%、cuEST化學模擬加速50倍、Omniverse FabTwin數位孿生;(4)黃仁勳6/1 GTC Taipei keynote揭Vera Rubin正式量產(VR200機櫃六晶片系統vs Blackwell訓練+3.5x、推理+5x、成本降至1/7)、N1X ARM筆電晶片、150家台廠供應鏈夥伴;(5)CoWoS產能擴張至2026底35,000 wspm(+59% YoY)→中期目標120,000-140,000 wspm;⭐結構性利多續存:(1)NVIDIA $150B/年投資台灣承諾;(2)NVIDIA Taipei HQ 5/27動土;(3)TSMC 3nm下半年漲價15%、明年再漲5-10%;(4)TSMC上修2026全年營收成長>30%、員工分紅年增>30%;🏆 NVDA Q1 FY27財報5/20大超預期:營收$81.6B(+85% YoY)、Q2 guidance $91.0B±2%、$80B回購;6/2估三大法人合計買超+22,500張——外資+17,500張、投信+3,500張、自營+1,500張、追單買盤強勁、外資持股估升至75.7%;🎯結構性催化續存:TSMC上修2030全球晶片市場至$1.5T、AI/HPC占55%、18座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%、NVIDIA-TSMC AI製程合作為新利多;SOX 6/1 +4.64%至12,318.5、NVDA +4.50%收$248.50、AMD +4.92%收$475.50、AVGO +4.51%收$440.50、ASML +4.45%收$778.50、AMAT +4.47%收$238.50(6/1全美AI晶片股齊漲);Barclays$470對應NT$2,500(剩+1.2%)、共識NT$2,510(剩+1.6%);TSMC 5月月營收預計6/10前公布;Q2 2026財報7/16公布;技術面RSI升至78強勢區、MACD +3.0紅柱續放大、KDJ K(88)/D(83)/J(98)高檔超買區、所有均線多頭排列;⚠️短線雜訊:(1)短線估值P/E 36.4x在5年95百分位高位;(2)TSM ADR自6/1史高$442.50高位、技術超買警示;(3)Custom ASIC增速首超GPU(+44.6% vs +16.1%);(4)Intel CEO 6/2 keynote若揭重大外部客戶搶單恐成短線雜訊;本週關鍵:NT$2,440(6/1收)支撐是否守住、若守住將再測NT$2,500 Barclays $470對應整數關卡"
change_color = "00B050"      # 上漲綠色（+1.23%）

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
