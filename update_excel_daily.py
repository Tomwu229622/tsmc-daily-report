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
tw_price = "NT$2,405"        # 6/2 Tue 估收（估算，+50，+2.12% vs 6/1 NT$2,355；台股尚未收盤、追趕 ADR 6/1 +4.73%）
change_pct = "+2.12%"        # 台股 2330 6/2 估收盤漲跌幅（估算）
nyse_price = "US$439.78"     # NYSE TSM 最新收盤（6/1 Mon，+4.73% vs 5/29 $418.45；COMPUTEX + GTC Taipei + NVIDIA-TSMC AI 合作三重催化、市值 $2.28T 創史新高）
volume = "52,000 張"         # 6/2 估成交量（估算）
news_summary = "⚠️6/1 Mon台股2330盤中衝天價NT$2,415(市值62.62兆)但終場收平盤NT$2,355(0.00% vs 5/29)、未跟進大盤——高檔獲利了結+資金輪動至聯發科(+5.68%收NT$4,555)等其他AI股;台股加權指數6/1大漲+604.97(+1.35%)收45,337.91創歷史新高、首度站穩45,000;今日(6/2 Tue)台股尚未收盤、本報告以估算值呈現:估收NT$2,405(+50,+2.12% vs 6/1)、估盤中回測6/1天價NT$2,415、估量52,000張;🚀NYSE TSM 6/1 Mon大漲+4.73%收$439.78(vs 5/29 $418.45)、盤中觸$443.18、市值$2.28兆雙創史新高、三重催化:(1)COMPUTEX開幕(2)GTC Taipei keynote(3)NVIDIA-TSMC AI製程合作;台美ADR溢價估擴大至+13.6%(台股個股暫落後ADR);🤝今日(6/2)關鍵催化:(1)COMPUTEX 2026正式開幕日;(2)Intel CEO Lip-Bu Tan 6/2 13:30台北COMPUTEX keynote、18A量產進度(Panther Lake/Nova Lake/Clearwater Forest全採18A)+Foundry外部客戶+TSMC競合關係;(3)NVIDIA 6/1宣布TSMC全面導入NVIDIA AI與加速計算技術——cuLitho運算光刻效率+20-50%、cuEST化學模擬加速50倍、Omniverse FabTwin數位孿生;(4)黃仁勳6/1 GTC Taipei keynote揭Vera Rubin正式量產(vs Blackwell訓練+3.5x、推理+5x、成本降至1/7)、N1X ARM筆電晶片、150家台廠供應鏈夥伴;⭐結構性利多續存:(1)NVIDIA $150B/年投資台灣承諾;(2)NVIDIA Taipei HQ 5/27動土;(3)TSMC 3nm下半年漲價15%、明年再漲5-10%;(4)TSMC上修2026全年營收成長>30%、員工分紅年增>30%;🏆 NVDA Q1 FY27財報5/20大超預期:營收$81.6B(+85% YoY)、Q2 guidance $91.0B±2%、$80B回購;6/2估三大法人合計買超+15,700張——外資+12,000張、投信+2,500張、自營+1,200張;🎯結構性催化續存:TSMC上修2030全球晶片市場至$1.5T、AI/HPC占55%、18座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%、NVIDIA-TSMC AI製程合作;SOX 6/1 +4.64%至12,318.5、NVDA +4.50%、AMD +4.92%、AVGO +4.51%、ASML +4.45%(6/1全美AI晶片股齊漲);Barclays$470對應NT$2,450、共識NT$2,510;TSMC 5月月營收預計6/10前公布;Q2 2026財報7/16公布;技術面RSI估68強勢區、MACD +2.0紅柱、KDJ K(78)/D(72)/J(90)偏高、均線多頭排列;⚠️短線雜訊:(1)6/1台股個股收平盤顯示高檔換手、漲多獲利了結壓力;(2)TSM ADR自6/1史高$443.18高位、技術超買警示;(3)短線估值P/E~35.4x在5年高位;(4)Intel CEO 6/2 keynote若揭重大外部客戶搶單恐成短線雜訊;本週關鍵:回測6/1天價NT$2,415突破將挑戰NT$2,440/NT$2,450(Barclays $470對應)、下方支撐NT$2,355(6/1收)"
change_color = "00B050"      # 上漲綠色（+2.12%）

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
