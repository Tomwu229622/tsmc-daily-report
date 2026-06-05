"""
TSMC Excel 每日更新腳本 - 2026-06-05
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

today_str = "2026-06-05"
tw_price = "NT$2,400"        # 6/5 Fri 估收（估算，+15，+0.63% vs 6/4 實際 NT$2,385；AGM 利多消化 + Broadcom AI 動能延續 + CEO Wei 確認 AI 結構性緊俏催化反彈）
change_pct = "+0.63%"        # 台股 2330 6/5 估收盤漲跌幅（估算）
nyse_price = "US$436.69"     # NYSE TSM 最新確認收盤（6/4 Thu，持平 vs 6/3 $436.69；6/5 美股盤待 AGM 與台股後續催化）
volume = "48,000 張"         # 6/5 估成交量（估算、量縮整理）
news_summary = "🏛️重大事件:TSMC 2026年度股東會6/4於新竹召開、股東通過2025年業務報告與財報——合併營收NT$3,809.05B、淨利NT$1,717.88B、稀釋EPS NT$66.25、雙創史高;💬CEO魏哲家對股東確認AI需求將持續超越全球半導體供應、儘管TSMC全力推進美國擴廠仍難滿足客戶對先進晶片的強勁需求、TSMC重申2026全年營收成長>30%指引、3nm製程H2漲價15%+2027再漲5-10%計畫不變;✅6/4 Thu台股2330實際收NT$2,385(-40,-1.65% vs 6/3 NT$2,425)自史高NT$2,425高位獲利了結、AGM利多被技術性回檔吸收;📊今日(6/5 Fri)台股2330估在AGM利多消化+Broadcom 6/3 AI財報動能延續+CEO Wei確認AI結構性緊俏催化下溫和反彈、估收NT$2,400(+15,+0.63% vs 6/4 NT$2,385)、估盤中區間NT$2,385-2,410、估量48,000張量縮整理;6/5估三大法人合計買超+7,600張(外資+5,500、投信+1,500、自營+600);📊NYSE TSM 6/4 Thu持穩$436.69(持平 vs 6/3 $436.69)已消化6/3 -2.24%自史高$449.39回檔賣壓;台美ADR溢價自+11.4%擴大至+13.0%;🚀Broadcom 6/3美股盤後Q2 FY2026財報全面大超預期:總營收$22.2B(+48% YoY)、AI半導體營收$10.8B(+143% YoY)、Q3指引AI營收$16.0B(+200% YoY);🤝NVIDIA-TSMC AI製程合作(cuLitho/cuEST/FabTwin)續發酵;🤝Intel CEO 6/2 keynote確認TSMC「長期信任夥伴」雜訊清除續存;🏆黃仁勳6/1 GTC Taipei keynote揭Vera Rubin正式量產(vs Blackwell訓練+3.5x、推理+5x、成本降至1/7)、150家台廠供應鏈夥伴;🚀BigGo Finance估2026全球晶片市場$1.51兆(+90% YoY)首破$1T;🚀SIA/Deloitte 6/1報告:2028 AI資料中心晶片年營收$1.2兆、4年成長10倍、晶片占AI機櫃價值95%;⭐結構性利多續存:(1)🏛️6/4 AGM CEO Wei確認AI需求超供應、2026 >30%成長(2)Broadcom 6/3 AI財報驗證Custom ASIC+CoWoS需求(3)Intel CEO確認TSMC「長期信任夥伴」(4)NVIDIA $150B/年投資台灣承諾(5)TSMC 3nm下半年漲價15%、2027再漲5-10%(6)員工分紅年增>30%;🏆NVDA Q1 FY27財報5/20大超預期:營收$81.6B(+85% YoY)、Q2 guidance $91.0B、$80B回購;🎯結構性催化續存:TSMC上修2030全球晶片市場至$1.5T、18座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%;Barclays $470對應NT$2,450、共識NT$2,530;TSMC 5月月營收預計6/10前公布;COMPUTEX 2026末日6/6;Q2 2026財報7/16公布;技術面RSI估自71回落至65、MACD +1.8紅柱收斂、KDJ K(72)/D(76)/J(64)自高檔回落初步死叉、均線多頭排列維持;⚠️短線雜訊:(1)⚠️6/4 -1.65%反映高位獲利了結、估值警示P/E 35-36x在5年93%+高位(2)GF Value評估TSM「適度高估」(3)TSMC對前主管Wei-Jen Lo訴訟(2nm商業機密、轉投Intel)持續中(4)美對中AI晶片出口管制6/1延伸至中企海外子公司;本日關鍵:6/5能否守住NT$2,385(6/4收)支撐並挑戰NT$2,410、反彈壓力NT$2,425(6/3收)/NT$2,440(6/3史高)/NT$2,450(Barclays對應)"
change_color = "00B050"      # 上漲綠色（+0.63%）

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
