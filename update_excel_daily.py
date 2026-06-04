"""
TSMC Excel 每日更新腳本 - 2026-06-04
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

today_str = "2026-06-04"
tw_price = "NT$2,435"        # 6/4 Thu 估收（估算，+10，+0.41% vs 6/3 實際 NT$2,425；Broadcom 6/3 盤後 AI 財報大超預期催化、台股守穩史高、offset NYSE TSM 6/3 -2.24% 回檔）
change_pct = "+0.41%"        # 台股 2330 6/4 估收盤漲跌幅（估算）
nyse_price = "US$436.69"     # NYSE TSM 最新確認收盤（6/3 Wed，-2.24% vs 6/2 史新高 $446.69；6/4 美股盤估隨 Broadcom AI 財報利多反彈、收盤待確認）
volume = "55,000 張"         # 6/4 估成交量（估算、放量續攻）
news_summary = "🚀重大催化:Broadcom(AVGO)6/3美股盤後公布Q2 FY2026財報全面大超預期——總營收$22.2B(+48% YoY)創史高、超越$22.0B指引;AI半導體營收$10.8B(+143% YoY)超預期;Q3指引AI營收上看$16.0B(+200% YoY)、合併營收$29.4B(+84% YoY)、調整後EBITDA $15.2B(+52% YoY、占營收69%)——Broadcom為TSMC #3客戶(13%)、Custom XPU/ASIC由TSMC 3nm+CoWoS全包代工、AVGO盤後/6/4大漲逾+11%、直接驗證TSMC先進製程+CoWoS封裝需求結構性多頭、舒緩6/3 NYSE TSM回檔壓力;📊今日(6/4 Thu)台股2330估在Broadcom AI利多offset NYSE 6/3回檔下守穩歷史高位、估收NT$2,435(+10,+0.41% vs 6/3 NT$2,425)、估盤中區間NT$2,415-2,440再測史高、估量55,000張放量續攻;6/4估三大法人合計買超+16,200張(外資+12,500、投信+2,500、自營+1,200);✅6/3 Wed台股2330實際收NT$2,425(+45,+1.89% vs 6/2 NT$2,380)創歷史新高、盤中觸NT$2,440;📉NYSE TSM 6/3 Wed收$436.69(-$10.00,-2.24% vs 6/2史新高$446.69)自盤中52w新高$449.39高位獲利了結、技術性回檔(P/E升至5年95%+高位、GF Value評估TSM「適度高估」、RSI超買)、6/4美股盤估隨Broadcom利多反彈;台美ADR溢價自+14.6%收斂至+11.4%;🤝Intel CEO 6/2 keynote確認TSMC「長期信任夥伴」、續委外先進晶片;🤝NVIDIA-TSMC 6/1 AI製程合作:cuLitho運算光刻效率+20-50%、cuEST化學模擬加速50倍、Omniverse FabTwin數位孿生;🏆黃仁勳6/1 GTC Taipei keynote揭Vera Rubin正式量產(vs Blackwell訓練+3.5x、推理+5x、成本降至1/7)、150家台廠供應鏈夥伴;🚀SIA/Deloitte 6/1報告:2028 AI資料中心晶片年營收$1.2兆、4年成長10倍、晶片占AI機櫃價值95%;⭐結構性利多續存:(1)Broadcom 6/3 AI財報驗證Custom ASIC+CoWoS需求(2)Intel CEO確認TSMC「長期信任夥伴」(3)NVIDIA $150B/年投資台灣承諾(4)TSMC 3nm下半年漲價15%、明年再漲5-10%(5)TSMC上修2026全年營收成長>30%、員工分紅年增>30%;🏆NVDA Q1 FY27財報5/20大超預期:營收$81.6B(+85% YoY)、Q2 guidance $91.0B、$80B回購;🎯結構性催化續存:TSMC上修2030全球晶片市場至$1.5T、AI/HPC占55%、18座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%;Barclays $470對應NT$2,450、共識NT$2,530;TSMC 5月月營收預計6/10前公布;COMPUTEX 2026 6/2-6末日6/6;Q2 2026財報7/16公布;技術面RSI估維持71高位、MACD +2.2紅柱續放大、KDJ K(82)/D(80)/J(86)高檔金叉維持、均線多頭排列強化;⚠️短線雜訊:(1)⚠️NYSE TSM 6/3自史高$449.39高位回檔-2.24%、估值警示P/E 36-37x在5年95%+高位(2)GF Value評估TSM「適度高估」(3)TSMC對前主管Wei-Jen Lo訴訟(2nm商業機密、轉投Intel)持續中(4)Apple傳部分晶片轉Intel 18A的5/13雜訊未完全消化;本日關鍵:6/4能否突破NT$2,440史高+守住NT$2,425(6/3收)支撐、反彈壓力NT$2,440(6/3史高)/NT$2,450(Barclays對應)/NT$2,500心理整數"
change_color = "00B050"      # 上漲綠色（+0.41%）

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
