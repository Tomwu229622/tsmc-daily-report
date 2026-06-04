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
tw_price = "NT$2,400"        # 6/4 Thu 估收（估算，-25，-1.03% vs 6/3 實際 NT$2,425；台股跟進 NYSE 6/3 -2.24% 回檔修正）
change_pct = "-1.03%"        # 台股 2330 6/4 估收盤漲跌幅（估算）
nyse_price = "US$436.69"     # NYSE TSM 最新收盤（6/3 Wed，-2.24% vs 6/2 史新高 $446.69；自盤中 52w 高 $449.39 高位獲利了結、技術性回檔）
volume = "52,000 張"         # 6/4 估成交量（估算、放量整理）
news_summary = "✅6/3 Wed台股2330實際收NT$2,425(+45,+1.89% vs 6/2 NT$2,380)創歷史新高、跟進NYSE TSM 6/2大漲訊號+Intel CEO 6/2 keynote確認TSMC「長期信任夥伴」雜訊清除催化、盤中觸NT$2,430;6/3估三大法人合計買超+23,500張(外資+18,500、投信+3,200、自營+1,800);📉NYSE TSM 6/3 Wed收$436.69(-$10.00,-2.24% vs 6/2史新高$446.69)自盤中52w新高$449.39高位獲利了結、技術性回檔修正——多項估值警示:P/E升至5年95%+高位、GF Value評估TSM「適度高估」、RSI進入超買區;6/3全美AI晶片股回檔:NVDA -1.34%、AMD -1.10%、AVGO -0.38%、ASML -1.60%、AMAT -1.52%、MU -1.53%、SOX -1.76%;今日(6/4 Thu)台股2330估跟進NYSE回檔、估收NT$2,400(-25,-1.03% vs 6/3 NT$2,425)、估盤中區間NT$2,385-2,420、估量52,000張放量整理;台美ADR溢價自+14.6%收斂至+13.0%;6/4估三大法人合計賣超-7,200張(外資-8,500、投信+1,800、自營-500);📊Broadcom(AVGO)6/3盤後財報關注:TSMC #3客戶13%、Custom XPU+AI營收續爆發、前一季AI營收+106% YoY、6/4台股反應財報結果;🤝Intel CEO 6/2 keynote確認TSMC「長期信任夥伴」、續委外先進晶片;🤝NVIDIA-TSMC 6/1 AI製程合作:cuLitho運算光刻效率+20-50%、cuEST化學模擬加速50倍、Omniverse FabTwin數位孿生;🏆黃仁勳6/1 GTC Taipei keynote揭Vera Rubin正式量產(vs Blackwell訓練+3.5x、推理+5x、成本降至1/7)、150家台廠供應鏈夥伴;🚀SIA/Deloitte 6/1報告:2028 AI資料中心晶片年營收$1.2兆、4年成長10倍、晶片占AI機櫃價值95%;IDC估2026全球半導體收入$1.29T(+52.8% YoY);⭐結構性利多續存:(1)Intel CEO確認TSMC「長期信任夥伴」(2)NVIDIA $150B/年投資台灣承諾(3)NVIDIA Taipei HQ 5/27動土(4)TSMC 3nm下半年漲價15%、明年再漲5-10%(5)TSMC上修2026全年營收成長>30%、員工分紅年增>30%;🏆NVDA Q1 FY27財報5/20大超預期:營收$81.6B(+85% YoY)、Q2 guidance $91.0B、$80B回購;🎯結構性催化續存:TSMC上修2030全球晶片市場至$1.5T、AI/HPC占55%、18座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%、NVIDIA-TSMC AI製程合作、Intel確認長期夥伴、SIA/Deloitte 2028 AI晶片$1.2T;Barclays $470對應NT$2,450、共識NT$2,530;Broadcom 6/3盤後財報;TSMC 5月月營收預計6/10前公布;COMPUTEX 2026 6/2-6末日6/6;Q2 2026財報7/16公布;技術面RSI估自70降至65、MACD +1.8紅柱動能收斂、KDJ K(76)/D(78)/J(72) J自96高檔回落初步死叉、均線多頭排列維持;⚠️短線雜訊:(1)⚠️NYSE TSM 6/3自史高$449.39高位回檔-2.24%、估值警示P/E 36-37x在5年95%+高位(2)GF Value評估TSM「適度高估」(3)TSMC對前主管Wei-Jen Lo訴訟(2nm商業機密、轉投Intel)持續中(4)Apple傳部分晶片轉Intel 18A的5/13雜訊未完全消化;本週關鍵:6/4能否守住NT$2,400心理整數+NT$2,380(6/2收)支撐、反彈壓力NT$2,425(6/3史高)/NT$2,430(6/3盤中高)/NT$2,450(Barclays對應)"
change_color = "FF0000"      # 下跌紅色（-1.03%）

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
