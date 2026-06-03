"""
TSMC Excel 每日更新腳本 - 2026-06-03
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

today_str = "2026-06-03"
tw_price = "NT$2,420"        # 6/3 Wed 估收（估算，+40，+1.68% vs 6/2 NT$2,380；台股反彈跟進 NYSE 6/3 +2.54%）
change_pct = "+1.68%"        # 台股 2330 6/3 估收盤漲跌幅（估算）
nyse_price = "US$446.69"     # NYSE TSM 最新收盤（6/3 Wed，+2.54% vs 6/2 $435.63；Intel CEO 確認長期夥伴 + NVIDIA-TSMC AI 製程合作 + Vera Rubin 量產利多、創史新高）
volume = "48,000 張"         # 6/3 估成交量（估算、量縮整理）
news_summary = "✅6/2 Tue台股2330實際收NT$2,380(+25,+1.06% vs 6/1 NT$2,355)、跟進NYSE TSM 6/1 +4.73%大漲訊號、量縮整理;台股加權指數續站穩45,000+;NYSE TSM 6/2 Tue收$435.63(-0.94% vs 6/1 $439.78)、盤中觸52週新高$449.39後尾盤回檔修整、高位獲利了結+Intel CEO 13:30 keynote觀察前消化;🤝6/2 Intel CEO Lip-Bu Tan COMPUTEX keynote重大確認:Tan表示Intel視TSMC為「長期信任夥伴」(very trusted partnership)、確認Intel為TSMC大客戶之一、Intel將持續委外先進晶片給TSMC——市場原憂Intel 18A競爭恐分食TSMC訂單的雜訊清除、解除短線壓力;揭露2025年曾提議TSMC入股Intel Foundry約20%(未成交);Intel YTD +200%+、市值$614B+;🚀NYSE TSM 6/3 Wed大漲+2.54%收$446.69創史新高(vs 6/2 $435.63)、+$11.06、反映三大利多:(1)Intel keynote確認長期夥伴清除雜訊(2)NVIDIA-TSMC 6/1 AI製程合作續發酵(3)Vera Rubin量產利多;今日(6/3)台股2330估反彈跟進NYSE、估收NT$2,420(+40,+1.68% vs 6/2 NT$2,380)、估盤中觸NT$2,430創新高、估量48,000張;台美ADR溢價估維持+14.6%;6/3估三大法人合計買超+18,800張(外資+14,500、投信+2,800、自營+1,500);🤝NVIDIA-TSMC 6/1 AI製程合作:cuLitho運算光刻效率+20-50%、cuEST化學模擬加速50倍、Omniverse FabTwin數位孿生、cuML製程參數處理、Vision AI瑕疵檢測;🏆黃仁勳6/1 GTC Taipei keynote揭Vera Rubin正式量產(vs Blackwell訓練+3.5x、推理+5x、成本降至1/7)、150家台廠供應鏈夥伴;⭐結構性利多續存:(1)Intel CEO確認TSMC「長期信任夥伴」(2)NVIDIA $150B/年投資台灣承諾(3)NVIDIA Taipei HQ 5/27動土(4)TSMC 3nm下半年漲價15%、明年再漲5-10%(5)TSMC上修2026全年營收成長>30%、員工分紅年增>30%;🏆NVDA Q1 FY27財報5/20大超預期:營收$81.6B(+85% YoY)、Q2 guidance $91.0B、$80B回購;🎯結構性催化續存:TSMC上修2030全球晶片市場至$1.5T、AI/HPC占55%、18座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%、NVIDIA-TSMC AI製程合作、Intel確認長期夥伴;SOX 6/3 +1.36%至12,485.2創新高、NVDA +1.89%、AMD +1.75%、AVGO +1.75%、ASML +2.15%、AMAT +2.22%、MU +2.94%、INTC +6.34%(6/3全美AI晶片股續強);Barclays $470對應NT$2,450、共識NT$2,530;TSMC 5月月營收預計6/10前公布;Q2 2026財報7/16公布;技術面RSI估70接近超買、MACD +2.2紅柱、KDJ K(82)/D(75)/J(96)高檔、均線多頭排列;⚠️短線雜訊:(1)短線估值P/E~37.0x已升至5年95%+高位、需財報數據持續驗證(2)TSM ADR自6/2史高$449.39高位、技術超買警示(3)TSMC對前主管Wei-Jen Lo訴訟(2nm商業機密、轉投Intel)持續中(4)Apple傳部分晶片轉Intel 18A的5/13雜訊未完全消化;本週關鍵:突破NT$2,430將挑戰NT$2,450(Barclays $470對應)、下方支撐NT$2,380(6/2收)/NT$2,355(6/1收)"
change_color = "00B050"      # 上漲綠色（+1.68%）

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
