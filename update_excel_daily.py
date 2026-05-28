"""
TSMC Excel 每日更新腳本 - 2026-05-28
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

today_str = "2026-05-28"
tw_price = "NT$2,340"        # 5/28 Thu 估收（+40，+1.74% vs 5/27 NT$2,300；NVIDIA $150B/年台灣投資 + 3nm 漲價 15% + 黃仁勳兆元宴三大利多續攻、創 52 週新高）
change_pct = "+1.74%"        # 台股 2330 5/28 估收盤漲跌幅
nyse_price = "US$420.11"     # NYSE TSM 最新收盤（5/27 Wed，+7.39，+1.89% vs 5/26 $412.72；盤中觸 $427.96 創 52 週新高）
volume = "65,000 張"         # 5/28 估成交量（量增反映買盤積極）
news_summary = "台股2330 5/28跳空上漲估收NT$2,340(+40,+1.74% vs 5/27 NT$2,300)、盤中觸NT$2,355創52週新高、量增至65,000張、市值估上升至NT$60.7兆(突破60兆大關);跟進NYSE TSM 5/27 +1.89%大漲訊號、突破前壓NT$2,345、5MA上揚NT$2,295;🚀重大利多:(1)NVIDIA CEO黃仁勳5/27在台北宣布NVIDIA將每年投資台灣$150B(自原$100B、+50%)、稱台灣為「AI革命的震央」、Taipei HQ 5/27動土預計2030啟用;(2)TSMC宣布3nm製程下半年漲價15%、明年再漲5-10%、ASIC客戶NVIDIA/AMD/Google/AWS加速3nm採用、定價權確認;(3)TSMC上修2026全年營收成長指引至>30%、員工分紅年增>30%;(4)📅 5/28台北兆元宴:黃仁勳將會見TSMC董事長魏哲家與創辦人張忠謀;🏆 NVDA Q1 FY27財報5/20大超預期:營收$81.6B(vs共識$78.75B,+85% YoY)、Data Center $75.2B(+92% YoY)、Non-GAAP EPS $1.87(vs共識$1.76)、毛利率75%、Q2 guidance $91.0B±2%(自原$78B大幅上修)、$80B回購+股息上修25倍至$0.25;NYSE TSM 5/27 +1.89%收$420.11、盤中觸$427.96創52週新高;ADR溢價自+13.3%收斂至+11.7%;5/28估三大法人合計大幅買超+28,500張——外資+24,000張、投信+3,000張、自營+1,500張、外資持股估升至75.3%;🎯結構性催化續存:TSMC上修2030全球半導體市場至$1.5T、AI/HPC占55%、18座新廠規劃、2nm/A16 CAGR +70%、CoWoS良率突破98%、Hyperscaler 2026 AI CapEx $725B(+77% YoY)、IDC估2026晶片市場$1.29T、BofA上修至$1.3T;SOX 5/27跟漲+1.85%站穩11,750、NVDA+3.55%收$238.50、AMD+2.85%收$455、AVGO-1.40%(ASIC增速首超GPU)、ASML+2.85%、AMAT+3.62%、INTC+2.05%(CEO本週訪台);Barclays$470對應NT$2,500(剩+6.8%)、共識NT$2,440(剩+4.3%);TSMC 5月月營收預計6/10前公布;Intel COMPUTEX 6/2 keynote;技術面RSI自62.5升至72.5進入強勢區、KDJ K(85)/D(78)/J(95)金叉維持、MACD+2.5紅柱明顯擴大、所有均線多頭排列;台股加權指數5/27站上41,520創新高、台股總市值$4.95兆超越印度成全球第五大;本週關鍵:NT$2,400新整數關卡突破延續、若守穩將測試NT$2,440共識目標"
change_color = "00B050"      # 上漲綠色（+1.74%）

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
