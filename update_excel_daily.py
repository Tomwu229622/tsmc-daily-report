"""
TSMC Excel 每日更新腳本 - 2026-06-10
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

today_str = "2026-06-12"
tw_price = "NT$2,300"          # 6/12 Fri 盤中估（+50，+2.22% vs 6/11 收 NT$2,250；跟進 NYSE TSM 6/11 +3.20% 反彈；官方收盤待確認）
change_pct = "+2.22%"          # 6/12 盤中估漲跌幅（vs 6/11 收 NT$2,250）
nyse_price = "US$421.07"       # NYSE TSM 6/11 Thu 實際收盤（+$12.32，+3.20% vs 6/10 $408.75；強勢反彈）
volume = "38,000"              # 6/12 盤中估成交量（張，估）
news_summary = "✅NYSE TSM 6/11強勢反彈+3.20%收$421.07(+$12.32 vs 6/10 $408.75)、消化6/10 -4.48% sell-the-news賣壓、record-revenue基本面(5月營收NT$4,169.75億、+30.1% YoY、創單月新高)獲市場重新確認；台股2330 6/11收NT$2,250(持平)、6/12預期跟進ADR反彈、盤中估~NT$2,300(+2.22%)、市值回升NT$59.6兆。⚠️籌碼面警訊(真實數據修正)：外資6/9-6/11連3日賣超(-15,008/-15,543/-4,904張、累計逾-3.5萬張)，惟6/11賣壓自-15,543大幅收斂至-4,904、投信轉買+600、自營轉買+340、拋售動能趨緩、外資尚未轉買為次階段關鍵。💰TSMC CFO 6/11確認先進製程H2漲價(3nm漲~15%、sub-5nm漲5-10%)；CEO魏哲家重申全球晶片供應數年內持續落後AI需求、Q2月產能160K-175K片仍供不應求。⚠️競爭雜訊：Google向Intel Foundry下單逾300萬顆自研TPU(2028交付)、Intel 6/8大漲逾11%、市場憂部分客戶在TSMC產能緊俏下尋求分散。⭐NVIDIA-TSMC AI製程合作(cuLitho/cuEST/FabTwin)導入晶圓廠提升良率；NVIDIA RTX Spark超晶片秋季登場由TSMC代工。台美ADR溢價(以6/11 ADR+6/12台股盤中估計)+13.69%。⭐中長線基底未變：5月營收NT$4,169.75億+30.1%創單月新高、1-5月累計NT$1.96兆、2nm量產70-80%良率領先全球、4大美國CSP 2026 AI CapEx合計$725B(+77% YoY)、PwC報告TSMC市值躍居全球第9。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+10.00%)。⚠️估值警示TSM P/E ~33.9x仍在5年90%高位、關鍵守住NT$2,250、6/12站回NT$2,295/5MA。下一里程碑：6月營收7月上旬、Q2法說會7/16。"
change_color = "00B050"        # 綠色（6/12 +2.22% 盤中估上漲）

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
