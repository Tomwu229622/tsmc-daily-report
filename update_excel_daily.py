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

today_str = "2026-06-11"
tw_price = "NT$2,275"          # 6/11 Thu 盤中（+20，+0.89% vs 6/10 NT$2,255；逢低買盤回補、利多消化後反彈）
change_pct = "+0.89%"          # 6/11 盤中漲跌幅（+20 vs 6/10 NT$2,255）
nyse_price = "US$408.75"       # NYSE TSM 6/10 Wed 收盤（US 6/11 尚未開盤、維持 6/10 收盤）
volume = "30,500"              # 6/11 盤中成交量（張，估、量縮整理 vs 6/10 42,439）
news_summary = "✅利多消化後反彈：6/11 Thu台股2330盤中翻紅至NT$2,275(+20,+0.89% vs 6/10 NT$2,255)、量縮整理約30,500張、市值回升NT$59.0兆——昨日6/10利多出盡-2.17% sell-the-news後逢低買盤回補、record-revenue基本面獲確認。⭐今日頭條(續)：TSMC 5月合併月營收NT$4,169.75億(NT$416.975億元、US$13.17B)：年增+30.1%、月增+1.5%、創史上單月新高；1-5月累計NT$1.96兆、年增+30%；Taipei Times 6/11確認record revenue、CEO魏哲家重申「全球半導體供應將數年內持續落後AI需求」；分析師認為公司well-positioned達成Q2指引($39.0-40.2B、+10% QoQ、+32% YoY)、可望挑戰單季營收新高(超越前高NT$1.23兆)。⚠️昨日(6/10)回顧：台股收NT$2,255(-2.17%)放量、NYSE TSM重挫-4.47%收$408.75(-$19.14)；US 6/11尚未開盤、今晚ADR走勢為次日台股關鍵。台美ADR溢價自+12.57%進一步收斂至+11.57%。6/11估三大法人合計轉買超+5,600張(外資+4,500、投信+800、自營+300)、外資持股回升75.1%。💰TSMC CFO暗示先進製程續漲價(3nm H2漲15%、sub-5nm漲5-10%)支撐毛利率。🏆PwC報告:TSMC市值年增+101%躋身全球第9大市值企業。⭐中長線基底未變：2nm量產70-80%良率領先全球、年底月產能100K片、AGM CEO魏哲家確認AI需求超供應、NVIDIA $150B/年投資台灣+RTX Spark由TSMC代工、4大美國CSP 2026 AI CapEx合計$725B(+77% YoY)、Intel CEO確認長期夥伴。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+11.21%)。⚠️估值警示TSM P/E ~33.5x仍在5年88-90%高位、關鍵守住NT$2,250、站回NT$2,295。下一里程碑：6月營收7月上旬、Q2法說會7/16。"
change_color = "00B050"        # 綠色（6/11 +0.89% 盤中翻紅上漲）

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
