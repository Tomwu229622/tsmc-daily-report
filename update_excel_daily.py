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

today_str = "2026-06-15"
tw_price = "NT$2,310"          # 6/15 Mon 報告日；採 6/12 Fri 實際收盤 NT$2,310（6/15 台股盤中、官方收盤待確認）
change_pct = "盤中待確認"        # 6/15 盤中待確認（最新確認為 6/12 收 +2.67%）
nyse_price = "US$423.93"       # NYSE TSM 6/12 Fri 收盤（週末休市、為最新確認數據）
volume = "—"                   # 6/15 盤中待確認
news_summary = "報告日2026-06-15(Mon)；NYSE週末休市、最新確認為6/12收$423.93(+0.68%)連2紅；台股2330最新確認6/12收NT$2,310(+2.67%)、6/15盤中官方收盤待確認。✅TSMC反彈延續、台美雙市6/12連2日收紅、加權指數+1,019.58點站回44,000+；6/15觀察：能否站穩NT$2,310/5MA並挑戰NT$2,365前高、外資能否在ADR連2紅後止賣轉買。💡重大技術進展：分析師郭明錤6/10揭露TSMC CoPoS(Chip-on-Panel-on-Structure)玻璃基板封裝規劃2028 H2量產、以310×310mm玻璃載板、支援超光罩9.5倍超大封裝、成本與良率優於CoWoS、NVIDIA Feynman次世代AI晶片潛在首批採用、解決CoWoS尺寸瓶頸。💰TSMC CFO確認先進製程H2漲價(3nm漲~15%、sub-5nm漲5-10%)；CEO魏哲家重申全球晶片供應數年內持續落後AI需求、Q2月產能160K-175K片仍供不應求。⚠️競爭雜訊：Google向Intel Foundry下單逾300萬顆自研TPU(2028交付)、Samsung推進HPB封裝對標CoPoS、Intel 6/8大漲逾11%。台美ADR溢價(以6/12 ADR$423.93+6/12台股NT$2,310計)+13.97%。⭐中長線基底未變：5月營收NT$4,169.75億+30.1%創單月新高、1-5月累計NT$1.96兆、2nm量產70-80%良率領先全球、4大美國CSP 2026 AI CapEx合計$725B(+77% YoY)、PwC報告TSMC市值躍居全球第9。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+9.52%)。⚠️估值警示TSM P/E ~35.1x仍在5年高位、關鍵守住NT$2,250。下一里程碑：6月營收7月上旬、Q2法說會7/16。"
change_color = "808080"        # 灰色（6/15 盤中待確認）

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
