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
tw_price = "NT$2,255"          # 6/10 Wed 實際收盤（最新確認資料；今日 6/11 Thu 盤中、預料跟進 NYSE 重挫開低）
change_pct = "-2.17%"          # 6/10 收盤漲跌幅（-50 vs 6/9 NT$2,305）
nyse_price = "US$408.75"       # NYSE TSM 最新確認收盤（6/10 Wed，-4.47% vs 6/9 $427.89；利多出盡重挫）
volume = "42,439"              # 6/10 實際成交量（張，放量下跌 vs 6/9 33,206）
news_summary = "⭐今日頭條：TSMC 5月合併月營收6/10正式公布NT$4,169.75億(NT$416.975億元、US$13.17B)：年增+30.1%、月增+1.5%、創史上單月新高(超越3月與4月前高)；1-5月累計NT$1.96兆、年增+30%；分析師認為公司well-positioned達成Q2指引($39.0-40.2B、+10% QoQ、+32% YoY)、可望挑戰單季營收新高(超越前高NT$1.23兆)。【更正】前報5月營收NT$320.52B(NT$3,205億)為公布前誤植、官方正式數字為NT$416.975億元。⚠️然而利多出盡：6/10 Wed台股2330實際收NT$2,255(-50,-2.17% vs 6/9 NT$2,305)、放量至42,439張、市值回落NT$58.5兆；NYSE TSM 6/10重挫-4.47%收$408.75(-$19.14)、盤中$407.70-$426.32——sell-the-news+半導體類股回檔共振+6/3史高NT$2,440後高位獲利了結延續。台美ADR溢價自+14.98%收斂至+12.57%。💰TSMC CFO 6/10暗示先進製程續漲價(3nm H2漲15%、sub-5nm漲5-10%)支撐毛利率。🏆PwC報告:TSMC市值年增+101%躋身全球第9大市值企業。⭐中長線基底未變：2nm量產70-80%良率領先全球、年底月產能100K片、AGM CEO魏哲家確認AI需求超供應、NVIDIA $150B/年投資台灣、Intel CEO確認長期夥伴、SIA估2028 AI資料中心晶片年營收$1.2兆。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+12.20%)。⚠️估值警示TSM P/E ~33.8x仍在5年88-90%高位、短線測試NT$2,200心理整數。下一里程碑：6月營收7月上旬、Q2法說會7/16。"
change_color = "FF0000"        # 紅色（6/10 -2.17% 利多出盡下跌）

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
