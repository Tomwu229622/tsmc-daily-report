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

today_str = "2026-06-10"
tw_price = "NT$2,305"          # 6/9 Tue 實際收盤（最新確認資料；今日 6/10 Wed 盤中、月營收公布為短線方向定錨）
change_pct = "+0.44%"          # 6/9 收盤漲跌幅（+10 vs 6/8 NT$2,295）
nyse_price = "US$427.89"       # NYSE TSM 最新確認收盤（6/9 Tue，+0.26% vs 6/8 $426.80；高位整理）
volume = "33,206"              # 6/9 實際成交量（張，量縮觀望 6/10 月營收公布）
news_summary = "⭐今日頭條：TSMC 5月合併月營收6/10 13:30公布NT$3,205.16億(NT$320.52B)：年增+39.6%(+39.59%)、月減-8.3%(季節性回落vs 4月史高NT$3,495.67億)、創史上單月次高；1-5月累計NT$1兆5,093.34億(NT$1,509.34B)、年增+42.6%創歷年同期新高。解讀：MoM-8.3%屬4月先進製程拉貨高峰後季節性回落，YoY近+40%與累計+42.6%強勁年增、印證AI晶片需求結構性多頭、與CEO魏哲家6/4 AGM『AI需求超供應』及2026 >30%成長指引一致；單月次高反映N3/N2先進製程滿載+CoWoS瓶頸下強勁出貨。✅6/9 Tue台股2330反彈站穩實際收NT$2,305(+10,+0.44% vs 6/8 NT$2,295)、盤中NT$2,295-2,320量縮至33,206張(月營收公布前觀望)、成交額NT$765.81億、市值守穩NT$59.8兆；加權指數6/9大漲+1,201.66點重返44K收44,704.44確認6/5-6/8賣壓結束。NYSE TSM 6/9高位整理收$427.89(+1.09,+0.26% vs 6/8 $426.80)、盤中$405.60-$438.16。🏆6/9 PwC報告:TSMC市值年增+101%至$1.427兆、躋身全球第9大市值企業(自第12名躍升)。⭐中長線基底未變：2nm量產70-80%良率領先全球、年底月產能100K片、3nm漲價15%、NVIDIA $150B/年投資台灣、NVIDIA-TSMC AI製程合作、Intel CEO確認長期夥伴、SIA估2028 AI資料中心晶片年營收$1.2兆。基本面：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(上行+9.76%)。⚠️估值警示TSM P/E ~33.9x仍在5年88-90%高位。下一里程碑：6月營收7月上旬、Q2法說會7/16。"
change_color = "00B050"        # 綠色（6/9 +0.44% 反彈上漲）

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
