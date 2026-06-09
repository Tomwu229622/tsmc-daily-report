"""
TSMC Excel 每日更新腳本 - 2026-06-09
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

today_str = "2026-06-09"
tw_price = "盤中、待盤後補"    # 6/9 Tue 台股盤中、6/9 收盤實際資料待 TWSE 盤後公告
change_pct = "盤中、待盤後補"  # 6/9 收盤漲跌幅待 TWSE 盤後公告
nyse_price = "US$426.80"       # NYSE TSM 最新確認收盤（6/8 Mon，+2.80% vs 6/5 $415.17；自崩跌反彈領先止跌）
volume = "盤中、待盤後補"      # 6/9 成交量待 TWSE 盤後公告
news_summary = "📊今日(6/9 Tue)台股2330盤中、6/9收盤實際資料待TWSE盤後公告。最新確認資料：(1)6/8 Mon 台股 2330 跳空補跌實際收 NT$2,305(-60,-2.54% vs 6/5 NT$2,365)反映NYSE 6/5 -6.69%崩跌、開盤跳空-135點-5.70%至NT$2,230、盤中區間NT$2,230-2,320、量約52,000張、市值回落NT$59.8兆；(2)✅跳空後盤中快速回升+75點顯示低接買盤湧現、跌幅-2.54%明顯低於NYSE 6/5崩跌-6.69%；(3)✅NYSE TSM 6/8反彈+2.80%收$426.80(+$11.63 vs 6/5 $415.17)消化Broadcom利空、領先確認止跌訊號；(4)SOX 6/8反彈+3.05%自12,431.25回升至~12,810、NVDA +2.20%、AVGO +2.68%、AMD +2.63%、ASML +3.41%；(5)6/8估三大法人合計賣超-20,800張(外資-18,500、投信-1,500、自營-800)、台美ADR溢價自+9.04%大幅擴大至+14.99%。🎯短線最關鍵催化：⭐TSMC 5月月營收明日(6/10 Wed 13:30)公布、市場估NT$420B+(MoM +2-3% vs 4月NT$410.73B)、若優於市場估為止跌反彈關鍵訊號；若不及預估恐加深下行壓力。⭐中長線結構性利多基底未變：🏛️6/4 AGM CEO魏哲家確認AI需求超供應、2026 >30%成長、3nm漲價15%；⭐TSMC 2nm開始量產、初期良率70-80%領先全球、年底月產能100K片；NVIDIA Computex 2026發表RTX Spark超晶片由TSMC代工；SIA估2028 AI資料中心晶片年營收$1.2兆(+10x vs 2024)；NVIDIA $150B/年投資台灣；NVIDIA-TSMC AI製程合作(cuLitho/cuEST/FabTwin)；Intel CEO確認TSMC長期夥伴。基本面未變：Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25。32位分析師全數看多、平均目標NT$2,530(自NT$2,305上行+9.76%)。⚠️估值警示TSM P/E ~33.9x仍在5年88-90%高位；TSMC對前主管Wei-Jen Lo訴訟持續中。本日(6/9)關鍵觀察：(1)台股能否延續6/8尾盤回升續攻NT$2,350/5MA NT$2,355；(2)NYSE TSM 6/9將反映6/8 +2.80%反彈是否延續；(3)6/10月營收公布前夕量縮觀望概率高；支撐NT$2,300/NT$2,250心理整數/NT$2,200。6/9收盤實際資料待TWSE盤後公告補資料。"
change_color = "808080"        # 灰色（待盤後補實際資料）

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
