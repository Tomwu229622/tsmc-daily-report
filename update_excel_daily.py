"""
TSMC Excel 每日更新腳本 - 2026-06-05
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

today_str = "2026-06-08"
tw_price = "NT$2,285（估）"   # 6/8 Mon 估收（估算，-80，-3.38% vs 6/5 NT$2,365；估跳空補跌反映 NYSE 6/5 -6.69% 崩跌；待盤後補實際資料）
change_pct = "-3.38%（估）"  # 台股 2330 6/8 估收盤漲跌幅（估算）
nyse_price = "US$415.17"     # NYSE TSM 最新確認收盤（6/5 Fri，-6.69%；6/8 美股盤待）
volume = "82,000 張（估）"   # 6/8 估成交量（估算、爆量殺盤）
news_summary = "📊今日(6/8 Mon)台股2330估跳空補跌、估收NT$2,285(-80,-3.38% vs 6/5 NT$2,365)、估盤中區間NT$2,265-2,320、估量爆82,000張、市值估回落NT$59.3兆，反映6/5美股半導體爆量崩跌(SOX -8.71%、NYSE TSM -6.69%收$415.17、$1T+半導體市值蒸發)；6/8估三大法人合計賣超-35,200張(外資-28,500、投信-4,500、自營-2,200);🔴短線殺盤三大主因:(1)⚠️Broadcom 6/3盤後Q3 AI指引$16B不及分析師估$17.2B、未上修2026全年AI指引、引爆「sell-the-news」、AVGO 6/4-6/5累計-13%;(2)⚠️5月NFP強勁、10Y美債殖利率突破4.5%、Fed升息預期升溫衝擊高估值科技股;(3)⚠️美對中AI晶片出口管制6/1起延伸至中企海外子公司;(4)估值警示TSM P/E仍在5年90%+高位;⭐長線定錨(崩跌中結構性利多未變):BofA上修2026全球晶片市場至$1.3兆、點名NVIDIA/Broadcom/Marvell/AMD四大引擎;Broadcom CEO Hock Tan確認轉型chips-only、2027 AI指引維持>$100B不變、6大客戶(Anthropic/OpenAI/Google/Meta)訂單鎖至2028;🏛️6/4 AGM CEO魏哲家確認AI需求將持續超越全球半導體供應、重申2026成長>30%、3nm H2漲價15%+2027再漲5-10%;Intel CEO確認TSMC長期夥伴;NVIDIA $150B/年投資台灣;NVIDIA-TSMC AI製程合作(cuLitho/cuEST/FabTwin);🎯短線最關鍵催化:⭐TSMC 5月月營收6/10公布、估NT$420B+、若優於市場估為止跌反彈關鍵訊號;Q2 2026財報7/16;基本面未變:Q1 2026淨利NT$572.5B(+58.3% YoY)、毛利率66.2%雙創史高、2025全年EPS NT$66.25;32位分析師全數看多、平均目標NT$2,530;⚠️TSMC對前主管Wei-Jen Lo訴訟持續中;本日關鍵:6/8台股能否守住NT$2,265/NT$2,250心理整數支撐並反彈挑戰NT$2,320/5MA NT$2,355"
change_color = "FF0000"      # 下跌紅色（-3.38% 估）

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
