"""
TSMC Excel 每日更新腳本 - 2026-07-01
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

today_str = "2026-07-15"
tw_price = "NT$2,420（7/14收）"   # 台股2330 7/14 Tue官方收NT$2,420(-20、-0.82% vs 7/13 NT$2,440、開2,410/高2,430/低2,390、量42,857張TWSE口徑、成交額1,033.71億、均價2,412、TWSE)——大盤重挫-642.57點(-1.42%)下相對抗跌、惟失守5MA、法說會前觀望
change_pct = "-0.82%（7/14收）"    # 7/14官方漲跌幅-0.82%（TWSE、法說會前觀望+獲利了結）
nyse_price = "US$420.39（7/14收）" # NYSE TSM 7/14 Tue收$420.39(-$1.19、-0.28% vs 7/13 $421.58、日內418.86-430.87、Yahoo)法說會前連3日小跌；同日SOX +2.54%大反彈、TSM明顯落後屬觀望
volume = "42,857（7/14收）"    # 台股7/14官方成交量42,857張（TWSE口徑）、成交額1,033.71億
news_summary = "報告日2026-07-15(Wed)。⭐⭐7/16 Q2法說會倒數1天(週三14:00)：法人估Q2營收~$40B(+32% YoY)、獲利+50%+連5季創高、毛利率挑戰70%(指引65.5-67.5%)、關注上修全年財測(現行>30%)；6月營收NT$4,426.80億(+67.9% YoY)已確認Q2約NT$1.27兆達標。📉台股2330 7/14收NT$2,420(-20、-0.82%、量42,857張TWSE口徑、額1,033.71億)——大盤重挫-642.57點(-1.42%)收44,737.95、相對抗跌惟失守5MA~2,436。📊三大法人7/14(T86)合計賣超8,005張：外資-12,416(連4賣、4日-31,365)、投信+1,841、自營+2,570。📉NYSE TSM 7/14收$420.39(-0.28%)法說會前觀望；⚡同日SOX +2.54%大反彈(NVDA +4.06%、MU +4.92%、AMKR +6.27%)、SK海力士7/15韓股盤中+11%，今日台股可望跟進。💰漲價題材：7nm以下全製程調漲5-10%、三星跟進4/5nm漲15%；CoWoS 2026產能12.5萬片/月遭NVIDIA/博通訂滿、缺口至2027。📊ADR溢價+11.7%($420.39 vs理論值$376.46、匯率32.14)。🚀分析師續挺：Citi NT$3,800(+57.0%)、Barclays $625(+48.7%)、Nomura NT$3,425(+41.5%)。⚠️風險：外資連4賣擴大、法說會前後劇烈波動、232關稅輸中晶片查驗收緊、Intel 18A、估值不低(P/E~32.1x)。技術面：MACD柱轉負、KDJ J=22近超賣；支撐2,390/2,366/2,336、壓力2,436/2,480/2,535史高。"
change_color = "FF0000"        # 紅色（7/14官方-0.82%下跌）

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
