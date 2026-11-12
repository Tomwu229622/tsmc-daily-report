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

today_str = '2026-11-12'
tw_price = 'NT$2,395（11/11收，+110、+4.81%；開2,330（跳空+45）、高2,395、低2,305，收在全日最高、紅K實體65元；2026年99個上漲日第8大，非極端；連3日上漲、10/29以來最高收盤；收回MA10 2,296、MA20 2,362.75與半年線MA120 2,390.58（收盤恰等於2,395門檻），未收回MA60 2,417.50（差22.50元）；回補2,345–2,360、2,380–2,390兩缺口並留下2,285–2,305向上缺口；MACD柱-9.42續收斂、KD黃金交叉第4日、RSI 52.50回到50之上；外資轉買6,038張中止連23賣；轉弱程度明顯緩和）'
change_pct = '+4.81%（11/11收，相對大盤+1.20pp（精算+1.1962），恢復相對抗跌；加權+1,545.23點、+3.62%收44,257.15，全市場成交8,762.95億）；11/10為+2.24%'
nyse_price = 'US$413.35（11/11收，+4.25%，2年266個上漲日並列第26大，非極端；連4日上漲；費半+5.66%；ADR溢價同日對齊+13.43%（ADR 11/11對台股11/11與匯率32.86），11/10同日對齊+14.07%；盤後約405.30（-1.95%，19:59 EST））；11/10收US$396.48'
volume = '33,056（11/11收，-9.62%；為前報基準58,446張的0.57倍（-43.44%）＝量縮上漲、新5日均量54,281.0張；2026年206日中第114大）；11/10為36,576張'
news_summary = """11/11收2,395(+4.81%)連3漲、收回半年線但量縮；外資轉買6,038張中止連23賣；輝達傳撤換CEO盤後重挫；維持看空。"""
change_color = '00B050'
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
