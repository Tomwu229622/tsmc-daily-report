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

today_str = "2026-07-21"
tw_price = "NT$2,320（7/20收）"   # 台股2330 7/20 Mon官方收NT$2,320(+30、+1.31% vs 7/17 NT$2,290、開2,300/高2,345/低2,300、量55,790張TWSE口徑、成交額1,298.16億、成交188,463筆、TWSE確認)——崩跌後首日量縮止穩反彈、大盤續跌-0.52%下逆勢撐盤
change_pct = "+1.31%（7/20收）"    # 7/20官方漲跌幅+1.31%（TWSE、崩跌後技術性反彈、三大法人轉買超）
nyse_price = "US$402.30（7/20收）" # NYSE TSM 7/20 Mon收$402.30(+$3.93、+0.99% vs 7/17 $398.37、Yahoo API確認)——收復$400；SOX同日+0.60%止跌(自6/22高仍-19.8%、熊市邊緣)
volume = "55,790（7/20收）"    # 台股7/20官方成交量55,790張（TWSE口徑、7/17爆量97,362張的57%、反彈量縮）、成交額1,298.16億
news_summary = "報告日2026-07-21(Tue)。✅崩跌後首日止穩：台股2330 7/20官方收NT$2,320(+30、+1.31%、開2,300/高2,345/低2,300、量縮55,790張=前日57%、額1,298.16億、TWSE確認)——大盤續跌-221.57點(-0.52%)收42,449.70(盤中震盪逾1,100點、融資減碼、聯電連2跌停、日月光-2.77%)下逆勢撐盤、貢獻約244點。📊籌碼大幅改善：三大法人7/20對2330轉買超+4,002張(T86官方)——外資-1,672張(連8賣、惟自7/17 -44,184收斂96%)、投信+1,049(連3買)、自營+4,626(連8買)；全市場三大法人轉買超21.7億(BFI82U、vs 7/17賣超2,631.5億史上最大)。✅NYSE TSM 7/20收$402.30(+0.99%)收復$400；SOX +0.60%收11,743.85止跌(自6/22史高仍-19.8%、熊市邊緣)——惟設備股續跌(LRCX-2.09%、KLAC-2.42%)、AAPL-2.14%、板塊分歧。🚀法說會後外資密集調升目標價：Susquehanna $600、Barclays $650、BofA $590——S&P Global統計19位平均$520.37、共識強力買進。⚠️Yardeni警告SOXX熊市恐再跌12%、AI情緒未確認止穩；韓股補跌(SK海力士7/20 -4.2%、三星-4.3%)後7/21早盤回穩。📊ADR溢價維持+12.0%(匯率~32.3)。技術面(官方收盤計算)：收2,320站回布林下軌2,307、仍壓在60MA 2,328之下；RSI 43.9回升、MACD綠柱-19.2慣性擴大、KDJ J5.0深超賣；支撐2,300-2,307/2,290/2,200、壓力2,328(60MA)/2,388(5MA)/2,424(20MA)——止穩待確認、站回60MA為反彈續航訊號。"
change_color = "00B050"        # 綠色（7/20官方+1.31%上漲）

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
