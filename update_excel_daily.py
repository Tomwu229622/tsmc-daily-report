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

today_str = "2026-07-16"
tw_price = "NT$2,440（7/15收）"   # 台股2330 7/15 Wed官方收NT$2,440(+20、+0.83% vs 7/14 NT$2,420、開2,425/高2,460/低2,415、量33,666張TWSE口徑、成交額822.14億、均價2,442、TWSE)——大盤大漲+893.64點(+2.00%)下明顯落後屬法說會前觀望、惟站回5MA/20MA
change_pct = "+0.83%（7/15收）"    # 7/15官方漲跌幅+0.83%（TWSE、大盤+2.00%下落後、法說會前觀望）
nyse_price = "US$419.48（7/15收）" # NYSE TSM 7/15 Wed收$419.48(-$0.91、-0.22% vs 7/14 $420.39、日內410.78-428.88、Yahoo)法說會前連4日小跌；同日SOX -2.08%回檔(MU -8.02%)、TSM相對抗跌
volume = "33,666（7/15收）"    # 台股7/15官方成交量33,666張（TWSE口徑）、成交額822.14億
news_summary = "報告日2026-07-16(Thu)。⭐⭐Q2法說會今日14:00登場：Q2初步營收約NT$1.27兆(~$39.2B、季度新高)已確認達成指引$39.0-40.2B；共識估營收~$40B(+33% YoY)、每ADR EPS $3.81-3.87(+55%+)、毛利率挑戰70%(指引65.5-67.5%)、獲利連5季創高；關注上修全年財測(現行>30%)與2027-28資本支出；台股13:30收盤、效應主要今晚美股與7/17台股反映。📈台股2330 7/15收NT$2,440(+20、+0.83%、量33,666張、額822.14億)——大盤大漲+893.64點(+2.00%)收45,631.59、2330明顯落後屬觀望、惟站回5MA~2,436/20MA~2,430。📊三大法人7/15(T86)合計賣超3,568張：外資-3,427(連5賣、5日-34,792、惟較7/14 -12,416大幅收斂)、投信-596(終止連買)、自營+456。📉NYSE TSM 7/15收$419.48(-0.22%)相對抗跌；⚠️同日SOX -2.08%回檔、MU -8.02%(HBM獲利了結)、記憶體重挫延續至7/16韓股盤中(SK海力士-10%、三星-7.7%)；⭐ASML二度上修2026展望(+2.23%)、AAPL +4.01%、NVDA +0.33%。🚀法說會前分析師密集上調：美銀$490→$590(7/14)、Barclays $625(+49.0%)、Citi NT$3,800(+55.7%)、Nomura NT$3,425。📊ADR溢價收斂至+10.5%($419.48 vs理論值$379.47、匯率32.15)。⚠️風險：法說會後劇烈波動、外資連5賣、AI記憶體鏈急回吐、232關稅輸中晶片查驗、Intel 18A、估值不低(P/E~32.3x)。技術面：RSI 54.6回中軸上、MACD柱仍負-7.0、KDJ J=26自超賣回升；支撐2,415/2,390/2,342、壓力2,460/2,480/2,535史高。"
change_color = "00B050"        # 綠色（7/15官方+0.83%上漲）

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
